import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import {
  AreaChart, Area, BarChart, Bar, CartesianGrid, Cell, Legend,
  PieChart, Pie, ResponsiveContainer, Tooltip, XAxis, YAxis
} from 'recharts';
import {
  AlertTriangle, Building2, CalendarDays, CheckCircle2, Clock3,
  Download, FileSpreadsheet, Filter, RefreshCw, Search, ShieldCheck,
  TrendingUp, Upload, UsersRound, XCircle
} from 'lucide-react';
import './styles.css';

const MONTHS = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SETIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
const STATUS = { C: 'Cumplida', F: 'Faltante', EP: 'En proceso' };
const STATUS_COLOR = { C: '#16a34a', F: '#dc2626', EP: '#f59e0b' };

const cleanStatus = value => {
  const v = String(value ?? '').trim().toUpperCase();
  if (v === 'P') return 'EP';
  return ['C','F','EP'].includes(v) ? v : '';
};

function normalizeRows(rows) {
  return rows.map(r => {
    const a = cleanStatus(r['Guardia A']);
    const b = cleanStatus(r['Guardia B']);
    const c = cleanStatus(r['Guardia C']);
    const guards = { A: a, B: b, C: c };
    const completed = Object.values(guards).filter(x => x === 'C').length;
    let estado = cleanStatus(r['Estado Programa']);
    if (!estado) {
      if (completed === 3) estado = 'C';
      else if (Object.values(guards).some(x => x === 'EP')) estado = 'EP';
      else estado = 'F';
    }
    return {
      area: String(r['Área Responsable'] ?? '').trim(),
      mes: String(r['Mes'] ?? '').trim().toUpperCase(),
      mesNumero: Number(r['N° Mes']) || MONTHS.indexOf(String(r['Mes'] ?? '').trim().toUpperCase()) + 1,
      tema: String(r['Tema'] ?? '').trim(),
      guardias: guards,
      cumplidas: completed,
      porcentaje: Math.round((completed / 3) * 100),
      estado
    };
  }).filter(x => x.area && x.tema && x.mesNumero > 0);
}

const pct = (a,b) => b ? Math.round((a/b)*100) : 0;

function Kpi({ icon: Icon, label, value, note, tone='blue' }) {
  return <article className={`kpi kpi-${tone}`}>
    <div className="kpi-icon"><Icon size={22}/></div>
    <div><span>{label}</span><strong>{value}</strong><small>{note}</small></div>
  </article>;
}

function StatusPill({ value }) {
  const v = cleanStatus(value) || 'F';
  return <span className={`pill pill-${v.toLowerCase()}`}>{v} · {STATUS[v]}</span>;
}

function App() {
  const [rows, setRows] = useState([]);
  const [area, setArea] = useState('');
  const [month, setMonth] = useState(7);
  const [guard, setGuard] = useState('TODAS');
  const [query, setQuery] = useState('');
  const [updatedAt, setUpdatedAt] = useState('Datos incluidos en el proyecto');
  const fileInput = useRef(null);

  useEffect(() => {
    fetch('./data.json').then(r => r.json()).then(data => {
      setRows(data);
      setArea([...new Set(data.map(x => x.area))][0] || '');
    });
  }, []);

  const areas = useMemo(() => [...new Set(rows.map(x => x.area))].sort(), [rows]);
  const programmed = useMemo(() => rows.filter(x => x.area === area && x.mesNumero <= month), [rows, area, month]);
  const visible = useMemo(() => programmed.filter(x => {
    const matchGuard = guard === 'TODAS' || !!x.guardias[guard];
    const text = `${x.tema} ${x.mes} ${x.estado}`.toLowerCase();
    return matchGuard && text.includes(query.toLowerCase());
  }), [programmed, guard, query]);

  const metrics = useMemo(() => {
    const total = programmed.length;
    const fulfilled = programmed.filter(x => x.estado === 'C').length;
    const missing = programmed.filter(x => x.estado === 'F').length;
    const progress = programmed.filter(x => x.estado === 'EP').length;
    const guardPct = ['A','B','C'].map(g => ({
      guardia: `Guardia ${g}`,
      C: programmed.filter(x => x.guardias[g] === 'C').length,
      F: programmed.filter(x => x.guardias[g] === 'F' || !x.guardias[g]).length,
      EP: programmed.filter(x => x.guardias[g] === 'EP').length,
      porcentaje: pct(programmed.filter(x => x.guardias[g] === 'C').length, total)
    }));
    return { total, fulfilled, missing, progress, compliance: pct(fulfilled,total), guardPct };
  }, [programmed]);

  const statusChart = [
    {name:'Cumplidas', value:metrics.fulfilled, color:STATUS_COLOR.C},
    {name:'Faltantes', value:metrics.missing, color:STATUS_COLOR.F},
    {name:'En proceso', value:metrics.progress, color:STATUS_COLOR.EP}
  ].filter(x => x.value > 0);

  const monthly = useMemo(() => MONTHS.slice(0, month).map((m, i) => {
    const set = rows.filter(x => x.area === area && x.mesNumero === i + 1);
    return { mes: m.slice(0,3), cumplimiento: pct(set.filter(x => x.estado === 'C').length,set.length), temas:set.length };
  }), [rows, area, month]);

  const allAreas = useMemo(() => areas.map(a => {
    const set = rows.filter(x => x.area === a && x.mesNumero <= month);
    return { area: a, cumplimiento: pct(set.filter(x => x.estado === 'C').length,set.length) };
  }), [rows, areas, month]);

  async function uploadExcel(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array' });
    const sheet = wb.Sheets['CUMPLIMIENTO GUARDIAS'];
    if (!sheet) return alert('No se encontró la hoja "CUMPLIMIENTO GUARDIAS".');
    const data = normalizeRows(XLSX.utils.sheet_to_json(sheet, { defval: '' }));
    if (!data.length) return alert('La hoja no contiene registros válidos.');
    setRows(data);
    setArea(data[0].area);
    setUpdatedAt(`${file.name} · ${new Date().toLocaleString('es-PE')}`);
    e.target.value = '';
  }

  function exportCSV() {
    const exportRows = visible.map(x => ({
      Área:x.area, Mes:x.mes, Tema:x.tema, Estado:x.estado,
      'Guardia A':x.guardias.A, 'Guardia B':x.guardias.B, 'Guardia C':x.guardias.C,
      'Cumplimiento':`${x.porcentaje}%`
    }));
    const sheet = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, sheet, 'Reporte');
    XLSX.writeFile(wb, `reporte-${area.toLowerCase().replaceAll(' ','-')}-${MONTHS[month-1].toLowerCase()}.xlsx`);
  }

  return <div className="app-shell">
    <aside className="sidebar">
      <div className="brand"><div className="brand-mark"><ShieldCheck/></div><div><b>ANEXO 6</b><span>Control de capacitaciones</span></div></div>
      <nav>
        <a className="active"><TrendingUp/> Dashboard</a>
        <a href="#pendientes"><AlertTriangle/> Pendientes</a>
        <a href="#guardias"><UsersRound/> Guardias</a>
        <a href="#comparativo"><Building2/> Áreas</a>
      </nav>
      <div className="sidebar-note"><FileSpreadsheet/><div><b>Fuente actual</b><span>{updatedAt}</span></div></div>
    </aside>

    <main>
      <header className="topbar">
        <div><span className="eyebrow">SEGURIDAD Y SALUD OCUPACIONAL</span><h1>Dashboard de Capacitaciones</h1><p>Seguimiento por área, mes, tema y guardia.</p></div>
        <div className="top-actions">
          <button className="btn secondary" onClick={() => fileInput.current?.click()}><Upload size={18}/> Cargar Excel</button>
          <button className="btn primary" onClick={exportCSV}><Download size={18}/> Exportar reporte</button>
          <input ref={fileInput} hidden type="file" accept=".xlsx,.xls" onChange={uploadExcel}/>
        </div>
      </header>

      <section className="filters card">
        <div className="filter-title"><Filter size={19}/><b>Filtros del reporte</b></div>
        <label>Área<select value={area} onChange={e=>setArea(e.target.value)}>{areas.map(x=><option key={x}>{x}</option>)}</select></label>
        <label>Mes de corte<select value={month} onChange={e=>setMonth(Number(e.target.value))}>{MONTHS.map((m,i)=><option value={i+1} key={m}>{m}</option>)}</select></label>
        <label>Guardia<select value={guard} onChange={e=>setGuard(e.target.value)}><option value="TODAS">Todas</option><option value="A">Guardia A</option><option value="B">Guardia B</option><option value="C">Guardia C</option></select></label>
        <button className="icon-button" title="Restablecer" onClick={()=>{setMonth(7);setGuard('TODAS');setQuery('')}}><RefreshCw size={18}/></button>
      </section>

      <section className="kpi-grid">
        <Kpi icon={CalendarDays} label="Temas programados" value={metrics.total} note={`Enero a ${MONTHS[month-1].toLowerCase()}`} tone="blue"/>
        <Kpi icon={CheckCircle2} label="Cumplidas" value={metrics.fulfilled} note={`${metrics.compliance}% al corte`} tone="green"/>
        <Kpi icon={XCircle} label="Faltantes" value={metrics.missing} note="Requieren atención" tone="red"/>
        <Kpi icon={Clock3} label="En proceso" value={metrics.progress} note="Seguimiento activo" tone="amber"/>
      </section>

      <section className="dashboard-grid">
        <article className="card chart-card"><div className="card-heading"><div><h2>Estado al mes de corte</h2><p>{area}</p></div><strong>{metrics.compliance}%</strong></div>
          <ResponsiveContainer width="100%" height={260}><PieChart><Pie data={statusChart} dataKey="value" nameKey="name" innerRadius={65} outerRadius={95} paddingAngle={4}>{statusChart.map(x=><Cell key={x.name} fill={x.color}/>)}</Pie><Tooltip/><Legend/></PieChart></ResponsiveContainer>
        </article>
        <article className="card chart-card"><div className="card-heading"><div><h2>Evolución mensual</h2><p>Porcentaje de temas cumplidos</p></div></div>
          <ResponsiveContainer width="100%" height={260}><AreaChart data={monthly}><defs><linearGradient id="fillTrend" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#0284c7" stopOpacity={.35}/><stop offset="95%" stopColor="#0284c7" stopOpacity={0}/></linearGradient></defs><CartesianGrid strokeDasharray="3 3" vertical={false}/><XAxis dataKey="mes"/><YAxis domain={[0,100]} tickFormatter={v=>`${v}%`}/><Tooltip formatter={v=>`${v}%`}/><Area type="monotone" dataKey="cumplimiento" stroke="#0284c7" strokeWidth={3} fill="url(#fillTrend)"/></AreaChart></ResponsiveContainer>
        </article>
      </section>

      <section id="pendientes" className="card table-card">
        <div className="card-heading"><div><h2>Temas faltantes o en proceso</h2><p>Detalle automático hasta {MONTHS[month-1].toLowerCase()}</p></div><div className="search"><Search size={17}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Buscar tema..."/></div></div>
        <div className="table-wrap"><table><thead><tr><th>Mes</th><th>Tema</th><th>Estado</th><th>Avance</th><th>Guardia A</th><th>Guardia B</th><th>Guardia C</th></tr></thead><tbody>
          {visible.filter(x=>x.estado!=='C').map((x,i)=><tr key={`${x.tema}-${i}`}><td>{x.mes}</td><td className="topic">{x.tema}</td><td><StatusPill value={x.estado}/></td><td><div className="progress-cell"><div><span style={{width:`${x.porcentaje}%`}}/></div><b>{x.porcentaje}%</b></div></td>{['A','B','C'].map(g=><td key={g}><StatusPill value={x.guardias[g]||'F'}/></td>)}</tr>)}
          {!visible.some(x=>x.estado!=='C') && <tr><td colSpan="7" className="empty">No hay temas pendientes con los filtros seleccionados.</td></tr>}
        </tbody></table></div>
      </section>

      <section id="guardias" className="dashboard-grid">
        <article className="card table-card"><div className="card-heading"><div><h2>Cumplimiento por guardia</h2><p>Resumen al corte seleccionado</p></div></div><div className="table-wrap"><table><thead><tr><th>Guardia</th><th>C</th><th>F</th><th>EP</th><th>% cumplimiento</th></tr></thead><tbody>{metrics.guardPct.map(g=><tr key={g.guardia}><td><b>{g.guardia}</b></td><td>{g.C}</td><td>{g.F}</td><td>{g.EP}</td><td><div className="progress-cell"><div><span style={{width:`${g.porcentaje}%`}}/></div><b>{g.porcentaje}%</b></div></td></tr>)}</tbody></table></div></article>
        <article className="card chart-card"><div className="card-heading"><div><h2>Comparación de guardias</h2><p>Temas cumplidos, faltantes y en proceso</p></div></div><ResponsiveContainer width="100%" height={280}><BarChart data={metrics.guardPct}><CartesianGrid strokeDasharray="3 3" vertical={false}/><XAxis dataKey="guardia"/><YAxis/><Tooltip/><Legend/><Bar dataKey="C" name="Cumplidas" fill={STATUS_COLOR.C} radius={[4,4,0,0]}/><Bar dataKey="F" name="Faltantes" fill={STATUS_COLOR.F} radius={[4,4,0,0]}/><Bar dataKey="EP" name="En proceso" fill={STATUS_COLOR.EP} radius={[4,4,0,0]}/></BarChart></ResponsiveContainer></article>
      </section>

      <section className="card table-card"><div className="card-heading"><div><h2>Cumplimiento por tema y guardia</h2><p>Todos los temas programados hasta el mes de corte</p></div></div><div className="table-wrap"><table><thead><tr><th>Mes</th><th>Tema</th><th>A</th><th>B</th><th>C</th><th>%</th><th>Estado</th></tr></thead><tbody>{visible.map((x,i)=><tr key={`${x.tema}-all-${i}`}><td>{x.mes}</td><td className="topic">{x.tema}</td>{['A','B','C'].map(g=><td key={g}><StatusPill value={x.guardias[g]||'F'}/></td>)}<td><b>{x.porcentaje}%</b></td><td><StatusPill value={x.estado}/></td></tr>)}</tbody></table></div></section>

      <section id="comparativo" className="card chart-card"><div className="card-heading"><div><h2>Cumplimiento por área</h2><p>Comparativo general hasta {MONTHS[month-1].toLowerCase()}</p></div></div><ResponsiveContainer width="100%" height={Math.max(340,areas.length*30)}><BarChart data={allAreas} layout="vertical" margin={{left:30,right:35}}><CartesianGrid strokeDasharray="3 3" horizontal={false}/><XAxis type="number" domain={[0,100]} tickFormatter={v=>`${v}%`}/><YAxis type="category" dataKey="area" width={150} tick={{fontSize:11}}/><Tooltip formatter={v=>`${v}%`}/><Bar dataKey="cumplimiento" fill="#2563eb" radius={[0,6,6,0]}/></BarChart></ResponsiveContainer></section>

      <footer>Dashboard Anexo 6 · Los datos permanecen en tu navegador y no se envían a servidores externos.</footer>
    </main>
  </div>;
}

createRoot(document.getElementById('root')).render(<App/>);
