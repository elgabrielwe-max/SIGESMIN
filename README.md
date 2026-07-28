# Dashboard web de Capacitaciones Anexo 6

Dashboard interactivo creado a partir del archivo **Capacitaciones Anexo 6 Dashboard.xlsx**.

## Funciones

- Filtros por área, mes de corte y guardia.
- Temas faltantes y en proceso identificados automáticamente.
- Cumplimiento por tema y por guardias A, B y C.
- Evolución mensual y comparativo entre áreas.
- Carga de una nueva versión del Excel desde el navegador.
- Exportación del reporte filtrado a Excel.
- Diseño adaptable a celular y computadora.

## Probar en la computadora

1. Instala Node.js 20 o superior.
2. Abre esta carpeta en VS Code.
3. Ejecuta:

```bash
npm install
npm run dev
```

4. Abre la dirección que muestra la terminal, normalmente `http://localhost:5173`.

## Publicar gratis con GitHub Pages

1. Crea un repositorio nuevo en GitHub.
2. Sube todos los archivos de esta carpeta a la rama `main`.
3. En GitHub abre **Settings → Pages**.
4. En **Build and deployment**, selecciona **GitHub Actions**.
5. Abre la pestaña **Actions** y espera a que termine `Deploy GitHub Pages`.
6. GitHub mostrará la dirección pública de la web.

## Actualizar datos

La página trae los datos actuales dentro de `public/data.json`. Para trabajar con una versión nueva no necesitas modificar código: pulsa **Cargar Excel** y selecciona el archivo. Debe conservar la hoja llamada `CUMPLIMIENTO GUARDIAS` y estas columnas:

- Área Responsable
- Mes
- N° Mes
- Tema
- Guardia A
- Guardia B
- Guardia C
- Estado Programa

Los valores reconocidos son `C`, `F` y `EP`; también convierte automáticamente `P` en `EP`.

> La carga del Excel ocurre únicamente en el navegador. El archivo no se envía a ningún servidor.
