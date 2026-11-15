# Sistema de Conciliación Bancaria Automática

Aplicación web interactiva diseñada para automatizar el cruce contable entre extractos bancarios y libros auxiliares en formatos Excel (`.xlsx`, `.xls`) y `.csv`. Ejecuta el procesamiento íntegramente en el navegador del usuario garantizando privacidad y rapidez, e incorpora un panel visual de indicadores y exportación de reportes auditables.

- **Desarrollado por:** Alexander Olmedo — [Portafolio Profesional](https://alexanderolmedo.vercel.app/)
- **Sitio Comercial:** [Sistemas-Majovit](https://sistemas-majovit.vercel.app/)

---

## Características Principales

- **Procesamiento Local y Seguro:** Lectura y análisis de archivos directamente en el navegador mediante **SheetJS**, sin enviar información financiera sensible a servidores externos.
- **Motor de Cruce Contable en 5 Niveles:** Algoritmo de conciliación uno a uno (evitando duplicidades) con cálculo exacto en centavos para prevenir errores de punto flotante:
  1. **Fecha y valor exactos:** Coincidencia exacta en fecha y monto.
  2. **Fecha cercana:** Valor exacto con margen de hasta $\pm 3$ días.
  3. **Validación por tercero:** Valor exacto, compatibilidad semántica de beneficiario/descripción y margen de hasta $\pm 15$ días.
  4. **Cruce por suma agrupada:** Búsqueda de subconjuntos del mismo signo que igualen exactamente el movimiento contraparte dentro del mismo periodo o ventana cercana.
  5. **Detección de partidas por revisar:** Identificación de registros con misma fecha y tercero compatible pero diferencia de monto de hasta $\$100.000$ COP para revisión manual.
- **Dashboard Analítico Interactivo:** Gráficas dinámicas con **Chart.js** que muestran la distribución de estados, efectividad por regla de cruce y concentración de partidas pendientes por origen.
- **Filtros y Búsqueda en Tiempo Real:** Filtrado por rango de fechas (`Desde` / `Hasta`), pestañas clasificadas (*Conciliados*, *Por revisar*, *Pendientes*) y buscador instantáneo.
- **Exportación Profesional:** Generación de informe consolidado descargable en formato **CSV/Excel** (con codificación UTF-8 BOM) e impresión limpia en **PDF** mediante hoja de estilos `@media print`.

---

## Formato de Archivos Soportados

### 1. Extracto Bancario (`.xlsx`, `.xls`, `.csv`)
Estructura sin encabezados por columnas:
- **Columna D:** Fecha del movimiento.
- **Columna F:** Valor con signo (`+` ingresos / `-` egresos).
- **Columna H:** Descripción o concepto bancario.

### 2. Libro Auxiliar Contable (`.xlsx`, `.xls`, `.csv`)
Estructura con fila de encabezados detectada automáticamente:
- `FECHA`
- `DETALLE` (o `CONCEPTO`)
- `NOMBRE BENEFICIARIO` (o `BENEFICIARIO` / `TERCERO`)
- `DÉBITO` y `CRÉDITO`

*(La aplicación incluye un botón **"Cargar ejemplo"** con datos de prueba precargados para evaluar el funcionamiento sin necesidad de subir archivos).*

---

## Stack Tecnológico

- **Frontend:** HTML5 Semántico, CSS3 (Grid, Flexbox, Variables y estilos de impresión), JavaScript Vanilla (ES6+)
- **Procesamiento de Hojas de Cálculo:** [SheetJS (xlsx)](https://sheetjs.com/)
- **Visualización de Datos:** [Chart.js](https://www.chartjs.org/)

---
