# Dashboard de Cobranzas y Facturación para PyMEs

App publicada: https://dashboard-cobranzas.streamlit.app/

Aplicación web interactiva desarrollada con Python, Streamlit, pandas y Plotly para analizar facturación, cobranzas, deuda pendiente y facturas críticas de una PyME.

> Caso de portfolio con dataset simulado. No utiliza datos reales ni referencias a empresas reales.

## Objetivo

Transformar una base de facturación y cobranzas en una herramienta ejecutiva, clara y fácil de compartir, apta para publicar en Streamlit Community Cloud o incluir en un portfolio freelance de Data / IA.

## Problema de negocio

Muchas PyMEs gestionan facturas, pagos y deuda pendiente en planillas dispersas. Esto dificulta responder preguntas clave:

- ¿Cuánto se facturó y cuánto se cobró?
- ¿Qué saldo sigue pendiente?
- ¿Qué deuda está vencida?
- ¿Qué clientes concentran la mayor deuda?
- ¿Qué facturas requieren acción prioritaria?
- ¿Cuál es la exposición por moneda?

## Solución desarrollada

La app lee un archivo Excel simulado, normaliza columnas, valida datos y presenta un tablero ejecutivo con filtros, KPIs dinámicos, visualizaciones interactivas y una tabla de gestión de facturas críticas.

La segunda iteración incorpora:

- Filtros interactivos por período, cliente, moneda y estado.
- Botones de selección para moneda y estado cuando la versión de Streamlit lo permite.
- Selección en gráficos para filtrar el dashboard por cliente, estado, moneda o mes.
- KPIs dinámicos recalculados con los filtros manuales y los filtros generados por gráficos.
- Navegación por pestañas: Resumen, Clientes, Facturas críticas y Proyecto.
- Pestaña Clientes con ranking de deuda, concentración, vencimientos y detalle descargable por cliente.
- Tabla de facturas críticas recalculada con todos los filtros.
- Descarga de la tabla filtrada como CSV.
- Diseño visual más sobrio, profesional y orientado a dashboard ejecutivo.

## Herramientas usadas

- Python
- Streamlit
- pandas
- Plotly Express
- openpyxl

## Estructura del proyecto

```text
dashboard-cobranzas-streamlit/
├── .github/
│   └── workflows/
│       └── update-dataset.yml
├── app.py
├── requirements.txt
├── README.md
├── scripts/
│   └── update_dataset.py
└── data/
    ├── dataset_cobranzas.xlsx
    └── update_log.csv
```

## Dataset simulado

El archivo Excel debe ubicarse en:

```text
data/dataset_cobranzas.xlsx
```

La app intenta leer una hoja llamada `Facturas`. Si no existe, busca una tabla llamada `FacturasTable`.

## Actualización automática del dataset

Este proyecto incluye una rutina automática que genera facturas ficticias y actualiza la base de datos usada por la app. Esto permite que el dashboard se mantenga dinámico sin usar datos reales, APIs externas ni información de empresas reales.

El script principal está en:

```text
scripts/update_dataset.py
```

La rutina:

- Lee `data/dataset_cobranzas.xlsx`.
- Usa la hoja `Facturas` y la tabla `FacturasTable` como fuente principal.
- Reutiliza clientes y servicios existentes del Excel.
- Agrega entre 3 y 8 facturas simuladas por ejecución.
- Calcula fechas, saldos, estado, mora, prioridad y próxima acción.
- Mantiene solo los últimos 12 meses móviles según `fecha_emision`.
- Registra cada ejecución en `data/update_log.csv`.

Para ejecutarla manualmente:

```bash
python scripts/update_dataset.py
```

La app muestra la fecha de última actualización, la cantidad total de registros simulados y la aclaración de que el dataset es ficticio para portfolio.

### GitHub Actions

El workflow `.github/workflows/update-dataset.yml` ejecuta la actualización una vez por día y también permite correrla manualmente desde la pestaña **Actions** de GitHub mediante `workflow_dispatch`.

Cuando el Excel o el log cambian, el workflow hace commit y push automático al repositorio para que Streamlit Community Cloud vuelva a desplegar la app con datos simulados actualizados.

## KPIs incluidos

- Total facturado ARS equivalente
- Total cobrado ARS equivalente
- Saldo pendiente ARS equivalente
- Saldo vencido ARS equivalente
- Facturas vencidas
- Saldo pendiente USD

## Visualizaciones incluidas

- Evolución mensual de facturación y cobranza.
- Clientes con mayor deuda pendiente.
- Estado general de cobranzas.
- Exposición de deuda por moneda.
- Vista de clientes con ranking, métricas de deuda vencida y perfil individual.
- Tabla de facturas críticas para gestionar.

## Interactividad

Los gráficos usan selección interactiva de Plotly en Streamlit. Al seleccionar elementos, el dashboard puede aplicar filtros cruzados:

- Cliente seleccionado desde el gráfico de deuda por cliente.
- Estado seleccionado desde el gráfico de estado general.
- Moneda seleccionada desde el gráfico de exposición por moneda.
- Mes seleccionado desde la evolución mensual.

La app muestra los filtros activos por interacción y permite limpiar solo las selecciones de gráficos o todos los filtros.

## Cómo ejecutar localmente

1. Crear y activar un entorno virtual, opcional pero recomendado.

```bash
python -m venv .venv
```

En Windows:

```bash
.venv\Scripts\activate
```

En macOS/Linux:

```bash
source .venv/bin/activate
```

2. Instalar dependencias.

```bash
pip install -r requirements.txt
```

3. Ejecutar la app.

```bash
streamlit run app.py
```

## Publicación en Streamlit Community Cloud

Para publicar la app:

1. Subir este proyecto a GitHub.
2. Verificar que `app.py`, `requirements.txt` y la carpeta `data/` estén incluidos.
3. Crear una nueva app en Streamlit Community Cloud.
4. Seleccionar el repositorio y apuntar a `app.py`.

## Si aparece un error

- Revisar que el Excel exista en `data/dataset_cobranzas.xlsx`.
- Confirmar que el archivo tenga una hoja `Facturas` o una tabla `FacturasTable`.
- Revisar que las columnas principales estén presentes, aunque tengan diferencias de mayúsculas/minúsculas.
- Verificar que las fechas e importes tengan formatos convertibles.
- Si un filtro deja el tablero vacío, la app mostrará: `No hay datos disponibles para los filtros seleccionados.`
