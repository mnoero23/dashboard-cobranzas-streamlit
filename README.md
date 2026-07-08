[README.md](https://github.com/user-attachments/files/29777776/README.md)
# Dashboard de Cobranzas y Facturación para PyMEs

Aplicación web interactiva desarrollada con Python y Streamlit para analizar facturación, cobranzas, deuda pendiente y facturas críticas de una PyME.

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

La app lee un archivo Excel simulado, normaliza columnas, valida datos y presenta un tablero interactivo con filtros, KPIs, visualizaciones y una tabla de gestión de facturas críticas.

## Herramientas usadas

- Python
- Streamlit
- pandas
- Plotly Express / Plotly Graph Objects
- openpyxl

## Estructura del proyecto

```text
dashboard-cobranzas-streamlit/
├── app.py
├── requirements.txt
├── README.md
└── data/
    └── dataset_cobranzas.xlsx
```

## Dataset simulado

El archivo Excel debe ubicarse en:

```text
data/dataset_cobranzas.xlsx
```

La app intenta leer una hoja llamada `Facturas`. Si no existe, busca una tabla llamada `FacturasTable`.

## KPIs incluidos

- Total facturado ARS equivalente
- Total cobrado ARS equivalente
- Saldo pendiente ARS equivalente
- Saldo vencido ARS equivalente
- Facturas vencidas
- Saldo pendiente USD
- Cantidad de facturas
- Clientes con deuda

## Visualizaciones incluidas

- Evolución mensual de facturación y cobranza
- Clientes con mayor deuda pendiente
- Estado general de cobranzas
- Exposición de deuda por moneda
- Tabla de facturas críticas para gestionar

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
