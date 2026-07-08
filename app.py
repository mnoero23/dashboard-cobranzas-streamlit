from __future__ import annotations

from datetime import date
from pathlib import Path
import re
import unicodedata

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries


APP_TITLE = "Dashboard de Cobranzas y Facturación para PyMEs"
DATA_PATH = Path("data") / "dataset_cobranzas.xlsx"

EXPECTED_COLUMNS = [
    "id_factura",
    "id_cliente",
    "cliente",
    "tipo_cliente",
    "responsable",
    "id_servicio",
    "tipo_servicio",
    "fecha_emision",
    "fecha_vencimiento",
    "fecha_pago",
    "moneda",
    "importe_facturado",
    "importe_cobrado",
    "saldo_pendiente",
    "tipo_cambio",
    "importe_equivalente_ars",
    "saldo_equivalente_ars",
    "dias_atraso",
    "estado",
    "prioridad_cobranza",
    "proxima_accion",
    "condicion_pago_dias",
    "perfil_pago",
]

REQUIRED_FOR_DASHBOARD = [
    "fecha_emision",
    "fecha_vencimiento",
    "cliente",
    "moneda",
    "estado",
    "importe_equivalente_ars",
    "saldo_equivalente_ars",
    "saldo_pendiente",
]

DATE_COLUMNS = ["fecha_emision", "fecha_vencimiento", "fecha_pago"]
NUMERIC_COLUMNS = [
    "importe_facturado",
    "importe_cobrado",
    "saldo_pendiente",
    "tipo_cambio",
    "importe_equivalente_ars",
    "saldo_equivalente_ars",
    "dias_atraso",
    "condicion_pago_dias",
]


st.set_page_config(
    page_title=APP_TITLE,
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    :root {
        --border: #d9e1ea;
        --muted: #5c6773;
        --panel: #ffffff;
        --soft: #f5f7fa;
    }

    .stApp {
        background: #f7f9fc;
        color: #1f2933;
    }

    [data-testid="stSidebar"] {
        background: #ffffff;
        border-right: 1px solid var(--border);
    }

    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2.5rem;
        max-width: 1320px;
    }

    .eyebrow {
        color: var(--muted);
        font-size: 0.9rem;
        margin-bottom: 0.15rem;
    }

    .notice {
        display: inline-block;
        margin-top: 0.65rem;
        padding: 0.45rem 0.75rem;
        border: 1px solid #cbd8e6;
        border-radius: 6px;
        background: #eef4fb;
        color: #334e68;
        font-size: 0.92rem;
    }

    .section-title {
        margin-top: 1.4rem;
        margin-bottom: 0.35rem;
        font-size: 1.25rem;
        font-weight: 700;
        color: #25313f;
    }

    [data-testid="stMetric"] {
        background: var(--panel);
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 0.9rem 1rem;
        box-shadow: 0 1px 2px rgba(31, 41, 51, 0.04);
    }

    [data-testid="stMetricLabel"] {
        color: #52606d;
        font-weight: 600;
    }

    [data-testid="stMetricValue"] {
        color: #1f2933;
        font-size: 1.45rem;
    }

    div[data-testid="stPlotlyChart"] {
        background: var(--panel);
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 0.65rem;
        box-shadow: 0 1px 2px rgba(31, 41, 51, 0.04);
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def normalize_name(value: object) -> str:
    text = str(value or "").strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = re.sub(r"[^a-zA-Z0-9]+", "_", text.lower())
    return re.sub(r"_+", "_", text).strip("_")


def format_ars(value: float | int | None) -> str:
    value = 0 if pd.isna(value) else float(value)
    return "$ " + f"{value:,.0f}".replace(",", ".")


def format_usd(value: float | int | None) -> str:
    value = 0 if pd.isna(value) else float(value)
    return "US$ " + f"{value:,.0f}".replace(",", ".")


def find_table_range(path: Path, table_name: str = "FacturasTable") -> tuple[str, str] | None:
    workbook = load_workbook(path, read_only=False, data_only=True)
    for sheet in workbook.worksheets:
        for table in sheet.tables.values():
            if normalize_name(table.name) == normalize_name(table_name):
                return sheet.title, table.ref
    return None


def read_table_from_range(path: Path, sheet_name: str, table_ref: str) -> pd.DataFrame:
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook[sheet_name]
    min_col, min_row, max_col, max_row = range_boundaries(table_ref)
    rows = sheet.iter_rows(
        min_row=min_row,
        max_row=max_row,
        min_col=min_col,
        max_col=max_col,
        values_only=True,
    )
    values = list(rows)
    if not values:
        return pd.DataFrame()
    headers = [str(header).strip() if header is not None else "" for header in values[0]]
    return pd.DataFrame(values[1:], columns=headers)


@st.cache_data(show_spinner=False)
def load_facturas(path: str) -> pd.DataFrame:
    excel_path = Path(path)
    xls = pd.ExcelFile(excel_path)

    facturas_sheet = next(
        (sheet for sheet in xls.sheet_names if normalize_name(sheet) == "facturas"),
        None,
    )
    if facturas_sheet:
        df = pd.read_excel(excel_path, sheet_name=facturas_sheet)
    else:
        table_info = find_table_range(excel_path)
        if table_info is None:
            raise ValueError("No se encontró una hoja 'Facturas' ni una tabla 'FacturasTable'.")
        sheet_name, table_ref = table_info
        df = read_table_from_range(excel_path, sheet_name, table_ref)

    rename_map = {}
    expected_by_norm = {normalize_name(col): col for col in EXPECTED_COLUMNS}
    for column in df.columns:
        normalized = normalize_name(column)
        if normalized in expected_by_norm:
            rename_map[column] = expected_by_norm[normalized]
    df = df.rename(columns=rename_map)

    missing_columns = [column for column in EXPECTED_COLUMNS if column not in df.columns]
    df.attrs["missing_columns"] = missing_columns

    for column in EXPECTED_COLUMNS:
        if column not in df.columns:
            df[column] = pd.NA

    for column in DATE_COLUMNS:
        df[column] = pd.to_datetime(df[column], errors="coerce")

    for column in NUMERIC_COLUMNS:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)

    for column in ["cliente", "moneda", "estado", "prioridad_cobranza", "proxima_accion"]:
        df[column] = df[column].fillna("Sin dato").astype(str).str.strip()
        df.loc[df[column].eq(""), column] = "Sin dato"

    return df


def validate_dashboard_columns(df: pd.DataFrame) -> list[str]:
    missing_columns = set(df.attrs.get("missing_columns", []))
    return [column for column in REQUIRED_FOR_DASHBOARD if column in missing_columns]


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("Filtros")

    valid_dates = df["fecha_emision"].dropna()
    if valid_dates.empty:
        st.sidebar.warning("No hay fechas de emisión válidas para filtrar por período.")
        start_date = end_date = None
    else:
        min_date = valid_dates.min().date()
        max_date = valid_dates.max().date()
        selected_period = st.sidebar.date_input(
            "Período",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            format="DD/MM/YYYY",
        )
        if isinstance(selected_period, tuple) and len(selected_period) == 2:
            start_date, end_date = selected_period
        else:
            start_date, end_date = min_date, max_date

    clients = sorted(df["cliente"].dropna().unique())
    currencies = sorted(df["moneda"].dropna().unique())
    states = sorted(df["estado"].dropna().unique())

    selected_clients = st.sidebar.multiselect("Cliente", clients, default=clients)
    selected_currencies = st.sidebar.multiselect("Moneda", currencies, default=currencies)
    selected_states = st.sidebar.multiselect("Estado", states, default=states)

    filtered = df.copy()
    if start_date and end_date:
        start_ts = pd.Timestamp(start_date)
        end_ts = pd.Timestamp(end_date)
        filtered = filtered[
            filtered["fecha_emision"].between(start_ts, end_ts, inclusive="both")
        ]
    if selected_clients:
        filtered = filtered[filtered["cliente"].isin(selected_clients)]
    if selected_currencies:
        filtered = filtered[filtered["moneda"].isin(selected_currencies)]
    if selected_states:
        filtered = filtered[filtered["estado"].isin(selected_states)]

    return filtered


def build_kpis(df: pd.DataFrame) -> dict[str, float | int]:
    today = pd.Timestamp(date.today())
    vencidas_mask = (df["fecha_vencimiento"] < today) & (df["saldo_equivalente_ars"] > 0)
    deuda_mask = df["saldo_equivalente_ars"] > 0

    total_facturado_ars = df["importe_equivalente_ars"].sum()
    saldo_pendiente_ars = df["saldo_equivalente_ars"].sum()

    return {
        "total_facturado_ars": total_facturado_ars,
        "total_cobrado_ars": total_facturado_ars - saldo_pendiente_ars,
        "saldo_pendiente_ars": saldo_pendiente_ars,
        "saldo_vencido_ars": df.loc[vencidas_mask, "saldo_equivalente_ars"].sum(),
        "facturas_vencidas": int(vencidas_mask.sum()),
        "saldo_pendiente_usd": df.loc[df["moneda"].str.upper().eq("USD"), "saldo_pendiente"].sum(),
        "cantidad_facturas": int(len(df)),
        "clientes_con_deuda": int(df.loc[deuda_mask, "cliente"].nunique()),
    }


def style_plot(fig):
    fig.update_layout(
        template="plotly_white",
        font=dict(color="#25313f"),
        margin=dict(l=20, r=20, t=55, b=25),
        title=dict(font=dict(size=18), x=0.01),
        legend_title_text="",
        paper_bgcolor="white",
        plot_bgcolor="white",
    )
    fig.update_xaxes(showgrid=True, gridcolor="#edf1f5", zeroline=False)
    fig.update_yaxes(showgrid=False, zeroline=False)
    return fig


def render_header() -> None:
    st.markdown('<div class="eyebrow">Caso de portfolio con dataset simulado</div>', unsafe_allow_html=True)
    st.title(APP_TITLE)
    st.caption("Seguimiento ejecutivo de facturación, cobranza, deuda pendiente y facturas críticas.")
    st.markdown('<div class="notice">Dataset simulado para fines de portfolio.</div>', unsafe_allow_html=True)


def render_metrics(kpis: dict[str, float | int]) -> None:
    st.markdown('<div class="section-title">KPIs principales</div>', unsafe_allow_html=True)
    first_row = st.columns(3)
    first_row[0].metric("Total facturado", format_ars(kpis["total_facturado_ars"]))
    first_row[1].metric("Total cobrado", format_ars(kpis["total_cobrado_ars"]))
    first_row[2].metric("Saldo pendiente", format_ars(kpis["saldo_pendiente_ars"]))

    second_row = st.columns(3)
    second_row[0].metric("Saldo vencido", format_ars(kpis["saldo_vencido_ars"]))
    second_row[1].metric("Facturas vencidas", f"{kpis['facturas_vencidas']:,.0f}".replace(",", "."))
    second_row[2].metric("Saldo pendiente USD", format_usd(kpis["saldo_pendiente_usd"]))


def render_charts(df: pd.DataFrame) -> None:
    df = df.copy()
    df["mes_emision"] = df["fecha_emision"].dt.to_period("M").dt.to_timestamp()

    monthly = (
        df.dropna(subset=["mes_emision"])
        .groupby("mes_emision", as_index=False)
        .agg(
            **{
                "Total facturado ARS Eq": ("importe_equivalente_ars", "sum"),
                "Saldo pendiente ARS Eq": ("saldo_equivalente_ars", "sum"),
            }
        )
    )
    if not monthly.empty:
        monthly["Total cobrado ARS Eq"] = (
            monthly["Total facturado ARS Eq"] - monthly["Saldo pendiente ARS Eq"]
        )
        monthly_long = monthly.melt(
            id_vars="mes_emision",
            value_vars=["Total facturado ARS Eq", "Total cobrado ARS Eq"],
            var_name="Indicador",
            value_name="Importe",
        )
        fig_monthly = px.line(
            monthly_long,
            x="mes_emision",
            y="Importe",
            color="Indicador",
            markers=True,
            title="Evolución mensual de facturación y cobranza",
            color_discrete_sequence=["#2563eb", "#16a34a"],
        )
        fig_monthly.update_yaxes(tickprefix="$ ", tickformat=",.0f")
        fig_monthly.update_xaxes(title="", tickformat="%m/%Y")
        st.plotly_chart(style_plot(fig_monthly), use_container_width=True)
    else:
        st.info("No hay datos suficientes para mostrar la evolución mensual.")

    left, right = st.columns(2)

    debt_by_client = (
        df.groupby("cliente", as_index=False)["saldo_equivalente_ars"]
        .sum()
        .sort_values("saldo_equivalente_ars", ascending=False)
        .head(10)
        .sort_values("saldo_equivalente_ars", ascending=True)
    )
    if not debt_by_client.empty:
        fig_clients = px.bar(
            debt_by_client,
            y="cliente",
            x="saldo_equivalente_ars",
            orientation="h",
            title="Clientes con mayor deuda pendiente",
            labels={"cliente": "", "saldo_equivalente_ars": "Saldo pendiente ARS Eq"},
            color_discrete_sequence=["#2f6f73"],
        )
        fig_clients.update_xaxes(tickprefix="$ ", tickformat=",.0f")
        left.plotly_chart(style_plot(fig_clients), use_container_width=True)
    else:
        left.info("No hay deuda pendiente por cliente para mostrar.")

    status_counts = (
        df.groupby("estado", as_index=False)
        .size()
        .rename(columns={"size": "cantidad_facturas"})
        .sort_values("cantidad_facturas", ascending=True)
    )
    if not status_counts.empty:
        fig_status = px.bar(
            status_counts,
            y="estado",
            x="cantidad_facturas",
            orientation="h",
            title="Estado general de cobranzas",
            labels={"estado": "", "cantidad_facturas": "Cantidad de facturas"},
            color_discrete_sequence=["#64748b"],
        )
        right.plotly_chart(style_plot(fig_status), use_container_width=True)
    else:
        right.info("No hay estados para mostrar.")

    currency_debt = (
        df.groupby("moneda", as_index=False)["saldo_equivalente_ars"]
        .sum()
        .sort_values("saldo_equivalente_ars", ascending=True)
    )
    if not currency_debt.empty:
        fig_currency = px.bar(
            currency_debt,
            y="moneda",
            x="saldo_equivalente_ars",
            orientation="h",
            title="Exposición de deuda por moneda",
            labels={"moneda": "", "saldo_equivalente_ars": "Saldo pendiente ARS Eq"},
            color_discrete_sequence=["#9a6b36"],
        )
        fig_currency.update_xaxes(tickprefix="$ ", tickformat=",.0f")
        st.plotly_chart(style_plot(fig_currency), use_container_width=True)
    else:
        st.info("No hay deuda por moneda para mostrar.")


def render_critical_invoices(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-title">Facturas críticas para gestionar</div>', unsafe_allow_html=True)

    critical = df[
        (df["saldo_equivalente_ars"] > 0)
        & (~df["estado"].str.casefold().eq("cobrada"))
    ].copy()

    if critical.empty:
        st.success("No hay facturas críticas para gestionar con los filtros seleccionados.")
        return

    columns = [
        "cliente",
        "id_factura",
        "fecha_vencimiento",
        "moneda",
        "saldo_pendiente",
        "saldo_equivalente_ars",
        "dias_atraso",
        "prioridad_cobranza",
        "proxima_accion",
        "estado",
    ]
    critical = critical[columns].sort_values("dias_atraso", ascending=False)

    display_df = critical.copy()
    display_df["fecha_vencimiento"] = display_df["fecha_vencimiento"].dt.strftime("%d/%m/%Y")
    display_df["saldo_pendiente"] = display_df.apply(
        lambda row: format_usd(row["saldo_pendiente"])
        if str(row["moneda"]).upper() == "USD"
        else format_ars(row["saldo_pendiente"]),
        axis=1,
    )
    display_df["saldo_equivalente_ars"] = display_df["saldo_equivalente_ars"].map(format_ars)
    display_df["dias_atraso"] = display_df["dias_atraso"].fillna(0).astype(int)

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "cliente": "Cliente",
            "id_factura": "Factura",
            "fecha_vencimiento": "Vencimiento",
            "moneda": "Moneda",
            "saldo_pendiente": "Saldo pendiente",
            "saldo_equivalente_ars": "Saldo ARS Eq",
            "dias_atraso": "Días atraso",
            "prioridad_cobranza": "Prioridad",
            "proxima_accion": "Próxima acción",
            "estado": "Estado",
        },
    )


def main() -> None:
    render_header()

    if not DATA_PATH.exists():
        st.error(
            "No se encontró el archivo de datos. Ubicá el Excel en "
            "`data/dataset_cobranzas.xlsx` y volvé a ejecutar la app."
        )
        st.stop()

    try:
        df = load_facturas(str(DATA_PATH))
    except Exception as exc:
        st.error("No se pudo leer el archivo Excel.")
        st.exception(exc)
        st.stop()

    missing_required = validate_dashboard_columns(df)
    if missing_required:
        st.error(
            "Faltan columnas necesarias para construir el dashboard: "
            + ", ".join(f"`{column}`" for column in missing_required)
        )
        st.stop()

    missing_optional = [
        column for column in EXPECTED_COLUMNS if column not in df.columns or df[column].isna().all()
    ]
    if missing_optional:
        with st.expander("Columnas opcionales no encontradas o vacías"):
            st.write(", ".join(missing_optional))

    filtered = apply_filters(df)
    if filtered.empty:
        st.info("No hay datos disponibles para los filtros seleccionados.")
        st.stop()

    kpis = build_kpis(filtered)
    render_metrics(kpis)
    render_charts(filtered)
    render_critical_invoices(filtered)

    st.markdown('<div class="section-title">Sobre este proyecto</div>', unsafe_allow_html=True)
    st.write(
        "Este dashboard fue desarrollado como caso de portfolio para mostrar cómo una PyME "
        "puede transformar una base de facturación y cobranzas en una herramienta ejecutiva "
        "de seguimiento. Permite identificar deuda pendiente, facturas vencidas, clientes "
        "críticos, exposición por moneda y acciones prioritarias de cobranza."
    )


if __name__ == "__main__":
    main()
