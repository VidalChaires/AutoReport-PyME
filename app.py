"""
app.py
======
Dashboard Comercial y de Calidad de Datos para PyMEs en Streamlit.

Funcionalidades:
- Carga de archivo demo o ingesta personalizada (.xlsx, .xls, .csv).
- Pestaña 1: Auditoría Antes/Después (Salud de datos, depuración, anomalías corregidas).
- Pestaña 2: Dashboard Ejecutivo Plotly (Ventas por sucursal, categoría, temporal, vendedores).
- Pestaña 3: Descarga de Entregables Directivos (Excel profesional estilizado y Reporte PDF).
"""

from __future__ import annotations

import io
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from cleaning_engine import clean_sales_data
from report_builder import export_to_excel_bytes, export_to_pdf_bytes
from sample_generator import generate_dirty_sales_data

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA Y ESTILOS
# ==============================================================================
st.set_page_config(
    page_title="AutoReport PyME | Data Quality & Sales Intelligence",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Inyección de estilos CSS corporativos
st.markdown(
    """
    <style>
    .main-title {
        font-size: 2.1rem;
        font-weight: 700;
        color: #1B365D;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 1.05rem;
        color: #64748B;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .audit-badge-success {
        background-color: #DCFCE7;
        color: #166534;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.9rem;
    }
    .audit-badge-warning {
        background-color: #FEF3C7;
        color: #92400E;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.9rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==============================================================================
# GESTIÓN DE ESTADO Y CARGA DE DATOS
# ==============================================================================

if "raw_df" not in st.session_state:
    st.session_state.raw_df = None
if "data_source_name" not in st.session_state:
    st.session_state.data_source_name = None


def load_demo_data():
    """Genera o carga el archivo Excel sucio de demostración."""
    demo_path = Path("ventas_sucias_demo.xlsx")
    if not demo_path.exists():
        with st.spinner("Generando datos demo con anomalías intencionales..."):
            df_demo = generate_dirty_sales_data(filepath=demo_path, n_rows=320, seed=42)
    else:
        df_demo = pd.read_excel(demo_path)

    st.session_state.raw_df = df_demo
    st.session_state.data_source_name = "ventas_sucias_demo.xlsx (Demostración PyME)"


# ==============================================================================
# BARRA LATERAL (SIDEBAR)
# ==============================================================================
with st.sidebar:
    st.image(
        "https://img.icons8.com/fluency/96/combo-chart.png",
        width=64,
    )
    st.title("AutoReport PyME")
    st.caption("Ingeniería de Datos y Automatización Ejecutiva")
    st.markdown("---")

    st.subheader("1. Ingesta de Datos")

    # Botón para cargar conjunto de datos de demostración
    if st.button("🚀 Cargar Excel sucio de demostración", use_container_width=True, type="primary"):
        load_demo_data()
        st.success("¡Datos demo cargados con éxito!")

    st.markdown("**O sube tu propio archivo:**")
    uploaded_file = st.file_uploader(
        "Formatos soportados: Excel (.xlsx, .xls) o CSV (.csv)",
        type=["xlsx", "xls", "csv"],
        help="Sube un archivo transaccional de ventas para auditarlo y generar los reportes.",
    )

    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith(".csv"):
                df_uploaded = pd.read_csv(uploaded_file)
            else:
                df_uploaded = pd.read_excel(uploaded_file)
            st.session_state.raw_df = df_uploaded
            st.session_state.data_source_name = uploaded_file.name
            st.toast(f"Archivo cargado: {uploaded_file.name}", icon="📁")
        except Exception as e:
            st.error(f"Error al leer archivo: {e}")

    st.markdown("---")
    st.markdown("### 📌 Acerca de esta solución")
    st.info(
        "Pipeline de ingeniería de datos para pequeñas y medianas empresas: "
        "detecta vicios de captura manual, unifica fechas/sucursales y genera entregables listos para dirección."
    )
    st.caption("Desarrollado para entornos de PyME en México (MXN).")


# ==============================================================================
# CUERPO PRINCIPAL DEL DASHBOARD
# ==============================================================================

# Si aún no hay datos cargados, mostrar pantalla de bienvenida
if st.session_state.raw_df is None:
    st.markdown('<div class="main-title">AutoReport PyME 🇲🇽</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-title">Limpieza Inteligente de Datos, Auditoría Transaccional y Generación Automatizada de Reportes Directivos.</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        ### 👋 ¡Bienvenido al Sistema de Auditoría Comercial!
        Para comenzar, utiliza la barra lateral izquierda y elige una opción:
        1. **Haz clic en 'Cargar Excel sucio de demostración'** para probar el motor con ~300 registros que contienen duplicados, fechas mixtas, textos con espacios y montos corruptos.
        2. **O sube tu propio archivo** de ventas (.xlsx / .csv).
        """
    )

    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        st.markdown("#### 🧹 Saneamiento de Datos")
        st.write("Detección y purga de duplicados, corrección de fechas desalineadas y normalización de textos.")
    with col_info2:
        st.markdown("#### 📈 Inteligencia Comercial")
        st.write("Cálculo automático de Venta Bruta, Costo, Utilidad, Margen Neto y ranking de ejecutivos.")
    with col_info3:
        st.markdown("#### 📑 Entregables Directivos")
        st.write("Descarga inmediata de Excel con formato corporativo profesional y Reporte Ejecutivo en PDF.")

    st.stop()


# ==============================================================================
# PIPELINE DE LIMPIEZA Y CÁLCULO
# ==============================================================================
raw_df = st.session_state.raw_df

with st.spinner("Procesando datos a través del motor de ingeniería de datos..."):
    try:
        pipeline_results = clean_sales_data(raw_df)
    except Exception as exc:
        st.error(f"Error durante el procesamiento de datos: {exc}")
        st.stop()

cleaned_df = pipeline_results["cleaned_df"]
audit_metrics = pipeline_results["audit_metrics"]
financial_kpis = pipeline_results["financial_kpis"]

# Encabezado activo
st.markdown('<div class="main-title">AutoReport PyME | Consola de Control</div>', unsafe_allow_html=True)
st.caption(f"📁 Origen activo: **{st.session_state.data_source_name}** | Registros procesados: **{audit_metrics['filas_iniciales']:,}**")

# Pestañas principales
tab_auditoria, tab_dashboard, tab_descargas = st.tabs(
    [
        "🔍 1. Auditoría Antes/Después",
        "📊 2. Dashboard Ejecutivo (Plotly)",
        "📥 3. Descargar Entregables",
    ]
)


# ==============================================================================
# PESTAÑA 1: AUDITORÍA ANTES / DESPUÉS
# ==============================================================================
with tab_auditoria:
    st.subheader("Auditoría de Calidad e Integridad de Datos")
    st.write(
        "Diagnóstico comparativo del estado de los datos antes y después del pipeline de saneamiento."
    )

    # Métricas de calidad superiores
    m_col1, m_col2, m_col3, m_col4, m_col5 = st.columns(5)
    with m_col1:
        st.metric("Filas Recibidas", f"{audit_metrics['filas_iniciales']:,}")
    with m_col2:
        st.metric(
            "Duplicados Eliminados",
            f"{audit_metrics['duplicados_eliminados']:,}",
            delta=f"-{audit_metrics['duplicados_eliminados']}",
            delta_color="inverse",
        )
    with m_col3:
        st.metric(
            "Celdas Nulas Subsanadas",
            f"{audit_metrics['nulos_detectados_inicial']:,}",
            delta=f"-{audit_metrics['nulos_detectados_inicial']}",
            delta_color="inverse",
        )
    with m_col4:
        st.metric(
            "Registros Certificados",
            f"{audit_metrics['filas_finales']:,}",
            delta=f"{audit_metrics['tasa_retencion_pct']}% retención",
        )
    with m_col5:
        salud_val = audit_metrics["indice_salud_datos_pct"]
        st.metric("Índice Salud de Datos", f"{salud_val:.1f}%")

    st.markdown("---")

    # Comparativa lado a lado
    comp_col1, comp_col2 = st.columns(2)

    with comp_col1:
        st.markdown("#### ❌ Estado Inicial (Datos Crudos con Vicios)")
        st.caption("Se aprecian valores nulos, espacios en blanco residuales y formatos inconsistentes.")
        st.dataframe(
            raw_df.head(10),
            use_container_width=True,
            height=340,
        )

    with comp_col2:
        st.markdown("#### ✅ Estado Final (Datos Estructurados y Certificados)")
        st.caption("Textos limpios, fechas estandarizadas y columnas financieras calculadas.")
        st.dataframe(
            cleaned_df.head(10),
            use_container_width=True,
            height=340,
        )

    st.markdown("---")

    # Registro de correcciones aplicadas
    st.markdown("#### 📋 Bitácora Detallada de Correcciones Aplicadas")
    audit_table_data = [
        {"Regla de Auditoría": "Deduplicación Estricta", "Anomalías Detectadas": audit_metrics["duplicados_eliminados"], "Acción Realizada": "Eliminación de transacciones idénticas repetidas"},
        {"Regla de Auditoría": "Normalización de Sucursales y Vendedores", "Anomalías Detectadas": audit_metrics["textos_normalizados"], "Acción Realizada": "Remoción de espacios en blanco y homologación de mayúsculas/minúsculas"},
        {"Regla de Auditoría": "Parseo Multiformato de Fechas", "Anomalías Detectadas": audit_metrics["fechas_invalidas_filtradas"], "Acción Realizada": "Conversión a datetime (ISO) y descarte de textos no cronológicos"},
        {"Regla de Auditoría": "Validación Numérica y Precios", "Anomalías Detectadas": audit_metrics["registros_invalidos_filtrados"], "Acción Realizada": "Limpieza de signos $, comas de miles y filtro de montos <= 0"},
    ]
    st.table(pd.DataFrame(audit_table_data))


# ==============================================================================
# PESTAÑA 2: DASHBOARD EJECUTIVO (PLOTLY)
# ==============================================================================
with tab_dashboard:
    st.subheader("Tablero de Desempeño Comercial y Rentabilidad")

    # Filtros interactivos
    f_col1, f_col2 = st.columns(2)
    with f_col1:
        sucursales_disp = ["Todas"] + sorted(cleaned_df["Sucursal"].dropna().unique().tolist())
        selected_sucursal = st.selectbox("Filtrar por Sucursal:", sucursales_disp)

    with f_col2:
        categorias_disp = ["Todas"] + sorted(cleaned_df["Categoria"].dropna().unique().tolist())
        selected_categoria = st.selectbox("Filtrar por Categoría:", categorias_disp)

    # Filtrar DataFrame según selección
    df_filtered = cleaned_df.copy()
    if selected_sucursal != "Todas":
        df_filtered = df_filtered[df_filtered["Sucursal"] == selected_sucursal]
    if selected_categoria != "Todas":
        df_filtered = df_filtered[df_filtered["Categoria"] == selected_categoria]

    # Recalcular métricas dinámicas para el filtro
    f_ventas = float(df_filtered["Venta_Total"].sum())
    f_utilidad = float(df_filtered["Utilidad"].sum())
    f_margen = round((f_utilidad / f_ventas * 100), 1) if f_ventas > 0 else 0.0
    f_ticket = round(f_ventas / len(df_filtered), 2) if len(df_filtered) > 0 else 0.0
    f_txs = len(df_filtered)

    # Fila de Tarjetas KPIs
    k_col1, k_col2, k_col3, k_col4, k_col5 = st.columns(5)
    k_col1.metric("Venta Neta Total", f"${f_ventas:,.2f} MXN")
    k_col2.metric("Utilidad Bruta", f"${f_utilidad:,.2f} MXN")
    k_col3.metric("Margen Bruto", f"{f_margen:.1f}%")
    k_col4.metric("Ticket Promedio", f"${f_ticket:,.2f} MXN")
    k_col5.metric("Transacciones", f"{f_txs:,} txs")

    st.markdown("---")

    # Gráficos en Filas de 2 columnas
    g_col1, g_col2 = st.columns(2)

    with g_col1:
        # Gráfico 1: Ventas y Utilidad por Sucursal
        suc_summary = (
            df_filtered.groupby("Sucursal")
            .agg(Ventas=("Venta_Total", "sum"), Utilidad=("Utilidad", "sum"))
            .reset_index()
        )
        fig_suc = px.bar(
            suc_summary,
            x="Sucursal",
            y=["Ventas", "Utilidad"],
            barmode="group",
            title="Ventas vs. Utilidad por Sucursal (MXN)",
            color_discrete_sequence=["#1B365D", "#0F766E"],
            labels={"value": "Monto en MXN", "variable": "Concepto"},
        )
        fig_suc.update_layout(
            hovermode="x unified",
            margin=dict(l=20, r=20, t=40, b=20),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        )
        fig_suc.update_yaxes(tickprefix="$", tickformat=",.0f")
        st.plotly_chart(fig_suc, use_container_width=True)

    with g_col2:
        # Gráfico 2: Ventas por Categoría (Barras Horizontales)
        cat_summary = (
            df_filtered.groupby("Categoria")
            .agg(Ventas=("Venta_Total", "sum"))
            .reset_index()
            .sort_values(by="Ventas", ascending=True)
        )
        fig_cat = px.bar(
            cat_summary,
            x="Ventas",
            y="Categoria",
            orientation="h",
            title="Distribución de Ventas por Categoría",
            color="Ventas",
            color_continuous_scale="Blues",
            labels={"Ventas": "Monto en MXN", "Categoria": "Categoría"},
        )
        fig_cat.update_layout(
            margin=dict(l=20, r=20, t=40, b=20),
            coloraxis_showscale=False,
        )
        fig_cat.update_xaxes(tickprefix="$", tickformat=",.0f")
        st.plotly_chart(fig_cat, use_container_width=True)

    g_col3, g_col4 = st.columns(2)

    with g_col3:
        # Gráfico 3: Serie de Tiempo Mensual
        mes_summary = (
            df_filtered.groupby("Mes_Anio")
            .agg(Ventas=("Venta_Total", "sum"), Utilidad=("Utilidad", "sum"))
            .reset_index()
            .sort_values(by="Mes_Anio")
        )
        fig_time = px.line(
            mes_summary,
            x="Mes_Anio",
            y=["Ventas", "Utilidad"],
            markers=True,
            title="Tendencia Mensual de Ventas y Utilidad",
            color_discrete_sequence=["#0284C7", "#10B981"],
            labels={"value": "Monto en MXN", "Mes_Anio": "Periodo"},
        )
        fig_time.update_layout(
            hovermode="x unified",
            margin=dict(l=20, r=20, t=40, b=20),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        )
        fig_time.update_yaxes(tickprefix="$", tickformat=",.0f")
        st.plotly_chart(fig_time, use_container_width=True)

    with g_col4:
        # Gráfico 4: Ranking de Vendedores
        vend_summary = (
            df_filtered.groupby("Vendedor")
            .agg(Ventas=("Venta_Total", "sum"))
            .reset_index()
            .sort_values(by="Ventas", ascending=False)
            .head(6)
        )
        fig_vend = px.bar(
            vend_summary,
            x="Vendedor",
            y="Ventas",
            title="Ranking de Ventas por Ejecutivo",
            color="Ventas",
            color_continuous_scale="Teal",
            labels={"Ventas": "Venta Total MXN"},
        )
        fig_vend.update_layout(
            margin=dict(l=20, r=20, t=40, b=20),
            coloraxis_showscale=False,
        )
        fig_vend.update_yaxes(tickprefix="$", tickformat=",.0f")
        st.plotly_chart(fig_vend, use_container_width=True)


# ==============================================================================
# PESTAÑA 3: DESCARGAR ENTREGABLES
# ==============================================================================
with tab_descargas:
    st.subheader("Generación y Descarga de Entregables Ejecutivos")
    st.write(
        "Exporta de manera instantánea los datos certificados y consolidados en formato corporativo listo para comités directivos."
    )

    d_col1, d_col2 = st.columns(2)

    with d_col1:
        st.markdown("### 📊 Libro de Cálculo Excel (.xlsx)")
        st.markdown(
            """
            **Contenido del archivo:**
            - **Hoja 'Resumen Ejecutivo':** Tarjetas de KPIs financieros y auditoría de ingeniería de datos.
            - **Hoja 'Transacciones Limpias':** Registro fila a fila con formato numérico de moneda, fechas ISO y zebra-striping.
            - **Hoja 'Analítica por Sucursal':** Resumen consolidado con márgenes operativos.
            """
        )

        with st.spinner("Construyendo archivo Excel estilizado en memoria..."):
            excel_data = export_to_excel_bytes(cleaned_df, audit_metrics, financial_kpis)

        st.download_button(
            label="📥 Descargar Excel Limpio (.xlsx)",
            data=excel_data,
            file_name=f"Reporte_Ventas_PyME_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

    with d_col2:
        st.markdown("### 📄 Reporte Ejecutivo en PDF (.pdf)")
        st.markdown(
            """
            **Contenido del documento:**
            - **Página 1:** Tablero de control de KPIs financieros, diagnóstico de calidad y gráficas vectoriales incrustadas.
            - **Página 2:** Desglose detallado por sucursal, ranking de ejecutivos de venta y recomendaciones estratégicas.
            - Diseñado bajo lineamientos ejecutivos y codificación segura sin conflictos de texto.
            """
        )

        with st.spinner("Compilando Reporte Ejecutivo PDF en memoria..."):
            pdf_data = export_to_pdf_bytes(audit_metrics, financial_kpis)

        st.download_button(
            label="📄 Descargar Reporte Ejecutivo PDF (.pdf)",
            data=pdf_data,
            file_name=f"Reporte_Ejecutivo_PyME_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
            use_container_width=True,
            type="primary",
        )

    st.markdown("---")
    st.caption("AutoReport PyME © 2026 - Solución Integral de Ingeniería de Datos y Automatización.")
