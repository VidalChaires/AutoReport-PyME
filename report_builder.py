"""
report_builder.py
=================
Módulo para la generación y exportación de entregables ejecutivos en memoria:
1. Libro de cálculo Excel (.xlsx) altamente estilizado con openpyxl.
2. Reporte Ejecutivo en PDF (.pdf) formal de nivel directivo usando fpdf2 y matplotlib.

Ambas funciones operan 100% en memoria (BytesIO/bytes) para integrarse sin fricción
con APIs, pipelines de datos y dashboards web como Streamlit.
"""

from __future__ import annotations

import io
import unicodedata
from datetime import datetime
from typing import Any, Dict, Tuple

import matplotlib
matplotlib.use("Agg")  # Backend no interactivo para entornos headless
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from fpdf import FPDF
import pandas as pd


# ==============================================================================
# 1. EXPORTADOR DE EXCEL PROFESIONAL (OPENPYXL)
# ==============================================================================

def export_to_excel_bytes(
    cleaned_df: pd.DataFrame,
    audit_metrics: Dict[str, Any],
    financial_kpis: Dict[str, Any],
) -> bytes:
    """
    Genera un archivo Excel (.xlsx) en memoria con formato corporativo profesional:
    - Pestaña 'Resumen Ejecutivo' con tarjetas de KPIs y tabla de auditoría.
    - Pestaña 'Transacciones Limpias' con formato numérico y zebra striping.
    - Pestaña 'Analítica por Sucursal' con desglose consolidado.

    Returns
    -------
    bytes
        Contenido binario del archivo .xlsx listo para descarga o envío.
    """
    wb = openpyxl.Workbook()
    # Eliminar hoja por defecto
    wb.remove(wb.active)

    # Paleta de colores corporativos
    NAVY_DARK = "1B365D"
    NAVY_LIGHT = "E8EEF5"
    GRAY_LIGHT = "F8FAFC"
    GRAY_BORDER = "CBD5E1"
    ACCENT_GREEN = "0F766E"
    ACCENT_GREEN_BG = "CCFBF1"

    font_title = Font(name="Calibri", size=16, bold=True, color="1B365D")
    font_section = Font(name="Calibri", size=12, bold=True, color="1B365D")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=11, bold=True, color="000000")
    font_regular = Font(name="Calibri", size=11, color="1E293B")

    fill_header = PatternFill(start_color=NAVY_DARK, end_color=NAVY_DARK, fill_type="solid")
    fill_zebra = PatternFill(start_color=GRAY_LIGHT, end_color=GRAY_LIGHT, fill_type="solid")
    fill_card = PatternFill(start_color=NAVY_LIGHT, end_color=NAVY_LIGHT, fill_type="solid")
    fill_kpi_green = PatternFill(start_color=ACCENT_GREEN_BG, end_color=ACCENT_GREEN_BG, fill_type="solid")

    thin_border_side = Side(border_style="thin", color=GRAY_BORDER)
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    thick_bottom = Border(bottom=Side(border_style="medium", color=NAVY_DARK))

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # --------------------------------------------------------------------------
    # Hoja 1: Resumen Ejecutivo
    # --------------------------------------------------------------------------
    ws_resumen = wb.create_sheet(title="Resumen Ejecutivo")
    ws_resumen.views.sheetView[0].showGridLines = True

    ws_resumen.cell(row=2, column=2, value="DISTRIBUIDORA COMERCIAL PyMEX S.A. DE C.V.").font = font_title
    ws_resumen.cell(row=3, column=2, value=f"Reporte de Auditoría y Desempeño Comercial | Generado el {datetime.now().strftime('%d/%m/%Y %H:%M')}").font = Font(name="Calibri", size=10, italic=True, color="64748B")

    # Tarjetas de KPIs Financieros
    ws_resumen.cell(row=5, column=2, value="1. INDICADORES CLAVE DE RENDIMIENTO (KPIs)").font = font_section

    kpis_list = [
        ("Ventas Brutas Totales (MXN)", financial_kpis.get("venta_total_mxn", 0.0), "$#,##0.00"),
        ("Costo Total de Ventas (MXN)", financial_kpis.get("costo_total_mxn", 0.0), "$#,##0.00"),
        ("Utilidad Bruta (MXN)", financial_kpis.get("utilidad_total_mxn", 0.0), "$#,##0.00"),
        ("Margen Bruto Promedio", financial_kpis.get("margen_promedio_pct", 0.0) / 100.0, "0.0%"),
        ("Ticket Promedio por Venta", financial_kpis.get("ticket_promedio_mxn", 0.0), "$#,##0.00"),
        ("Transacciones Procesadas", financial_kpis.get("total_transacciones", 0), "#,##0"),
        ("Unidades Totales Vendidas", financial_kpis.get("total_unidades_vendidas", 0), "#,##0"),
    ]

    r_idx = 7
    for label, val, num_fmt in kpis_list:
        c_label = ws_resumen.cell(row=r_idx, column=2, value=label)
        c_label.font = font_bold
        c_label.fill = fill_card
        c_label.border = border_cell
        c_label.alignment = align_left

        c_val = ws_resumen.cell(row=r_idx, column=3, value=val)
        c_val.font = font_bold
        c_val.fill = fill_card
        c_val.border = border_cell
        c_val.alignment = align_right
        c_val.number_format = num_fmt
        r_idx += 1

    # Tabla de Auditoría de Calidad de Datos
    ws_resumen.cell(row=5, column=5, value="2. AUDITORÍA DE INGENIERÍA DE DATOS").font = font_section

    audit_list = [
        ("Registros Iniciales Recibidos", audit_metrics.get("filas_iniciales", 0), "#,##0"),
        ("Registros Duplicados Eliminados", audit_metrics.get("duplicados_eliminados", 0), "#,##0"),
        ("Celdas Nulas Iniciales Detectadas", audit_metrics.get("nulos_detectados_inicial", 0), "#,##0"),
        ("Fechas Inválidas / Descartadas", audit_metrics.get("fechas_invalidas_filtradas", 0), "#,##0"),
        ("Registros No Viables Descartados", audit_metrics.get("registros_invalidos_filtrados", 0), "#,##0"),
        ("Correcciones de Texto Aplicadas", audit_metrics.get("textos_normalizados", 0), "#,##0"),
        ("Registros Finales Certificados", audit_metrics.get("filas_finales", 0), "#,##0"),
        ("Índice de Salud de Datos", audit_metrics.get("indice_salud_datos_pct", 0.0) / 100.0, "0.0%"),
    ]

    r_idx_audit = 7
    for label, val, num_fmt in audit_list:
        c_label = ws_resumen.cell(row=r_idx_audit, column=5, value=label)
        c_label.font = font_regular
        c_label.fill = fill_kpi_green if "Salud" in label or "Certificados" in label else fill_zebra
        c_label.border = border_cell
        c_label.alignment = align_left

        c_val = ws_resumen.cell(row=r_idx_audit, column=6, value=val)
        c_val.font = font_bold
        c_val.fill = fill_kpi_green if "Salud" in label or "Certificados" in label else fill_zebra
        c_val.border = border_cell
        c_val.alignment = align_right
        c_val.number_format = num_fmt
        r_idx_audit += 1

    # Desempeño Destacado
    ws_resumen.cell(row=17, column=2, value="3. LIDERAZGO COMERCIAL").font = font_section
    leaders = [
        ("Sucursal Líder", financial_kpis.get("top_sucursal", "N/A")),
        ("Vendedor Destacado", financial_kpis.get("top_vendedor", "N/A")),
        ("Categoría Principal", financial_kpis.get("top_categoria", "N/A")),
    ]
    r_lead = 19
    for title_lead, val_lead in leaders:
        cl = ws_resumen.cell(row=r_lead, column=2, value=title_lead)
        cl.font = font_bold
        cl.fill = fill_card
        cl.border = border_cell
        cv = ws_resumen.cell(row=r_lead, column=3, value=val_lead)
        cv.font = font_regular
        cv.fill = fill_card
        cv.border = border_cell
        r_lead += 1

    # --------------------------------------------------------------------------
    # Hoja 2: Transacciones Limpias
    # --------------------------------------------------------------------------
    ws_data = wb.create_sheet(title="Transacciones Limpias")
    ws_data.views.sheetView[0].showGridLines = True

    display_df = cleaned_df.copy()
    if "Fecha" in display_df.columns:
        display_df["Fecha"] = display_df["Fecha"].dt.strftime("%Y-%m-%d")

    # Columnas seleccionadas para el reporte final
    report_columns = [
        "ID_Transaccion",
        "Fecha",
        "Sucursal",
        "Vendedor",
        "Categoria",
        "Producto",
        "Cantidad",
        "Precio_Unitario",
        "Costo_Unitario",
        "Venta_Total",
        "Costo_Total",
        "Utilidad",
        "Margen_Pct",
    ]
    report_columns = [c for c in report_columns if c in display_df.columns]

    # Escribir encabezados
    for col_idx, col_name in enumerate(report_columns, start=1):
        cell = ws_data.cell(row=1, column=col_idx, value=col_name.replace("_", " "))
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_cell

    # Escribir filas
    for row_idx, row_data in enumerate(display_df[report_columns].itertuples(index=False), start=2):
        is_even = (row_idx % 2 == 0)
        row_fill = fill_zebra if is_even else None

        for col_idx, val in enumerate(row_data, start=1):
            col_name = report_columns[col_idx - 1]
            cell = ws_data.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_regular
            cell.border = border_cell
            if row_fill:
                cell.fill = row_fill

            # Formateo condicional según tipo de columna
            if col_name in ["Precio_Unitario", "Costo_Unitario", "Venta_Total", "Costo_Total", "Utilidad"]:
                cell.number_format = "$#,##0.00"
                cell.alignment = align_right
            elif col_name == "Margen_Pct":
                cell.value = val / 100.0 if pd.notna(val) else 0.0
                cell.number_format = "0.0%"
                cell.alignment = align_right
            elif col_name in ["Cantidad"]:
                cell.number_format = "#,##0"
                cell.alignment = align_right
            elif col_name in ["ID_Transaccion", "Fecha", "Sucursal"]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left

    # --------------------------------------------------------------------------
    # Hoja 3: Analítica por Sucursal
    # --------------------------------------------------------------------------
    ws_branch = wb.create_sheet(title="Analitica por Sucursal")
    ws_branch.views.sheetView[0].showGridLines = True

    ws_branch.cell(row=2, column=2, value="DESGLOSE FINANCIERO CONSOLIDADO POR SUCURSAL").font = font_section

    agregado_suc = financial_kpis.get("agregado_sucursal", pd.DataFrame())
    if not agregado_suc.empty:
        branch_cols = ["Sucursal", "Transacciones", "Venta_Total", "Utilidad", "Margen_Pct"]
        for c_idx, c_name in enumerate(branch_cols, start=2):
            cell = ws_branch.cell(row=4, column=c_idx, value=c_name.replace("_", " "))
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = border_cell

        for r_offset, b_row in enumerate(agregado_suc.itertuples(index=False), start=5):
            for c_idx, c_name in enumerate(branch_cols, start=2):
                val = getattr(b_row, c_name)
                c = ws_branch.cell(row=r_offset, column=c_idx, value=val)
                c.font = font_regular
                c.border = border_cell
                if c_name in ["Venta_Total", "Utilidad"]:
                    c.number_format = "$#,##0.00"
                    c.alignment = align_right
                elif c_name == "Margen_Pct":
                    c.value = val / 100.0
                    c.number_format = "0.0%"
                    c.alignment = align_right
                elif c_name == "Transacciones":
                    c.number_format = "#,##0"
                    c.alignment = align_center
                else:
                    c.alignment = align_left

    # Autoajuste de anchos de columnas en todas las hojas
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Exportar a buffer en memoria
    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream.getvalue()


# ==============================================================================
# 2. GENERADOR DE REPORTES PDF EJECUTIVOS (FPDF2 + MATPLOTLIB)
# ==============================================================================

def _sanitize_for_latin1(text: Any) -> str:
    """
    Sanitiza cualquier texto para garantizar compatibilidad absoluta con fuentes
    estándar Latin-1 en FPDF2, evitando excepciones de codificación Unicode.
    """
    if pd.isna(text) or text is None:
        return ""
    text_str = str(text)
    # Reemplazar comillas inteligentes, guiones especiales y viñetas
    replacements = {
        "\u201c": '"',
        "\u201d": '"',
        "\u2018": "'",
        "\u2019": "'",
        "\u2013": "-",
        "\u2014": "-",
        "\u2022": "*",
        "\u2026": "...",
        "\u00a0": " ",
    }
    for orig, rep in replacements.items():
        text_str = text_str.replace(orig, rep)

    # Asegurar que se puede encodear en latin-1
    return text_str.encode("latin-1", errors="replace").decode("latin-1")


class ExecutivePDF(FPDF):
    """Subclase de FPDF2 con encabezados y pies de página corporativos."""

    def header(self):
        # Franja azul corporativa superior
        self.set_fill_color(27, 54, 93)  # #1B365D
        self.rect(0, 0, 210, 8, "F")

        # Título y logotipo textual
        self.set_xy(10, 12)
        self.set_font("Helvetica", "B", 14)
        self.set_text_color(27, 54, 93)
        self.cell(120, 8, _sanitize_for_latin1("AUTO REPORT PyME | REPORTE EJECUTIVO"), new_x="RIGHT", new_y="TOP", align="L")

        self.set_font("Helvetica", "I", 9)
        self.set_text_color(100, 116, 139)
        fecha_str = _sanitize_for_latin1(datetime.now().strftime("Fecha: %d/%m/%Y | Confidencial"))
        self.cell(70, 8, fecha_str, new_x="LMARGIN", new_y="NEXT", align="R")

        self.set_draw_color(203, 213, 225)
        self.set_line_width(0.3)
        self.line(10, 21, 200, 21)
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_draw_color(203, 213, 225)
        self.set_line_width(0.3)
        self.line(10, self.get_y(), 200, self.get_y())

        self.set_y(-12)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(100, 116, 139)
        self.cell(100, 6, _sanitize_for_latin1("AutoReport PyME México - Auditoría de Datos y Business Intelligence"), align="L")
        self.cell(90, 6, _sanitize_for_latin1(f"Página {self.page_no()}"), align="R")


def _generate_matplotlib_charts(financial_kpis: Dict[str, Any]) -> Tuple[io.BytesIO, io.BytesIO]:
    """Genera 2 gráficas de alta definición en memoria con estética profesional."""
    # Configuración de estilo sobrio corporativo
    plt.style.use("seaborn-v0_8-whitegrid" if "seaborn-v0_8-whitegrid" in plt.style.available else "default")
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = ["DejaVu Sans", "Arial", "Helvetica"]

    # 1. Gráfica de Ventas y Margen por Sucursal
    fig1, ax1 = plt.subplots(figsize=(6.5, 3.2), dpi=250)
    agregado_suc = financial_kpis.get("agregado_sucursal", pd.DataFrame())

    if not agregado_suc.empty:
        sucursales = agregado_suc["Sucursal"].tolist()
        ventas_millones = (agregado_suc["Venta_Total"] / 1000.0).tolist()
        colores = ["#1B365D", "#0284C7", "#0D9488"][:len(sucursales)]

        bars = ax1.bar(sucursales, ventas_millones, color=colores, width=0.55, edgecolor="none")
        ax1.set_title("Ventas Netas por Sucursal (Miles MXN)", fontsize=11, fontweight="bold", color="#1B365D", pad=10)
        ax1.set_ylabel("Monto ($k MXN)", fontsize=9, color="#475569")
        ax1.tick_params(axis="both", labelsize=9)

        # Etiquetas de datos sobre cada barra
        for bar in bars:
            height = bar.get_height()
            ax1.text(
                bar.get_x() + bar.get_width() / 2.0,
                height + (max(ventas_millones) * 0.02),
                f"${height:,.1f}k",
                ha="center",
                va="bottom",
                fontsize=8.5,
                fontweight="bold",
                color="#1E293B",
            )
        ax1.set_ylim(0, max(ventas_millones) * 1.20)
    else:
        ax1.text(0.5, 0.5, "Sin datos disponibles", ha="center", va="center")

    fig1.tight_layout()
    buf1 = io.BytesIO()
    fig1.savefig(buf1, format="png", dpi=250, bbox_inches="tight")
    buf1.seek(0)
    plt.close(fig1)

    # 2. Gráfica de Distribución de Ventas por Categoría (Barras horizontales)
    fig2, ax2 = plt.subplots(figsize=(6.5, 3.2), dpi=250)
    agregado_cat = financial_kpis.get("agregado_categoria", pd.DataFrame())

    if not agregado_cat.empty:
        cats = agregado_cat["Categoria"].tolist()[:5]  # Top 5
        ventas_cat = (agregado_cat["Venta_Total"][:5] / 1000.0).tolist()
        cats.reverse()
        ventas_cat.reverse()

        h_bars = ax2.barh(cats, ventas_cat, color="#0F766E", height=0.55, edgecolor="none")
        ax2.set_title("Top Categorías por Venta Total (Miles MXN)", fontsize=11, fontweight="bold", color="#1B365D", pad=10)
        ax2.set_xlabel("Venta ($k MXN)", fontsize=9, color="#475569")
        ax2.tick_params(axis="both", labelsize=8.5)

        for bar in h_bars:
            width = bar.get_width()
            ax2.text(
                width + (max(ventas_cat) * 0.02),
                bar.get_y() + bar.get_height() / 2.0,
                f"${width:,.1f}k",
                ha="left",
                va="center",
                fontsize=8,
                fontweight="bold",
                color="#1E293B",
            )
        ax2.set_xlim(0, max(ventas_cat) * 1.25)
    else:
        ax2.text(0.5, 0.5, "Sin datos disponibles", ha="center", va="center")

    fig2.tight_layout()
    buf2 = io.BytesIO()
    fig2.savefig(buf2, format="png", dpi=250, bbox_inches="tight")
    buf2.seek(0)
    plt.close(fig2)

    return buf1, buf2


def export_to_pdf_bytes(
    audit_metrics: Dict[str, Any],
    financial_kpis: Dict[str, Any],
) -> bytes:
    """
    Genera un informe ejecutivo en PDF de 2 páginas con diseño corporativo formal,
    tablas de auditoría de datos, KPIs financieros y dos gráficos incrustados.

    Returns
    -------
    bytes
        Contenido binario del documento PDF.
    """
    pdf = ExecutivePDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)

    # --------------------------------------------------------------------------
    # PÁGINA 1: AUDITORÍA DE DATOS & KPIS FINANCIEROS
    # --------------------------------------------------------------------------
    pdf.add_page()

    # Subtítulo de bienvenida
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(27, 54, 93)
    pdf.cell(0, 6, _sanitize_for_latin1("1. TABLERO DE CONTROL FINANCIERO Y AUDITORIA"), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    # Resumen de 4 tarjetas destacadas (Venta, Utilidad, Margen, Salud Datos)
    col_w = 46.0
    h_card = 16.0

    cards_data = [
        ("Ventas Totales", f"${financial_kpis.get('venta_total_mxn', 0.0):,.0f} MXN"),
        ("Utilidad Neta", f"${financial_kpis.get('utilidad_total_mxn', 0.0):,.0f} MXN"),
        ("Margen Promedio", f"{financial_kpis.get('margen_promedio_pct', 0.0):.1f}%"),
        ("Salud de Datos", f"{audit_metrics.get('indice_salud_datos_pct', 0.0):.1f}%"),
    ]

    for title, val in cards_data:
        x_curr = pdf.get_x()
        y_curr = pdf.get_y()

        # Fondo de tarjeta
        pdf.set_fill_color(241, 245, 249)  # Slate-100
        pdf.set_draw_color(203, 213, 225)
        pdf.rect(x_curr, y_curr, col_w, h_card, "DF")

        # Texto Título
        pdf.set_xy(x_curr + 2, y_curr + 2)
        pdf.set_font("Helvetica", "B", 8)
        pdf.set_text_color(100, 116, 139)
        pdf.cell(col_w - 4, 4, _sanitize_for_latin1(title), align="C")

        # Texto Valor
        pdf.set_xy(x_curr + 2, y_curr + 7)
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_text_color(15, 23, 42)
        pdf.cell(col_w - 4, 7, _sanitize_for_latin1(val), align="C")

        pdf.set_xy(x_curr + col_w + 2, y_curr)

    pdf.set_y(y_curr + h_card + 6)

    # Tablas Paralelas: Auditoría (Izq) y Métricas Financieras (Der)
    w_tbl = 92.0
    y_tbl_start = pdf.get_y()

    # Tabla 1: Auditoría
    pdf.set_xy(10, y_tbl_start)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(27, 54, 93)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(w_tbl, 6, _sanitize_for_latin1("Auditoría de Ingesta y Limpieza"), border=1, new_x="LMARGIN", new_y="NEXT", align="C", fill=True)

    audit_rows = [
        ("Registros Iniciales", f"{audit_metrics.get('filas_iniciales', 0):,}"),
        ("Duplicados Depurados", f"{audit_metrics.get('duplicados_eliminados', 0):,}"),
        ("Celdas Nulas Subsanadas", f"{audit_metrics.get('nulos_detectados_inicial', 0):,}"),
        ("Fechas Inválidas Filtradas", f"{audit_metrics.get('fechas_invalidas_filtradas', 0):,}"),
        ("Textos Normalizados", f"{audit_metrics.get('textos_normalizados', 0):,}"),
        ("Registros Certificados", f"{audit_metrics.get('filas_finales', 0):,}"),
    ]

    for label, val in audit_rows:
        pdf.set_x(10)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_text_color(30, 41, 59)
        pdf.cell(60, 5.5, _sanitize_for_latin1(label), border=1, align="L")
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(32, 5.5, _sanitize_for_latin1(val), border=1, align="R", new_x="LMARGIN", new_y="NEXT")

    # Tabla 2: Métricas Comerciales Complementarias
    pdf.set_xy(108, y_tbl_start)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(27, 54, 93)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(w_tbl, 6, _sanitize_for_latin1("Eficiencia y Rentabilidad Comercial"), border=1, new_x="LMARGIN", new_y="NEXT", align="C", fill=True)

    fin_rows = [
        ("Costo Total Mercancía", f"${financial_kpis.get('costo_total_mxn', 0.0):,.2f}"),
        ("Ticket Promedio de Venta", f"${financial_kpis.get('ticket_promedio_mxn', 0.0):,.2f}"),
        ("Unidades Totales Vendidas", f"{financial_kpis.get('total_unidades_vendidas', 0):,} pzs"),
        ("Transacciones Válidas", f"{financial_kpis.get('total_transacciones', 0):,} txs"),
        ("Sucursal Líder", str(financial_kpis.get("top_sucursal", "N/A")).split("(")[0].strip()),
        ("Vendedor Destacado", str(financial_kpis.get("top_vendedor", "N/A")).split("(")[0].strip()),
    ]

    for label, val in fin_rows:
        pdf.set_x(108)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_text_color(30, 41, 59)
        pdf.cell(56, 5.5, _sanitize_for_latin1(label), border=1, align="L")
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(36, 5.5, _sanitize_for_latin1(val), border=1, align="R", new_x="LMARGIN", new_y="NEXT")

    pdf.ln(6)

    # Gráficos generados con Matplotlib
    chart1_buf, chart2_buf = _generate_matplotlib_charts(financial_kpis)

    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(27, 54, 93)
    pdf.cell(0, 6, _sanitize_for_latin1("2. VISUALIZACIÓN DE DESEMPEÑO COMERCIAL"), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    y_charts = pdf.get_y()
    pdf.image(chart1_buf, x=10, y=y_charts, w=92)
    pdf.image(chart2_buf, x=108, y=y_charts, w=92)

    # --------------------------------------------------------------------------
    # PÁGINA 2: DESGLOSE DETALLADO & RECOMENDACIONES ESTRATÉGICAS
    # --------------------------------------------------------------------------
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(27, 54, 93)
    pdf.cell(0, 6, _sanitize_for_latin1("3. DESGLOSE CONSOLIDADO POR SUCURSAL Y VENDEDOR"), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    # Tabla Sucursales
    agregado_suc = financial_kpis.get("agregado_sucursal", pd.DataFrame())
    if not agregado_suc.empty:
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_fill_color(27, 54, 93)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(45, 6, _sanitize_for_latin1("Sucursal"), border=1, fill=True, align="L")
        pdf.cell(35, 6, _sanitize_for_latin1("Transacciones"), border=1, fill=True, align="C")
        pdf.cell(40, 6, _sanitize_for_latin1("Venta Total"), border=1, fill=True, align="R")
        pdf.cell(40, 6, _sanitize_for_latin1("Utilidad"), border=1, fill=True, align="R")
        pdf.cell(30, 6, _sanitize_for_latin1("Margen %"), border=1, fill=True, align="R", new_x="LMARGIN", new_y="NEXT")

        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_text_color(30, 41, 59)
        for row in agregado_suc.itertuples():
            pdf.cell(45, 5.5, _sanitize_for_latin1(row.Sucursal), border=1, align="L")
            pdf.cell(35, 5.5, f"{row.Transacciones:,}", border=1, align="C")
            pdf.cell(40, 5.5, f"${row.Venta_Total:,.2f}", border=1, align="R")
            pdf.cell(40, 5.5, f"${row.Utilidad:,.2f}", border=1, align="R")
            pdf.cell(30, 5.5, f"{row.Margen_Pct:.1f}%", border=1, align="R", new_x="LMARGIN", new_y="NEXT")

    pdf.ln(6)

    # Tabla Top 5 Vendedores
    agregado_vend = financial_kpis.get("agregado_vendedores", pd.DataFrame())
    if not agregado_vend.empty:
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_text_color(27, 54, 93)
        pdf.cell(0, 5, _sanitize_for_latin1("Desempeño de la Fuerza de Ventas (Top 5)"), new_x="LMARGIN", new_y="NEXT")
        pdf.ln(1)

        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_fill_color(15, 118, 110)  # Accent Teal
        pdf.set_text_color(255, 255, 255)
        pdf.cell(60, 6, _sanitize_for_latin1("Ejecutivo de Ventas"), border=1, fill=True, align="L")
        pdf.cell(35, 6, _sanitize_for_latin1("Unidades"), border=1, fill=True, align="C")
        pdf.cell(45, 6, _sanitize_for_latin1("Venta Lograda"), border=1, fill=True, align="R")
        pdf.cell(50, 6, _sanitize_for_latin1("Contribución a Utilidad"), border=1, fill=True, align="R", new_x="LMARGIN", new_y="NEXT")

        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_text_color(30, 41, 59)
        for row in agregado_vend.head(5).itertuples():
            pdf.cell(60, 5.5, _sanitize_for_latin1(row.Vendedor), border=1, align="L")
            pdf.cell(35, 5.5, f"{row.Unidades:,}", border=1, align="C")
            pdf.cell(45, 5.5, f"${row.Venta_Total:,.2f}", border=1, align="R")
            pdf.cell(50, 5.5, f"${row.Utilidad:,.2f}", border=1, align="R", new_x="LMARGIN", new_y="NEXT")

    pdf.ln(8)

    # Sección de Conclusiones y Diagnóstico de Calidad
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(27, 54, 93)
    pdf.cell(0, 6, _sanitize_for_latin1("4. DIAGNÓSTICO DE CALIDAD Y ACCIONES SUGERIDAS"), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(1)

    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(51, 65, 85)
    rec_text = (
        "- Calidad de Ingesta: Se detectaron inconsistencias recurrentes de captura manual en nombres "
        "de sucursales y vendedores (espacios y casing irregular). Se recomienda implementar validación de listas en origen.\n"
        "- Estandarización de Fechas: Se identificaron fechas ambiguas y textos no numéricos. Se sugiere fijar formato ISO YYYY-MM-DD en el ERP.\n"
        "- Desempeño Comercial: Se observa un margen operativo saludable. La sucursal con mayor tracción "
        f"es {str(financial_kpis.get('top_sucursal', 'N/A')).split('(')[0].strip()}.\n"
        "- Política de Duplicados: Se eliminaron registros repetidos idénticos para evitar distorsión en la contabilidad."
    )
    pdf.multi_cell(0, 5.2, _sanitize_for_latin1(rec_text), border=0)

    # Retorno en bytes usando fpdf2
    pdf_bytes = pdf.output()
    if isinstance(pdf_bytes, bytearray):
        return bytes(pdf_bytes)
    return pdf_bytes
