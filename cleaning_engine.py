"""
cleaning_engine.py
==================
Motor de ingeniería de datos para la ingesta, auditoría, saneamiento y cálculo
de KPIs comerciales y financieros para PyMEs mexicanas.

Responsabilidades:
1. Auditoría inicial del dataset (conteo de duplicados, valores ausentes, tipos erróneos).
2. Deduplicación determinista.
3. Normalización y estandarización de cadenas de texto (Sucursal, Vendedor, Categoría).
4. Saneamiento robusto de fechas multiformato (ISO, DD/MM/AAAA, MM-DD-AAAA).
5. Limpieza de columnas numéricas (remoción de símbolos $, MXN, comas, unidades).
6. Reglas de validación de negocio (filtros de cantidades/precios no positivos o nulos).
7. Cálculo de variables financieras derivadas (Venta Total, Costo Total, Utilidad, Margen %).
8. Generación de diccionario unificado con métricas de auditoría y KPIs ejecutivos.
"""

from __future__ import annotations

import re
from typing import Any, Dict, Tuple

import numpy as np
import pandas as pd


def _clean_numeric_series(series: pd.Series) -> pd.Series:
    """
    Convierte una serie con texto o números a flotante limpio.
    Remueve símbolos de divisa ($), texto MXN, comas de miles y sufijos como 'pzs'.
    """
    def parse_value(val: Any) -> float:
        if pd.isna(val):
            return np.nan
        if isinstance(val, (int, float)):
            return float(val)
        val_str = str(val).strip()
        # Eliminar signos de divisa, comas, palabras como MXN, pzs
        cleaned = re.sub(r"[^\d.-]", "", val_str)
        try:
            return float(cleaned)
        except (ValueError, TypeError):
            return np.nan

    return series.apply(parse_value)


def _normalize_branch_name(val: Any) -> Any:
    """Normaliza nombres de sucursales a sus valores canónicos en México."""
    if pd.isna(val):
        return np.nan
    val_clean = str(val).strip().upper()
    val_clean = re.sub(r"\s+", " ", val_clean)

    if "MONTERREY" in val_clean:
        return "Monterrey"
    if "GUADALAJARA" in val_clean:
        return "Guadalajara"
    if "CDMX" in val_clean or "MEXICO" in val_clean or "DISTRITO" in val_clean:
        return "CDMX"
    return str(val).strip().title()


def _normalize_text_field(val: Any, uppercase: bool = False) -> Any:
    """Limpia espacios en blanco residuales y estandariza mayúsculas/minúsculas."""
    if pd.isna(val):
        return np.nan
    cleaned = re.sub(r"\s+", " ", str(val).strip())
    if not cleaned or cleaned.upper() in {"NAN", "NONE", "NULL", "N/A"}:
        return np.nan
    return cleaned.upper() if uppercase else cleaned.title()


def clean_sales_data(raw_df: pd.DataFrame) -> Dict[str, Any]:
    """
    Ejecuta el pipeline completo de limpieza, validación y analítica financiera.

    Parameters
    ----------
    raw_df : pd.DataFrame
        DataFrame de entrada sin procesar.

    Returns
    -------
    Dict[str, Any]
        Diccionario con las claves:
        - 'cleaned_df': pd.DataFrame limpio y estructurado con tipos nativos correctos.
        - 'audit_metrics': Diccionario con estadísticas de calidad y depuración.
        - 'financial_kpis': Diccionario con métricas agregadas y de rendimiento financiero.
    """
    if raw_df is None or raw_df.empty:
        raise ValueError("El DataFrame proporcionado está vacío o es nulo.")

    df = raw_df.copy()
    initial_rows = len(df)
    initial_null_cells = int(df.isna().sum().sum())
    textos_modificados_count = 0

    # --- 1. NORMALIZACIÓN DE TEXTO ---
    # Limpiar espacios residuales y homologar casing ANTES de deduplicar
    if "ID_Transaccion" in df.columns:
        df["ID_Transaccion"] = df["ID_Transaccion"].apply(
            lambda x: re.sub(r"\s+", " ", str(x).strip()) if pd.notna(x) else np.nan
        )

    if "Sucursal" in df.columns:
        original_branches = df["Sucursal"].copy()
        df["Sucursal"] = df["Sucursal"].apply(_normalize_branch_name)
        textos_modificados_count += int((original_branches.astype(str) != df["Sucursal"].astype(str)).sum())

    if "Vendedor" in df.columns:
        original_sellers = df["Vendedor"].copy()
        df["Vendedor"] = df["Vendedor"].apply(lambda v: _normalize_text_field(v, uppercase=False))
        textos_modificados_count += int((original_sellers.astype(str) != df["Vendedor"].astype(str)).sum())

    if "Categoria" in df.columns:
        original_cats = df["Categoria"].copy()
        df["Categoria"] = df["Categoria"].apply(lambda c: _normalize_text_field(c, uppercase=False))
        textos_modificados_count += int((original_cats.astype(str) != df["Categoria"].astype(str)).sum())

    if "Producto" in df.columns:
        df["Producto"] = df["Producto"].apply(lambda p: re.sub(r"\s+", " ", str(p).strip()) if pd.notna(p) else np.nan)

    # --- 2. NORMALIZACIÓN DE FECHAS ---
    fechas_invalidas_count = 0
    if "Fecha" in df.columns:
        # Intentar parseo multiformato moderno con Pandas 2.0+
        fechas_parseadas = pd.to_datetime(df["Fecha"], errors="coerce", format="mixed")
        fechas_invalidas_count = int(fechas_parseadas.isna().sum())
        df["Fecha"] = fechas_parseadas

    # --- 3. SANEAMIENTO NUMÉRICO ---
    # Forzar conversión a numérico (removiendo $, MXN, comas) para que las columnas
    # numéricas tengan tipos homogéneos y se puedan detectar duplicados reales
    for col in ["Cantidad", "Precio_Unitario", "Costo_Unitario"]:
        if col in df.columns:
            df[col] = _clean_numeric_series(df[col])

    # --- 4. DEDUPLICACIÓN INTELIGENTE TRAS HOMOLOGACIÓN ---
    # Una vez que los textos no tienen espacios residuales y los números son flotantes puros,
    # identificamos los registros duplicados que antes estaban enmascarados por ruido
    filas_antes_dedup = len(df)
    df = df.drop_duplicates().reset_index(drop=True)
    if "ID_Transaccion" in df.columns:
        df = df.drop_duplicates(subset=["ID_Transaccion"], keep="first").reset_index(drop=True)
    duplicates_removed = filas_antes_dedup - len(df)

    # --- 5. FILTRADO DE REGLAS DE NEGOCIO Y CALIDAD ---
    # Una transacción válida debe tener:
    # - Sucursal y Vendedor no nulos
    # - Fecha válida
    # - Cantidad > 0 (entera)
    # - Precio_Unitario > 0
    # - Costo_Unitario >= 0
    cond_sucursal = df["Sucursal"].notna()
    cond_vendedor = df["Vendedor"].notna()
    cond_fecha = df["Fecha"].notna()
    cond_cantidad = df["Cantidad"].notna() & (df["Cantidad"] > 0)
    cond_precio = df["Precio_Unitario"].notna() & (df["Precio_Unitario"] > 0)
    cond_costo = df["Costo_Unitario"].notna() & (df["Costo_Unitario"] >= 0)

    mascara_valida = (
        cond_sucursal
        & cond_vendedor
        & cond_fecha
        & cond_cantidad
        & cond_precio
        & cond_costo
    )

    invalid_rows_count = int((~mascara_valida).sum())
    df_clean = df[mascara_valida].copy()

    # Tipado explícito de columnas
    df_clean["Cantidad"] = df_clean["Cantidad"].round().astype(int)
    df_clean["Precio_Unitario"] = df_clean["Precio_Unitario"].round(2)
    df_clean["Costo_Unitario"] = df_clean["Costo_Unitario"].round(2)

    # --- 6. ASEGURAMIENTO FINAL DE DEDUPLICACIÓN ---
    # Garantiza que ninguna fila idéntica persista en el entregable certificado
    filas_clean_antes = len(df_clean)
    if "ID_Transaccion" in df_clean.columns:
        df_clean = df_clean.drop_duplicates(subset=["ID_Transaccion"], keep="first").reset_index(drop=True)
    df_clean = df_clean.drop_duplicates().reset_index(drop=True)
    duplicates_removed += (filas_clean_antes - len(df_clean))

    # --- 7. CÁLCULO DE KPIs FINANCIEROS Y DERIVADOS ---
    df_clean["Venta_Total"] = (df_clean["Cantidad"] * df_clean["Precio_Unitario"]).round(2)
    df_clean["Costo_Total"] = (df_clean["Cantidad"] * df_clean["Costo_Unitario"]).round(2)
    df_clean["Utilidad"] = (df_clean["Venta_Total"] - df_clean["Costo_Total"]).round(2)
    # Margen porcentual protegido contra división por cero
    df_clean["Margen_Pct"] = np.where(
        df_clean["Venta_Total"] > 0,
        ((df_clean["Utilidad"] / df_clean["Venta_Total"]) * 100).round(2),
        0.0,
    )

    # Añadir columna Año-Mes para análisis temporal
    df_clean["Mes_Anio"] = df_clean["Fecha"].dt.strftime("%Y-%m")
    df_clean["Dia_Semana"] = df_clean["Fecha"].dt.day_name(locale="Spanish") if hasattr(df_clean["Fecha"].dt, "day_name") else df_clean["Fecha"].dt.strftime("%A")

    # Ordenar cronológicamente
    df_clean = df_clean.sort_values(by=["Fecha", "ID_Transaccion"]).reset_index(drop=True)

    final_rows = len(df_clean)
    retention_rate = round((final_rows / initial_rows) * 100, 2) if initial_rows > 0 else 0.0

    # --- 8. MÉTRICAS DE AUDITORÍA ---
    audit_metrics: Dict[str, Any] = {
        "filas_iniciales": initial_rows,
        "filas_finales": final_rows,
        "filas_descartadas_total": initial_rows - final_rows,
        "duplicados_eliminados": duplicates_removed,
        "nulos_detectados_inicial": initial_null_cells,
        "fechas_invalidas_filtradas": fechas_invalidas_count,
        "registros_invalidos_filtrados": invalid_rows_count,
        "textos_normalizados": textos_modificados_count,
        "tasa_retencion_pct": retention_rate,
        "indice_salud_datos_pct": round(100.0 - ((initial_rows - final_rows) / initial_rows * 100), 2) if initial_rows > 0 else 100.0,
    }

    # --- 8. KPIs FINANCIEROS Y COMERCIALES ---
    total_sales = float(df_clean["Venta_Total"].sum())
    total_cost = float(df_clean["Costo_Total"].sum())
    total_profit = float(df_clean["Utilidad"].sum())
    overall_margin = round((total_profit / total_sales * 100), 2) if total_sales > 0 else 0.0
    avg_ticket = round(total_sales / final_rows, 2) if final_rows > 0 else 0.0
    total_units = int(df_clean["Cantidad"].sum())

    # Agregaciones estratégicas
    ventas_sucursal_df = (
        df_clean.groupby("Sucursal")
        .agg(
            Venta_Total=("Venta_Total", "sum"),
            Utilidad=("Utilidad", "sum"),
            Transacciones=("ID_Transaccion", "count"),
        )
        .reset_index()
        .sort_values(by="Venta_Total", ascending=False)
    )
    ventas_sucursal_df["Margen_Pct"] = (
        (ventas_sucursal_df["Utilidad"] / ventas_sucursal_df["Venta_Total"]) * 100
    ).round(2)

    ventas_categoria_df = (
        df_clean.groupby("Categoria")
        .agg(
            Venta_Total=("Venta_Total", "sum"),
            Utilidad=("Utilidad", "sum"),
            Unidades=("Cantidad", "sum"),
        )
        .reset_index()
        .sort_values(by="Venta_Total", ascending=False)
    )
    ventas_categoria_df["Margen_Pct"] = (
        (ventas_categoria_df["Utilidad"] / ventas_categoria_df["Venta_Total"]) * 100
    ).round(2)

    ventas_mes_df = (
        df_clean.groupby("Mes_Anio")
        .agg(
            Venta_Total=("Venta_Total", "sum"),
            Utilidad=("Utilidad", "sum"),
            Transacciones=("ID_Transaccion", "count"),
        )
        .reset_index()
        .sort_values(by="Mes_Anio")
    )

    ranking_vendedores_df = (
        df_clean.groupby("Vendedor")
        .agg(
            Venta_Total=("Venta_Total", "sum"),
            Utilidad=("Utilidad", "sum"),
            Transacciones=("ID_Transaccion", "count"),
            Unidades=("Cantidad", "sum"),
        )
        .reset_index()
        .sort_values(by="Venta_Total", ascending=False)
    )

    top_sucursal = (
        f"{ventas_sucursal_df.iloc[0]['Sucursal']} (${ventas_sucursal_df.iloc[0]['Venta_Total']:,.2f} MXN)"
        if not ventas_sucursal_df.empty
        else "N/A"
    )
    top_vendedor = (
        f"{ranking_vendedores_df.iloc[0]['Vendedor']} (${ranking_vendedores_df.iloc[0]['Venta_Total']:,.2f} MXN)"
        if not ranking_vendedores_df.empty
        else "N/A"
    )
    top_categoria = (
        f"{ventas_categoria_df.iloc[0]['Categoria']} (${ventas_categoria_df.iloc[0]['Venta_Total']:,.2f} MXN)"
        if not ventas_categoria_df.empty
        else "N/A"
    )

    financial_kpis: Dict[str, Any] = {
        "venta_total_mxn": total_sales,
        "costo_total_mxn": total_cost,
        "utilidad_total_mxn": total_profit,
        "margen_promedio_pct": overall_margin,
        "ticket_promedio_mxn": avg_ticket,
        "total_transacciones": final_rows,
        "total_unidades_vendidas": total_units,
        "top_sucursal": top_sucursal,
        "top_vendedor": top_vendedor,
        "top_categoria": top_categoria,
        "agregado_sucursal": ventas_sucursal_df,
        "agregado_categoria": ventas_categoria_df,
        "agregado_mes": ventas_mes_df,
        "agregado_vendedores": ranking_vendedores_df,
    }

    return {
        "cleaned_df": df_clean,
        "audit_metrics": audit_metrics,
        "financial_kpis": financial_kpis,
    }
