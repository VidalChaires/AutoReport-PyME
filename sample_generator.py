"""
sample_generator.py
===================
Generador de datos sintéticos con anomalías intencionales de calidad de datos
para simular el ecosistema transaccional típico de una PyME distribuidora en México.

Genera ~300 registros con sucursales (Monterrey, Guadalajara, CDMX), montos en MXN,
vendedores, categorías, productos, y diversos vicios de captura manual:
- Registros duplicados exactos y parciales.
- Espacios en blanco adicionales y casing inconsistente (" monterrey ", "CDMX  ").
- Fechas en formatos mixtos (ISO, DD/MM/YYYY, timestamps, textos no parseables).
- Valores nulos (None, NaN, celdas vacías).
- Cifras con formato de texto (ej. "$1,250.50", "N/A", cantidades negativas por error).
"""

from __future__ import annotations

import random
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, List, Optional

import numpy as np
import pandas as pd


def generate_dirty_sales_data(
    filepath: Optional[str | Path] = "ventas_sucias_demo.xlsx",
    n_rows: int = 320,
    seed: int = 42,
) -> pd.DataFrame:
    """
    Genera un DataFrame y archivo Excel con transacciones comerciales 'sucias'.

    Parameters
    ----------
    filepath : str | Path | None, default 'ventas_sucias_demo.xlsx'
        Ruta donde se guardará el archivo Excel generado. Si es None, solo devuelve el DataFrame.
    n_rows : int, default 320
        Número aproximado de registros base antes de inyectar duplicados.
    seed : int, default 42
        Semilla para asegurar reproducibilidad.

    Returns
    -------
    pd.DataFrame
        DataFrame crudo con anomalías intencionales de calidad de datos.
    """
    random.seed(seed)
    np.random.seed(seed)

    # Catálogos base de la PyME mexicana
    sucursales = ["Monterrey", "Guadalajara", "CDMX"]
    vendedores = [
        "Ana Sofía Garza",
        "Carlos Mendoza",
        "Luis Fernando Morales",
        "Valeria Peña",
        "Jorge Ramírez",
        "Mariana Hernández",
    ]
    catalogo_productos = [
        # (Categoria, SKU - Producto, Costo_Base, Precio_Base)
        ("Electrónica y Cómputo", "LAP-001 Laptop Corporativa 15.6\"", 11500.0, 16800.0),
        ("Electrónica y Cómputo", "MON-002 Monitor UltraWide 29\"", 4200.0, 6499.0),
        ("Electrónica y Cómputo", "TEC-003 Combo Teclado y Mouse Inalámbrico", 450.0, 890.0),
        ("Mobiliario de Oficina", "ESC-101 Escritorio Ergonómico Ajustable", 3800.0, 5950.0),
        ("Mobiliario de Oficina", "SIL-102 Silla Ejecutiva Mesh Lumbar", 2100.0, 3600.0),
        ("Mobiliario de Oficina", "ARC-103 Archivero Metálico 3 Gavetas", 1850.0, 2990.0),
        ("Papelería Corporativa", "PAP-201 Resma Papel Bond Carta 75g (Caja 5k)", 620.0, 980.0),
        ("Papelería Corporativa", "BOL-202 Bolígrafos Gel Negro (Caja 50 pzs)", 180.0, 320.0),
        ("Papelería Corporativa", "TON-203 Tóner Láser Alto Rendimiento", 1200.0, 2150.0),
        ("Insumos y Limpieza", "DES-301 Sanitizante Grado Hospitalario 20L", 480.0, 850.0),
        ("Insumos y Limpieza", "PAP-302 Dispensador Toalla Interdoblada", 290.0, 540.0),
        ("Seguridad Industrial", "BOT-401 Bota de Seguridad Dieléctrica", 750.0, 1350.0),
        ("Seguridad Industrial", "CAS-402 Casco de Protección con Barboquejo", 160.0, 310.0),
    ]

    base_date = datetime(2024, 1, 1)
    rows: List[dict[str, Any]] = []

    # 1. Generación de registros base consistentes
    for i in range(1, n_rows + 1):
        prod_tuple = random.choice(catalogo_productos)
        cat, prod_nombre, costo_base, precio_base = prod_tuple

        # Variación aleatoria de costo y precio (+/- 10%)
        factor_costo = random.uniform(0.92, 1.08)
        factor_precio = random.uniform(0.95, 1.12)
        costo_unitario = round(costo_base * factor_costo, 2)
        precio_unitario = round(precio_base * factor_precio, 2)

        # Cantidad vendida habitual entre 1 y 25 unidades
        cantidad = random.randint(1, 25)

        # Fecha distribuida en un lapso de 180 días
        delta_days = random.randint(0, 180)
        trans_date = base_date + timedelta(days=delta_days, hours=random.randint(8, 19))

        rows.append(
            {
                "ID_Transaccion": f"TX-2024-{i:04d}",
                "Fecha": trans_date.strftime("%Y-%m-%d"),
                "Sucursal": random.choice(sucursales),
                "Vendedor": random.choice(vendedores),
                "Categoria": cat,
                "Producto": prod_nombre,
                "Cantidad": cantidad,
                "Precio_Unitario": precio_unitario,
                "Costo_Unitario": costo_unitario,
            }
        )

    df = pd.DataFrame(rows)

    # 2. Inyección intencional de anomalías y vicios de captura de datos

    # A) Duplicados exactos (copiar 15 filas al final)
    duplicate_indices = random.sample(range(len(df)), k=15)
    df_duplicates = df.iloc[duplicate_indices].copy()
    df = pd.concat([df, df_duplicates], ignore_index=True)

    # Convertir columnas a tipo object para permitir inyección de tipos heterogéneos
    # (strings en columnas numéricas, textos en fechas, etc.) sin error en pandas 3.0+
    df = df.astype(object)

    # B) Textos desalineados, espacios irregulares y casing inconsistente en Sucursal
    dirty_sucursales = [
        " monterrey ",
        "  MONTERREY",
        "monterrey",
        " CDMX ",
        "cdmx",
        "CDMX   ",
        "  guadalajara",
        "GUADALAJARA ",
        "Guadalajara   ",
    ]
    for idx in random.sample(range(len(df)), k=40):
        df.at[idx, "Sucursal"] = random.choice(dirty_sucursales)

    # Casing y espacios en Vendedor
    for idx in random.sample(range(len(df)), k=30):
        val = str(df.at[idx, "Vendedor"])
        if random.random() < 0.5:
            df.at[idx, "Vendedor"] = f"  {val.lower()}  "
        else:
            df.at[idx, "Vendedor"] = f"{val.upper()} "

    # Casing y espacios en Categoría
    for idx in random.sample(range(len(df)), k=25):
        cat_val = str(df.at[idx, "Categoria"])
        df.at[idx, "Categoria"] = f"  {cat_val.upper()}  "

    # C) Fechas en formatos mixtos y valores no parseables
    for idx in random.sample(range(len(df)), k=12):
        # Formato mexicano / europeo DD/MM/YYYY
        df.at[idx, "Fecha"] = "15/03/2024" if idx % 2 == 0 else "28/05/2024"
    for idx in random.sample(range(len(df)), k=8):
        # Formato estadounidense MM-DD-YYYY
        df.at[idx, "Fecha"] = "04-18-2024"
    for idx in random.sample(range(len(df)), k=6):
        # Timestamp completo
        df.at[idx, "Fecha"] = "2024-02-14 17:45:22"
    for idx in random.sample(range(len(df)), k=7):
        # Errores de texto tipográfico en fecha
        df.at[idx, "Fecha"] = random.choice(["PENDIENTE", "SIN_FECHA", "2024-02-30", "N/A"])

    # D) Inyección de valores nulos (NaN / None)
    for idx in random.sample(range(len(df)), k=8):
        df.at[idx, "Sucursal"] = np.nan
    for idx in random.sample(range(len(df)), k=6):
        df.at[idx, "Vendedor"] = None
    for idx in random.sample(range(len(df)), k=7):
        df.at[idx, "Precio_Unitario"] = np.nan
    for idx in random.sample(range(len(df)), k=5):
        df.at[idx, "Cantidad"] = None
    for idx in random.sample(range(len(df)), k=6):
        df.at[idx, "Costo_Unitario"] = np.nan

    # E) Formato de texto sucio en números (símbolos de moneda, comas, strings)
    for idx in random.sample(range(len(df)), k=18):
        orig_val = df.at[idx, "Precio_Unitario"]
        if pd.notna(orig_val) and isinstance(orig_val, (int, float)):
            # Formato string como "$1,500.50 MXN" o " $ 1,500.50 "
            df.at[idx, "Precio_Unitario"] = f"$ {orig_val:,.2f} MXN"

    for idx in random.sample(range(len(df)), k=10):
        orig_cost = df.at[idx, "Costo_Unitario"]
        if pd.notna(orig_cost) and isinstance(orig_cost, (int, float)):
            df.at[idx, "Costo_Unitario"] = f"${orig_cost:,.2f}"

    for idx in random.sample(range(len(df)), k=6):
        orig_qty = df.at[idx, "Cantidad"]
        if pd.notna(orig_qty):
            df.at[idx, "Cantidad"] = f" {orig_qty} pzs "

    # F) Errores lógicos en cantidades (valores negativos o cero accidentales)
    for idx in random.sample(range(len(df)), k=5):
        df.at[idx, "Cantidad"] = -random.randint(1, 10)

    # Mezclar las filas para realismo
    df = df.sample(frac=1.0, random_state=seed).reset_index(drop=True)

    # Guardar en archivo Excel si se especificó ruta
    if filepath is not None:
        target_path = Path(filepath)
        target_path.parent.mkdir(parents=True, exist_ok=True)
        df.to_excel(target_path, index=False, engine="openpyxl")

    return df


if __name__ == "__main__":
    output_file = "ventas_sucias_demo.xlsx"
    print(f"Generando conjunto de datos sucio de demostración: {output_file}...")
    df_raw = generate_dirty_sales_data(output_file)
    print(f"Éxito: Se generaron {len(df_raw)} filas con vicios intencionales en '{output_file}'.")
