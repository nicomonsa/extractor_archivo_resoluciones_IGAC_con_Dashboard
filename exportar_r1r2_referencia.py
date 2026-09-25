#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
 EXPORTADOR R1 / R2 AL FORMATO DE LOS EXCEL DE REFERENCIA
 Convenio IGAC - Municipio de Mosquera (Cundinamarca)
================================================================================

QUE HACE
    Toma los archivos planos R1 y R2 y los escribe en Excel con EXACTAMENTE las
    mismas columnas, en el mismo orden y con los mismos nombres que los archivos
    de referencia del IGAC:

        GLPI..._R1.xlsx -> 18 columnas
        GLPI..._R2.xlsx -> 46 columnas (incluye las columnas de espacio)

    A diferencia del parser, aqui NO se agrega ni una sola columna calculada.
    Es una transcripcion fiel del archivo plano: los valores salen tal como
    estan en el texto, conservando ceros a la izquierda y coma decimal.

COMO SE USA
    python exportar_r1r2_referencia.py <carpeta_entrada> <carpeta_salida>
    o presionar play y usa las rutas de abajo.

SALIDA
    R1_<nombre del archivo>.xlsx
    R2_<nombre del archivo>.xlsx
    y su equivalente en CSV.
================================================================================
"""

import os
import sys
import glob
from datetime import datetime
import pandas as pd

ENCODING = "latin-1"

RUTA_ENTRADA = r"D:\ESCRITORIO POR D\SEC_PLANEACION\SEPTIEMBRE\CATASTRO\ACTUALIZACION\RESOLUCIONES\R1_R2\Entrada"
RUTA_SALIDA = r"D:\ESCRITORIO POR D\SEC_PLANEACION\SEPTIEMBRE\CATASTRO\ACTUALIZACION\RESOLUCIONES\R1_R2\Salida"

# =============================================================================
# ESTRUCTURA EN EL ORDEN EXACTO DE LOS EXCEL DE REFERENCIA
# (nombre de columna, posicion inicial, posicion final) 1-based inclusivas
# =============================================================================

CABECERA = [
    ("Departamento",   1,  2),
    ("Municipio",      3,  5),
    ("NoPredial",      6, 30),
    ("TipoRegistro",  31, 31),
    ("NoOrden",       32, 34),
    ("TotalRegistro", 35, 37),
]

# --- R1: 312 caracteres, 18 columnas ---------------------------------------
COLUMNAS_R1 = CABECERA + [
    ("Nombre",               38, 137),
    ("EstadoCivil",         138, 138),
    ("TipoDocumento",       139, 139),
    ("NoDocumento",         140, 151),
    ("Direccion",           152, 251),
    ("Comuna",              252, 252),
    ("DestinoEconomico",    253, 253),
    ("AreaTerreno (m2)",    254, 268),
    ("AreaConstruida (m2)", 269, 274),
    ("Avaluo ($)",          275, 289),
    ("Vigencia",            290, 297),
    ("NoPredialAnterior",   298, 312),
]


def _bloque_construccion(ini):
    """Los tres bloques de construccion del R2 son identicos y estan separados
    por 48 posiciones entre si (26 de datos mas 22 de espacio)."""
    return [
        ("Habitaciones",         ini,      ini + 3),
        ("Baños",                ini + 4,  ini + 7),
        ("Locales",              ini + 8,  ini + 11),
        ("Pisos",                ini + 12, ini + 13),
        ("Estrato",              ini + 14, ini + 14),
        ("Uso",                  ini + 15, ini + 17),
        ("Puntaje",              ini + 18, ini + 19),
        ("AreaConstruida (m2)",  ini + 20, ini + 25),
    ]


# --- R2: 330 caracteres, 46 columnas ---------------------------------------
COLUMNAS_R2 = (
    CABECERA
    + [("MatriculaInmobiliaria", 38, 55),
       ("Espacio1",              56, 77),
       ("ZonaFisica",            78, 80),
       ("ZonaEconomica",         81, 83),
       ("AreaTerreno (m2)",      84, 98),
       ("Espacio2",              99, 120),
       ("ZonaFisica",           121, 123),
       ("ZonaEconomica",        124, 126),
       ("AreaTerreno (m2)",     127, 141),
       ("Espacio3",             142, 163)]
    + _bloque_construccion(164)
    + [("Espacio4", 190, 211)]
    + _bloque_construccion(212)
    + [("Espacio5", 238, 259)]
    + _bloque_construccion(260)
    + [("Espacio6",          286, 307),
       ("Vigencia",          308, 315),
       ("NoPredialAnterior", 316, 330)]
)

LARGOS = {312: ("R1", COLUMNAS_R1), 330: ("R2", COLUMNAS_R2)}

# Columnas que el Excel de referencia entrega como NUMERO, sin ceros a la
# izquierda. Las demas quedan como texto para no perder los ceros del numero
# predial, el documento ni la vigencia.
NUMERICAS = {
    "AreaTerreno (m2)", "AreaConstruida (m2)", "Avaluo ($)",
    "Habitaciones", "Baños", "Locales", "Pisos", "Estrato", "Puntaje",
}


def extraer(ruta):
    """Devuelve (tipo, DataFrame) con las columnas en el orden de referencia."""
    nombre = os.path.basename(ruta)
    with open(ruta, encoding=ENCODING) as fh:
        lineas = [l.rstrip("\n\r") for l in fh if l.strip()]
    if not lineas:
        print(f"  {nombre}: vacio, se omite")
        return None, None

    largos = [len(l) for l in lineas]
    largo = max(set(largos), key=largos.count)
    if largo not in LARGOS:
        print(f"  {nombre}: longitud {largo} no reconocida (se esperaba 312 o 330). Se omite.")
        return None, None

    tipo, columnas = LARGOS[largo]
    descartadas = sum(1 for n in largos if n != largo)
    lineas = [l for l in lineas if len(l) == largo]
    print(f"  {nombre}: {len(lineas)} registros -> {tipo}"
          + (f"  ({descartadas} lineas descartadas por longitud)" if descartadas else ""))

    # Se extrae posicion por posicion, sin convertir tipos: transcripcion fiel.
    datos = {}
    for i, (col, a, b) in enumerate(columnas):
        datos[i] = [l[a - 1:b].strip() for l in lineas]
    df = pd.DataFrame(datos)
    df.columns = [c for c, _, _ in columnas]

    # Convertir a numero las columnas que la referencia entrega sin ceros a la
    # izquierda. Se usa la posicion y no el nombre, porque en el R2 hay nombres
    # repetidos (tres bloques de construccion con los mismos encabezados).
    for i, (col, _, _) in enumerate(columnas):
        if col in NUMERICAS:
            serie = pd.to_numeric(
                df.iloc[:, i].str.replace(",", ".", regex=False), errors="coerce")
            df.isetitem(i, serie)
    return tipo, df


def main():
    if len(sys.argv) >= 3:
        entrada, salida_base = sys.argv[1], sys.argv[2]
    else:
        entrada, salida_base = RUTA_ENTRADA, RUTA_SALIDA
        print("Sin argumentos: usando las rutas por defecto del script.")
    print(f"  Entrada: {entrada}\n  Salida : {salida_base}\n")

    if not os.path.isdir(entrada):
        print(f"ERROR: la carpeta de entrada no existe:\n  {entrada}")
        sys.exit(1)

    salida = os.path.join(salida_base, "referencia_" + datetime.now().strftime("%Y%m%d_%H%M"))
    os.makedirs(salida, exist_ok=True)

    rutas = []
    for pat in ("*.TXT", "*.txt", "*.CSV", "*.csv"):
        rutas += glob.glob(os.path.join(entrada, pat))
    rutas = sorted(set(rutas))
    if not rutas:
        print(f"No se encontraron archivos en {entrada}")
        sys.exit(1)

    generados = 0
    for ruta in rutas:
        tipo, df = extraer(ruta)
        if df is None:
            continue
        base = os.path.splitext(os.path.basename(ruta))[0]
        nombre = f"{tipo}_{base}"

        # Excel: todo como texto, para no perder ceros a la izquierda ni que
        # Excel convierta el numero predial a notacion cientifica.
        destino = os.path.join(salida, nombre + ".xlsx")
        from openpyxl.styles import Font
        with pd.ExcelWriter(destino, engine="openpyxl") as xw:
            df.to_excel(xw, sheet_name=nombre[:31], index=False)
            hoja = xw.sheets[nombre[:31]]
            for celda in hoja[1]:
                celda.font = Font(bold=True)
            hoja.freeze_panes = "A2"

        df.to_csv(os.path.join(salida, nombre + ".csv"),
                  index=False, encoding="utf-8-sig", sep=";")

        print(f"     -> {nombre}.xlsx  ({len(df)} filas, {len(df.columns)} columnas)")
        generados += 1

    print("\n" + "=" * 62)
    print(f"  Archivos generados : {generados}")
    print(f"  Salida             : {os.path.abspath(salida)}")
    print("=" * 62)


if __name__ == "__main__":
    main()
