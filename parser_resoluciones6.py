#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
 PARSER DE ARCHIVOS PLANOS DE RESOLUCIONES DE CONSERVACION CATASTRAL
 Convenio IGAC - Municipio de Mosquera (Cundinamarca)
 Secretaria de Planeacion - Supervision del convenio
================================================================================
"""

import sys
import io
import glob
import os
import unicodedata
import hashlib
from datetime import datetime
import pandas as pd

ENCODING = "latin-1"

RUTA_ENTRADA = r"D:\ESCRITORIO POR D\SEC_PLANEACION\SEPTIEMBRE\CATASTRO\ACTUALIZACION\RESOLUCIONES\SCRIPTS\Entrada"
RUTA_SALIDA  = r"D:\ESCRITORIO POR D\SEC_PLANEACION\SEPTIEMBRE\CATASTRO\ACTUALIZACION\RESOLUCIONES\SCRIPTS\Salida"

# --- Cabecera comun ---
SPEC_COMUN = [
    ("departamento",       1,   2),
    ("municipio",          3,   5),
    ("resolucion_nro",     6,  14),
    ("resolucion_anio",   15,  18),
    ("radicacion_nro",    19,  29),
    ("radicacion_anio",   30,  33),
    ("tipo_tramite",      34,  35),
    ("clase_mutacion",    36,  36),
    ("npn_zona",          37,  38),
    ("npn_sector",        39,  40),
    ("npn_comuna",        41,  42),
    ("npn_barrio",        43,  44),
    ("npn_manzana",       45,  48),
    ("npn_terreno",       49,  52),
    ("npn_condicion",     53,  53),
    ("npn_edificio",      54,  55),
    ("npn_piso",          56,  57),
    ("npn_unidad",        58,  61),
    ("cancela_inscribe",  62,  62),
    ("tipo_registro",     63,  63),
    ("numero_orden",      64,  66),
    ("total_registros",   67,  69),
]

SPEC_T1 = [
    ("nombre",                   70, 169),
    ("estado_civil",            170, 170),
    ("tipo_documento",          171, 171),
    ("numero_documento",        172, 183),
    ("direccion",               184, 283),
    ("comuna",                  284, 284),
    ("destino_economico",       285, 285),
    ("area_terreno_raw",        286, 300),
    ("area_construida_raw",     301, 306),
    ("avaluo_raw",              307, 321),
    ("vigencia_raw",            322, 329),
    ("numero_predial_anterior", 330, 344),
]

SPEC_T2 = [
    ("matricula",                70,  87),
    ("zona_fisica_1",           121, 123),
    ("zona_economica_1",        124, 126),
    ("area_terreno_1",          127, 141),
    ("zona_fisica_2",           175, 177),
    ("zona_economica_2",        178, 180),
    ("area_terreno_2",          181, 195),
    ("habitaciones_1",          229, 230),
    ("banos_1",                 231, 232),
    ("locales_1",               233, 234),
    ("pisos_1",                 235, 236),
    ("tipificacion_1",          237, 238),
    ("uso_1",                   239, 241),
    ("puntaje_1",               242, 243),
    ("area_construida_1",       244, 249),
    ("habitaciones_2",          283, 284),
    ("banos_2",                 285, 286),
    ("locales_2",               287, 288),
    ("pisos_2",                 289, 290),
    ("tipificacion_2",          291, 292),
    ("uso_2",                   293, 295),
    ("puntaje_2",               296, 297),
    ("area_construida_2",       298, 303),
    ("habitaciones_3",          337, 338),
    ("banos_3",                 339, 340),
    ("locales_3",               341, 342),
    ("pisos_3",                 343, 344),
    ("tipificacion_3",          345, 346),
    ("uso_3",                   347, 349),
    ("puntaje_3",               350, 351),
    ("area_construida_3",       352, 357),
    ("numero_predial_anterior", 414, 428),
]

SPEC_T3 = [
    ("decretos",                 70, 139),
    ("motivacion",              140, 395),
    ("numero_predial_anterior", 396, 410),
]

MOVIMIENTO = {"C": "Cancelacion", "I": "Inscripcion"}
TIPO_REGISTRO = {"1": "Predio y propietario", "2": "Registral y construcciones", "3": "Acto administrativo"}
CLASE_MUTACION_DOM = {
    "1": "Primera clase - cambio de propietario, poseedor u ocupante",
    "2": "Segunda clase - linderos, agregacion o segregacion (incluye PH)",
    "3": "Tercera clase - construcciones, uso y destino economico",
    "4": "Cuarta clase - reajuste, revision de avaluo o autoestimacion",
    "5": "Quinta clase - inscripcion de predios no registrados",
}
TIPO_DOCUMENTO = {"C": "Cedula de ciudadania", "N": "NIT"}
ZONA = {"00": "Rural", "01": "Urbana", "04": "Centro poblado"}

def estructura_npn(fila):
    torre = fila["npn_edificio"] != "00"
    piso = fila["npn_piso"] != "00"
    unidad = fila["npn_unidad"] != "0000"
    if torre and piso and unidad: return "Torre, piso y unidad"
    if unidad and not torre and not piso: return "Solo unidad"
    if not torre and not piso and not unidad: return "Sin torre, piso ni unidad"
    return "Combinacion atipica"

def colspecs(spec):
    return [(a - 1, b) for _, a, b in spec], [n for n, _, _ in spec]

def leer_bloque(lineas, spec):
    if not lineas: return pd.DataFrame(columns=[n for n, _, _ in spec])
    cs, nombres = colspecs(spec)
    return pd.read_fwf(io.StringIO("\n".join(lineas)), colspecs=cs, names=nombres, dtype=str, header=None, keep_default_na=False)

def limpiar(df):
    for c in df.columns: df[c] = df[c].astype(str).str.strip()
    return df

def a_entero(serie):
    return pd.to_numeric(serie.str.strip(), errors="coerce").astype("Int64")

def normaliza_doc(v):
    if not isinstance(v, str): return ""
    return v.replace("-", "").replace(" ", "").lstrip("0")

def normaliza_texto(v):
    if not isinstance(v, str): return ""
    v = unicodedata.normalize("NFKD", v).encode("ascii", "ignore").decode()
    return " ".join(v.upper().split())

def a_fecha(serie):
    return pd.to_datetime(serie, format="%d%m%Y", errors="coerce")

def procesar_archivo(ruta):
    nombre = os.path.basename(ruta)
    with open(ruta, "rb") as fh:
        huella = hashlib.md5(fh.read()).hexdigest()[:10]
    with open(ruta, encoding=ENCODING) as fh:
        crudas = [l.rstrip("\n\r") for l in fh]
    lineas = [l for l in crudas if l.strip()]

    incidencias = []
    if not lineas:
        incidencias.append({"archivo": nombre, "tipo": "Archivo vacio", "detalle": "No se encontraron registros", "npn": ""})
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(incidencias)

    if any(sep in lineas[0] for sep in [";", ","]) and len(lineas[0]) < 400:
        incidencias.append({"archivo": nombre, "tipo": "Formato delimitado no soportado", "detalle": "El archivo tiene delimitadores.", "npn": ""})
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(incidencias)

    largos_esperados = {410, 428}
    malas = [(i, len(l)) for i, l in enumerate(lineas, 1) if len(l) not in largos_esperados]
    for nro, largo in malas:
        incidencias.append({"archivo": nombre, "tipo": "Longitud de linea inesperada", "detalle": f"linea {nro}: {largo} caracteres", "npn": ""})
    lineas = [l for l in lineas if len(l) in largos_esperados]

    por_tipo = {}
    for l in lineas:
        por_tipo.setdefault(l[62:63], []).append(l)

    def armar(tipos, spec_cola):
        lin = [l for t in tipos for l in por_tipo.get(t, [])]
        if not lin: return pd.DataFrame()
        df = pd.concat([limpiar(leer_bloque(lin, SPEC_COMUN)), limpiar(leer_bloque(lin, spec_cola))], axis=1)
        df.insert(0, "archivo", nombre)
        df.insert(1, "huella_archivo", huella)
        return df

    t1, t2, t3 = armar(["1"], SPEC_T1), armar(["2"], SPEC_T2), armar(["3"], SPEC_T3)

    for df in (t1, t2, t3):
        if df.empty: continue
        df["npn"] = (df["departamento"] + df["municipio"] + df["npn_zona"] + df["npn_sector"]
                     + df["npn_comuna"] + df["npn_barrio"] + df["npn_manzana"]
                     + df["npn_terreno"] + df["npn_condicion"]
                     + df["npn_edificio"] + df["npn_piso"] + df["npn_unidad"])
        df["movimiento"] = df["cancela_inscribe"].map(MOVIMIENTO)
        df["tipo_registro_desc"] = df["tipo_registro"].map(TIPO_REGISTRO)
        df["clase_mutacion_desc"] = df["clase_mutacion"].map(CLASE_MUTACION_DOM).fillna("Sin descripcion: " + df["clase_mutacion"])
        df["zona"] = df["npn_zona"].map(ZONA).fillna("Por confirmar: " + df["npn_zona"])
        df["estructura_npn"] = df.apply(estructura_npn, axis=1)
        df["resolucion"] = df["resolucion_nro"].str.lstrip("0").replace("", "0") + " de " + df["resolucion_anio"]

    if not t1.empty:
        t1["area_terreno_m2"] = a_entero(t1["area_terreno_raw"])
        t1["area_construida_m2"] = a_entero(t1["area_construida_raw"])
        t1["avaluo_pesos"] = a_entero(t1["avaluo_raw"])
        t1["vigencia"] = a_fecha(t1["vigencia_raw"])
        t1["tipo_documento_desc"] = t1["tipo_documento"].map(TIPO_DOCUMENTO).fillna("Por confirmar")
        t1["documento_norm"] = t1["numero_documento"].map(normaliza_doc)
        t1["nombre_norm"] = t1["nombre"].map(normaliza_texto)
        t1["numero_predial_anterior"] = t1["numero_predial_anterior"].str.strip()

    if not t2.empty:
        for c in ("area_terreno_1", "area_terreno_2"):
            t2[c + "_m2"] = pd.to_numeric(t2[c].str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce")
        t2["area_m2_registral"] = t2["area_terreno_1_m2"]
        for c in ("area_construida_1", "area_construida_2", "area_construida_3"):
            t2[c + "_m2"] = a_entero(t2[c])
        t2["orip"] = t2["matricula"].str.split("-").str[0].str.strip()

    if not t3.empty:
        t3["concepto"] = t3["decretos"].str.split("$").str[0].str.strip()
        t3["avaluo_texto"] = pd.to_numeric(t3["decretos"].str.extract(r"\$([\d,]+)")[0].str.replace(",", "", regex=False), errors="coerce").astype("Int64")
        t3["vigencia_fiscal"] = pd.to_datetime(t3["decretos"].str.extract(r"(\d{2}/\d{2}/\d{4})")[0], format="%d/%m/%Y", errors="coerce")

    return t1, t2, t3, pd.DataFrame(incidencias)

def validar_duplicidad(t1, t2, t3, rutas):
    inc = []
    partes = [d for d in (t1, t2, t3) if not d.empty]
    if not partes: return pd.DataFrame(inc)
    todo = pd.concat(partes, ignore_index=True)

    # 1. Archivos idénticos
    hu = todo[["archivo", "huella_archivo"]].drop_duplicates()
    for h, g in hu.groupby("huella_archivo"):
        if len(g) > 1:
            inc.append({
                "nivel": "1. Archivo duplicado",
                "tipo": "Archivo duplicado real",
                "detalle": " | ".join(sorted(g["archivo"])),
                "accion": "Dejar un solo archivo en Entrada",
                "npn": ""
            })

    # 2. Resolución idéntica entre archivos distintos
    def huella_resolucion(df1, df2, df3, resolucion):
        partes_res = []
        for nombre_tipo, df in (("T1", df1), ("T2", df2), ("T3", df3)):
            if df.empty: continue
            r = df[df["resolucion"] == resolucion].copy()
            if r.empty: continue
            cols = [c for c in r.columns if c not in ("archivo", "huella_archivo")]
            r = r[cols].astype(str)
            r = r.sort_values(cols, kind="mergesort").reset_index(drop=True)
            partes_res.append(nombre_tipo + "\n" + r.to_csv(index=False, lineterminator="\n"))
        if not partes_res: return ""
        return hashlib.md5("\n---TIPO---\n".join(partes_res).encode("utf-8")).hexdigest()

    resoluciones_archivo = todo[["resolucion", "archivo"]].drop_duplicates()
    datos_resoluciones = []
    for _, fila in resoluciones_archivo.iterrows():
        resolucion, archivo = fila["resolucion"], fila["archivo"]
        huella = huella_resolucion(t1[t1["archivo"] == archivo], t2[t2["archivo"] == archivo], t3[t3["archivo"] == archivo], resolucion)
        datos_resoluciones.append({"resolucion": resolucion, "archivo": archivo, "huella_resolucion": huella})

    rr = pd.DataFrame(datos_resoluciones)
    if not rr.empty:
        for (resolucion, huella), g in rr.groupby(["resolucion", "huella_resolucion"]):
            if len(g) > 1:
                inc.append({
                    "nivel": "2. Resolución idéntica entre archivos",
                    "tipo": f"La resolución {resolucion} es idéntica en varios archivos",
                    "detalle": " | ".join(sorted(g["archivo"].unique())),
                    "accion": "Conservar una sola versión",
                    "npn": ""
                })

    # 3. Informativo
    pr = todo[["npn", "resolucion"]].drop_duplicates()
    multi = pr.groupby("npn")["resolucion"].agg(list)
    for npn, lista in multi[multi.str.len() > 1].items():
        inc.append({
            "nivel": "3. Informativo",
            "tipo": "Predio afectado por varias resoluciones del lote",
            "detalle": " | ".join(sorted(lista)),
            "accion": "Informativo. No constituye inconsistencia",
            "npn": npn
        })

    return pd.DataFrame(inc)

def deduplicar_resoluciones(t1, t2, t3):
    """Deja una sola copia de cada resolucion antes de consolidar.

    Cuando la misma resolucion llega en dos archivos (por ejemplo un reenvio o
    una entrega parcial que repite lo ya entregado), el consolidado la contaria
    dos veces e inflaria predios y avaluos. Aqui se elige un archivo por
    resolucion y se descartan las filas provenientes de los demas.

    Criterio de eleccion: el archivo que aporte MAS registros para esa
    resolucion, por ser el mas completo. En caso de empate, el primero en orden
    alfabetico. La decision queda registrada en las alertas.

    Antes de descartar se compara el contenido. Si las copias son identicas, el
    descarte es seguro. Si difieren, se descarta igual para no inflar las cifras
    pero se marca como CRITICO, porque significa que hay dos versiones
    distintas de la misma resolucion y alguien debe decidir cual es la buena.
    """
    alertas = []
    partes = [d for d in (t1, t2, t3) if not d.empty]
    if not partes:
        return t1, t2, t3, pd.DataFrame(alertas)

    todo = pd.concat(partes, ignore_index=True)

    # Firma del contenido de cada (archivo, resolucion): numero de registros y
    # huella de la lista ordenada de predios con su lado y numero de orden.
    def firma(g):
        clave = (g["npn"] + "|" + g["cancela_inscribe"] + "|"
                 + g["tipo_registro"] + "|" + g["numero_orden"])
        return pd.Series({
            "registros": len(g),
            "huella": hashlib.md5("".join(sorted(clave)).encode()).hexdigest()[:10],
        })

    fir = todo.groupby(["resolucion", "archivo"]).apply(
        firma, include_groups=False).reset_index()

    descartar = set()
    for res, g in fir.groupby("resolucion"):
        if len(g) == 1:
            continue
        g = g.sort_values(["registros", "archivo"], ascending=[False, True])
        elegido = g.iloc[0]["archivo"]
        sobrantes = list(g.iloc[1:]["archivo"])
        identicas = g["huella"].nunique() == 1
        for a in sobrantes:
            descartar.add((res, a))
        alertas.append({
            "nivel": ("2. Resolución idéntica entre archivos" if identicas
                      else "2. CRITICO: versiones distintas"),
            "tipo": (f"La resolucion {res} viene en {len(g)} archivos"
                     + ("" if identicas else " CON CONTENIDO DIFERENTE")),
            "detalle": ("Se conservo: " + elegido
                        + " (" + str(int(g.iloc[0]["registros"])) + " registros). "
                        + "Se descarto para el consolidado: " + " | ".join(sobrantes)),
            "accion": ("Retirar de Entrada el archivo sobrante." if identicas else
                       "REVISAR: hay dos versiones distintas de la misma "
                       "resolucion. Confirmar con el IGAC cual es la vigente y "
                       "mover la otra a Reemplazados."),
            "npn": "",
        })

    if descartar:
        def filtrar(df):
            if df.empty:
                return df
            m = pd.Series(list(zip(df["resolucion"], df["archivo"])),
                          index=df.index).isin(descartar)
            return df[~m].reset_index(drop=True)
        t1, t2, t3 = filtrar(t1), filtrar(t2), filtrar(t3)

    return t1, t2, t3, pd.DataFrame(alertas)


def consolidar(t1, t2, t3):
    """Un renglon por predio y resolucion: como estaba, como queda y con que
    soporte registral, documental y el predial anterior."""
    if t1.empty: return pd.DataFrame()
    idx = ["archivo", "resolucion", "npn", "zona", "estructura_npn"]
    
    # --- lado tipo 1: area, avaluo, propietarios, titular y predial anterior ---
    base = t1.drop_duplicates(subset=idx + ["movimiento", "numero_orden"])
    agr = (base.groupby(idx + ["movimiento"], as_index=False)
                .agg(area_m2=("area_terreno_m2", "first"),
                     avaluo_pesos=("avaluo_pesos", "first"),
                     propietarios=("numero_orden", "nunique"),
                     direccion=("direccion", "first"),
                     area_construida=("area_construida_m2", "first"),
                     destino=("destino_economico", "first"),
                     predial_anterior=("numero_predial_anterior", "first"))) # <- CAPTURA EL PREDIAL ANTERIOR
                     
    tit = (t1[t1["numero_orden"] == "001"]
           .drop_duplicates(subset=idx + ["movimiento"])
           .set_index(idx + ["movimiento"])[["nombre", "documento_norm", "nombre_norm"]]
           .rename(columns={"nombre": "titular", "documento_norm": "titular_doc", "nombre_norm": "titular_norm"}))
    agr = agr.join(tit, on=idx + ["movimiento"])

    piv = agr.pivot_table(index=idx, columns="movimiento",
                          values=["area_m2", "avaluo_pesos", "propietarios", "titular", 
                                  "titular_doc", "titular_norm", "direccion", "destino", 
                                  "area_construida", "predial_anterior"], # <- INCLUIDO EN LA TABLA PIVOTE
                          aggfunc="first")
    piv.columns = [f"{a}_{b.lower()}" for a, b in piv.columns]
    piv = piv.reset_index()

    # Asegurar columnas base por si alguna viene vacía
    columnas_requeridas = [
        "area_m2_cancelacion", "area_m2_inscripcion",
        "avaluo_pesos_cancelacion", "avaluo_pesos_inscripcion",
        "propietarios_cancelacion", "propietarios_inscripcion",
        "titular_cancelacion", "titular_inscripcion",
        "titular_doc_cancelacion", "titular_doc_inscripcion",
        "direccion_cancelacion", "direccion_inscripcion",
        "destino_cancelacion", "destino_inscripcion",
        "titular_norm_cancelacion", "titular_norm_inscripcion",
        "predial_anterior_cancelacion", "predial_anterior_inscripcion" # <- GARANTIZA LAS COLUMNAS
    ]
    for c in columnas_requeridas:
        if c not in piv.columns: piv[c] = pd.NA

    if not t2.empty:
        reg = (t2[t2["movimiento"] == "Inscripcion"]
               .drop_duplicates(subset=["archivo", "resolucion", "npn"])
               .set_index(["archivo", "resolucion", "npn"])
               [["matricula", "orip", "area_m2_registral", "zona_fisica_1", "zona_economica_1"]])
        piv = piv.join(reg, on=["archivo", "resolucion", "npn"])

    if not t3.empty:
        act = (t3.sort_values("vigencia_fiscal")
               .drop_duplicates(subset=["archivo", "resolucion", "npn"], keep="last")
               .set_index(["archivo", "resolucion", "npn"])
               [["concepto", "avaluo_texto", "vigencia_fiscal"]])
        piv = piv.join(act, on=["archivo", "resolucion", "npn"])

    # --- clase de mutacion y tipo de tramite -------------------------------
    # Son constantes por resolucion. Se traen aqui para que el consolidado
    # responda que clase de mutacion afecto al predio sin abrir otra pestana.
    mut = (t1.drop_duplicates(subset=["archivo", "resolucion"])
             .set_index(["archivo", "resolucion"])
             [["clase_mutacion", "clase_mutacion_desc", "tipo_tramite"]])
    piv = piv.join(mut, on=["archivo", "resolucion"])

    # --- variaciones --------------------------------------------------------
    piv["variacion_avaluo"] = (piv["avaluo_pesos_inscripcion"].fillna(0)
                               - piv["avaluo_pesos_cancelacion"].fillna(0))
    piv["variacion_area"] = (piv["area_m2_inscripcion"].fillna(0)
                             - piv["area_m2_cancelacion"].fillna(0))

    # --- que cambio realmente entre cancelacion e inscripcion ---------------
    piv["cambia_titular"] = (piv["titular_doc_cancelacion"].notna()
                             & piv["titular_doc_inscripcion"].notna()
                             & (piv["titular_doc_cancelacion"] != piv["titular_doc_inscripcion"]))
    piv["corrige_nombre"] = (piv["titular_norm_cancelacion"].notna()
                             & piv["titular_norm_inscripcion"].notna()
                             & (piv["titular_norm_cancelacion"] != piv["titular_norm_inscripcion"])
                             & (piv["titular_doc_cancelacion"] == piv["titular_doc_inscripcion"]))
    piv["cambia_destino"] = (piv["destino_cancelacion"].notna()
                             & piv["destino_inscripcion"].notna()
                             & (piv["destino_cancelacion"] != piv["destino_inscripcion"]))

    def situacion(r):
        if pd.isna(r["avaluo_pesos_cancelacion"]):
            return "Solo inscripcion"
        if pd.isna(r["avaluo_pesos_inscripcion"]):
            return "Solo cancelacion"
        return "Cancelacion e inscripcion"

    def observado(r):
        """Lo que efectivamente cambia en el dato. Es lectura de los campos,
        NO la clasificacion oficial: esa va en clase_mutacion_desc."""
        s = r["situacion"]
        if s == "Solo inscripcion":
            return "NPN nuevo en la resolucion"
        if s == "Solo cancelacion":
            return "NPN retirado de la base"
        cambios = []
        if r["cambia_titular"]:
            cambios.append("titular")
        if r["variacion_area"] != 0:
            cambios.append("area")
        if r["variacion_avaluo"] != 0:
            cambios.append("avaluo")
        if r["cambia_destino"]:
            cambios.append("destino")
        if r["corrige_nombre"]:
            cambios.append("escritura del nombre")
        if not cambios:
            return "Mismo NPN sin cambios detectables"
        if cambios == ["destino"]:
            return "Reclasificacion de destino, sin cambio patrimonial"
        if cambios == ["escritura del nombre"]:
            return "Correccion del nombre del titular, mismo documento"
        return "Mismo NPN, cambia " + " y ".join(cambios)

    piv["situacion"] = piv.apply(situacion, axis=1)
    piv["movimiento_observado"] = piv.apply(observado, axis=1)

    # --- orden de columnas: identificacion, clasificacion, luego el detalle --
    orden = (idx
             + ["clase_mutacion", "clase_mutacion_desc", "tipo_tramite",
                "situacion", "movimiento_observado",
                "area_m2_cancelacion", "area_m2_inscripcion", "variacion_area",
                "avaluo_pesos_cancelacion", "avaluo_pesos_inscripcion", "variacion_avaluo",
                "area_construida_cancelacion", "area_construida_inscripcion",
                "propietarios_cancelacion", "propietarios_inscripcion",
                "titular_cancelacion", "titular_inscripcion",
                "cambia_titular", "corrige_nombre",
                "destino_cancelacion", "destino_inscripcion", "cambia_destino",
                "titular_doc_cancelacion", "titular_doc_inscripcion",
                "direccion_inscripcion", "direccion_cancelacion",
                "predial_anterior_inscripcion", "predial_anterior_cancelacion"]
             + [c for c in ("matricula", "orip", "area_m2_registral",
                            "zona_fisica_1", "zona_economica_1",
                            "concepto", "avaluo_texto", "vigencia_fiscal")
                if c in piv.columns])
    return piv[[c for c in orden if c in piv.columns]]

def resumen_resoluciones(t1, t2, t3):
    partes = [d for d in (t1, t2, t3) if not d.empty]
    if not partes: return pd.DataFrame()
    todo = pd.concat(partes, ignore_index=True)
    g = todo.groupby(["archivo", "resolucion", "resolucion_nro", "resolucion_anio", "radicacion_nro", "tipo_tramite", "clase_mutacion", "clase_mutacion_desc"], as_index=False).agg(
        registros=("npn", "size"), predios=("npn", "nunique"))
    if not t1.empty:
        ins = (t1[t1["movimiento"] == "Inscripcion"]
               .drop_duplicates(subset=["archivo", "resolucion", "npn"])
               .groupby(["archivo", "resolucion"], as_index=False)
               .agg(predios_inscritos=("npn", "nunique"), avaluo_inscrito=("avaluo_pesos", "sum")))
        can = (t1[t1["movimiento"] == "Cancelacion"]
               .drop_duplicates(subset=["archivo", "resolucion", "npn"])
               .groupby(["archivo", "resolucion"], as_index=False)
               .agg(predios_cancelados=("npn", "nunique"), avaluo_cancelado=("avaluo_pesos", "sum")))
        g = g.merge(ins, on=["archivo", "resolucion"], how="left").merge(can, on=["archivo", "resolucion"], how="left")
    return g

def main():
    entrada, salida_base = (sys.argv[1], sys.argv[2]) if len(sys.argv) >= 3 else (RUTA_ENTRADA, RUTA_SALIDA)
    if not os.path.isdir(entrada): sys.exit(1)
    
    salida = os.path.join(salida_base, "procesamiento_" + datetime.now().strftime("%Y%m%d_%H%M"))
    os.makedirs(salida, exist_ok=True)

    rutas = sorted(glob.glob(os.path.join(entrada, "*.TXT")) +
                   glob.glob(os.path.join(entrada, "*.txt")) +
                   glob.glob(os.path.join(entrada, "*.CSV")) +
                   glob.glob(os.path.join(entrada, "*.csv")))
    if not rutas: sys.exit(1)

    # --- ORDEN CRONOLÓGICO: Archivos más antiguos primero, más recientes al final ---
    rutas = sorted(rutas, key=os.path.getmtime)

    T1, T2, T3, INC = [], [], [], []
    for r in rutas:
        print(f"  procesando {os.path.basename(r)} ...")
        a, b, c, d = procesar_archivo(r)
        for lista, df in ((T1, a), (T2, b), (T3, c), (INC, d)):
            if not df.empty: lista.append(df)

    def unir(lista):
        return pd.concat(lista, ignore_index=True) if lista else pd.DataFrame()

    t1, t2, t3, inc = unir(T1), unir(T2), unir(T3), unir(INC)

    # --- POLÍTICA DE VERSIÓN AUTOMÁTICA (UPSET POR ORDEN DE ARCHIVO) ---
    # Llave de unicidad de un registro catastral
    llave_version = ["resolucion", "npn", "cancela_inscribe", "tipo_registro", "numero_orden"]
    
    for df_name in ['t1', 't2', 't3']:
        df_obj = locals()[df_name]
        if not df_obj.empty:
            # keep='last' asegura que si la misma resolución/predio viene en un archivo más nuevo, 
            # se sobrescribe el dato viejo y se queda con la versión fresca.
            locals()[df_name] = df_obj.drop_duplicates(subset=llave_version, keep='last')

    t1, t2, t3 = locals()['t1'], locals()['t2'], locals()['t3']

    # El resumen (01) se construye ANTES de depurar: es el inventario de TODO
    # lo que se recibio y proceso, incluidas las copias repetidas.
    res_inventario = resumen_resoluciones(t1, t2, t3)
    if not res_inventario.empty:
        res_inventario = res_inventario.copy()

    # Depurar resoluciones repetidas antes de generar las demas salidas, para
    # que t1, t2, t3 y el consolidado no queden inflados. La depuracion afecta
    # a las cuatro por igual: la coherencia entre pestanas es lo que importa.
    n_antes = (len(t1), len(t2), len(t3))
    t1, t2, t3, ded = deduplicar_resoluciones(t1, t2, t3)
    if not ded.empty:
        inc = pd.concat([inc, ded], ignore_index=True)
        print(f"  Depuracion: t1 {n_antes[0]}->{len(t1)}, "
              f"t2 {n_antes[1]}->{len(t2)}, t3 {n_antes[2]}->{len(t3)} registros")

    dup = validar_duplicidad(t1, t2, t3, rutas)
    if not dup.empty:
        inc = pd.concat([inc, dup], ignore_index=True)

    cons = consolidar(t1, t2, t3)

    # Marcar en el inventario cuales copias entraron a las cifras y cuales no.
    if not res_inventario.empty:
        if not t1.empty:
            vigentes = set(zip(t1["archivo"], t1["resolucion"]))
        else:
            vigentes = set()
        res_inventario["usada_en_cifras"] = [
            "Si" if (a, r) in vigentes else "No - copia descartada"
            for a, r in zip(res_inventario["archivo"], res_inventario["resolucion"])]
    res = res_inventario

    tablas = {
        "01_resoluciones_resumen": res,
        "02_t1_propietarios_avaluos": t1,
        "03_t2_matriculas_fisico": t2,
        "04_t3_acto_administrativo": t3,
        "05_consolidado": cons,
        "06_alertas": inc,
    }

    for nombre, df in tablas.items():
        df.to_csv(os.path.join(salida, f"{nombre}.csv"), index=False, encoding="utf-8-sig", sep=";", decimal=",")

    with pd.ExcelWriter(os.path.join(salida, "resoluciones.xlsx"), engine="openpyxl") as xw:
        for nombre, df in tablas.items():
            if df.empty: continue
            df.head(100000).to_excel(xw, sheet_name=nombre[:31], index=False)

    print("\n" + "=" * 62)
    print(f"  Archivos procesados : {len(rutas)}")
    print(f"  Resoluciones        : {len(res) if not res.empty else 0}")
    print(f"  Predios distintos   : {t1['npn'].nunique() if not t1.empty else 0}")
    print(f"  Salida              : {os.path.abspath(salida)}")
    print("=" * 62)

if __name__ == "__main__":
    main()