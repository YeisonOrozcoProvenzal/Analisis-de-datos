# -*- coding: utf-8 -*-
"""Analisis de movimiento: segmentacion actual de Provenzal vs modelo v3 (3V y 4V).
Lee los resultados que deja run_v3_3V_4V.py en resultados_2019_con_unico/_work/ y
escribe resultados_2019_con_unico/Analisis_Movimiento_v3.xlsx."""
import json
import os
import pathlib
import time
import urllib.parse

import numpy as np
import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine

HERE = pathlib.Path(__file__).resolve().parent
os.chdir(HERE)
load_dotenv()

OUT_DIR = HERE / "resultados_2019_con_unico"
WORK = OUT_DIR / "_work"
OUT_XLSX = OUT_DIR / "Analisis_Movimiento_v3.xlsx"
XLSX_V3_ACTUAL = OUT_DIR / "segmentacion_2019_con_unico.xlsx"  # "Excel v3 actual" (referencia)

FECHA_CORTE = pd.Timestamp("2026-08-31")
FECHA_INICIO = pd.Timestamp("2019-01-01")
CLIENTES_EXCLUIR = [
    "C222222222222", "C22222222222", "CN", "C1",
    "C111", "C1111111111", "C77", "C777", "C000000000",
    "C12235456", "C123456", "C1234567", "C900790149",
]
SKUS_EXCLUIR = ["2069"]
CANALES_EXCLUIR = ["Shopify", "CHATCENTER WEB"]

ORDEN_SEG_PROV = ["Diamante", "Oro", "Plata", "Bronce", "Sin Segmento"]
RANGO_SEG = {s: i for i, s in enumerate(["Sin Segmento", "Bronce", "Plata", "Oro", "Diamante"])}  # mayor = mejor
SIN_REGISTRO = "Sin registro en Provenzal"
FUERA = "Fuera del modelo"
ORDEN_FREC = ["Alta", "Media", "Baja", "Único"]
ORDEN_ESTADO = ["Activo", "Por inactivar", "Inactivo"]

# ---------------------------------------------------------------- conexion
params = urllib.parse.quote_plus(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    f"SERVER={os.getenv('DB_SERVER')};DATABASE=Ventas_Comerssia;"
    f"UID={os.getenv('DB_USER')};PWD={os.getenv('DB_PASSWORD')};"
)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}", pool_pre_ping=True, pool_recycle=180)


def leer(query, intentos=4, espera=10):
    for i in range(1, intentos + 1):
        try:
            return pd.read_sql(query, engine)
        except Exception as e:
            engine.dispose()
            if i == intentos:
                raise
            print(f"intento {i} fallo ({type(e).__name__}); reintento en {espera}s")
            time.sleep(espera)


checks = []  # (chequeo, esperado, observado, estado, nota)


def chk(nombre, esperado, observado, ok, nota=""):
    checks.append({"Chequeo": nombre, "Esperado": esperado, "Observado": observado,
                   "Estado": "OK" if ok else "REVISAR", "Nota": nota})
    print(("OK      " if ok else "REVISAR ") + nombre, "|", observado, "|", nota, flush=True)


# ---------------------------------------------------------------- modelo 3V / 4V
cl3 = pd.read_pickle(WORK / "clientes_3V.pkl")
cl4 = pd.read_pickle(WORK / "clientes_4V.pkl")
rf3 = pd.read_pickle(WORK / "resultado_final_3V.pkl")
excl_neg = pd.read_pickle(WORK / "excluidos_margen_negativo_3V.pkl")

assert cl3["ID_Cliente_Unico"].is_unique and cl4["ID_Cliente_Unico"].is_unique
modelo = cl3[["ID_Cliente_Unico", "Metal", "GrupoEstructural", "Estado", "NivelFrecuencia", "Valor", "Frecuencia", "n_compras"]].rename(
    columns={"Metal": "Metal_3V", "GrupoEstructural": "GrupoEstructural_3V", "Estado": "Estado_3V", "NivelFrecuencia": "Nivel_3V"})
m4 = cl4[["ID_Cliente_Unico", "Metal", "GrupoEstructural", "Estado", "NivelFrecuencia", "Valor", "Frecuencia", "n_compras"]].rename(
    columns={"Metal": "Metal_4V", "GrupoEstructural": "GrupoEstructural_4V", "Estado": "Estado_4V", "NivelFrecuencia": "Nivel_4V",
             "Valor": "Valor_4V", "Frecuencia": "Frecuencia_4V", "n_compras": "n_compras_4V"})
modelo = modelo.merge(m4, on="ID_Cliente_Unico", how="outer", validate="one_to_one", indicator=True)
assert (modelo["_merge"] == "both").all(), "los universos 3V y 4V no coinciden"
modelo = modelo.drop(columns="_merge")

UNIVERSO = len(modelo)
chk("Universo 3V = 4V = 273.399 clientes", 273399, f"{len(cl3):,} (3V) / {len(cl4):,} (4V) / {UNIVERSO:,} IDs en ambos",
    len(cl3) == len(cl4) == UNIVERSO == 273399)
chk("Valor, Frecuencia y n_compras idénticos entre 3V y 4V", "idénticos",
    "idénticos" if (modelo["Valor"].equals(modelo["Valor_4V"]) and modelo["Frecuencia"].equals(modelo["Frecuencia_4V"]) and modelo["n_compras"].equals(modelo["n_compras_4V"])) else "difieren",
    modelo["Valor"].equals(modelo["Valor_4V"]) and modelo["Frecuencia"].equals(modelo["Frecuencia_4V"]) and modelo["n_compras"].equals(modelo["n_compras_4V"]))
n_dif_nivel = int((modelo["Nivel_3V"] != modelo["Nivel_4V"]).sum())
chk("NivelFrecuencia idéntico por cliente entre 3V y 4V (solo se mueve Metal)", "0 diferencias", f"{n_dif_nivel} diferencias", n_dif_nivel == 0)

# Metal_3V vs cifras del pedido y vs Excel v3 actual
esperado = {"Diamante": 54745, "Oro": 120646, "Plata": 98008}
obs = modelo["Metal_3V"].value_counts().to_dict()
for m, e in esperado.items():
    o = obs.get(m, 0)
    dif = o - e
    chk(f"Metal_3V {m}", e, o, abs(dif) / e <= 0.0005, f"dif={dif:+d} ({dif / e * 100:+.3f}%)")
chk("Metal_3V sin otros niveles", "solo Diamante/Oro/Plata", sorted(obs), set(obs) == set(esperado))

# Contra el Excel v3 actual (Resumen_Estado y Clientes)
v3_res = pd.read_excel(XLSX_V3_ACTUAL, sheet_name="Resumen_Estado").set_index("Estado")["Clientes"]
est3 = modelo["Estado_3V"].value_counts()
for e in ORDEN_ESTADO:
    chk(f"Estado_3V {e} vs Excel v3 actual (Resumen_Estado)", int(v3_res[e]), int(est3.get(e, 0)),
        abs(int(est3.get(e, 0)) - int(v3_res[e])) / int(v3_res[e]) <= 0.0005, f"dif={int(est3.get(e, 0)) - int(v3_res[e]):+d}")
chk("Total Resumen_Estado v3 actual = 273.399", 273399, int(v3_res.sum()), int(v3_res.sum()) == 273399)

# ---------------------------------------------------------------- MASTER_ID y diferencia de 17
master = leer("SELECT Cliente, ID_Cliente_Unico FROM Ventas_Comerssia.dbo.MASTER_ID")
n_master_raw = len(master)
master = master.drop_duplicates()
print(f"MASTER_ID: {n_master_raw:,} filas, {n_master_raw - len(master):,} duplicadas exactas eliminadas")
codigos_por_id = master.groupby("ID_Cliente_Unico")["Cliente"].agg(lambda s: set(s))

ids_en_resultado = set(rf3["ID_Cliente_Unico"])
ids_modelo = set(modelo["ID_Cliente_Unico"])
faltan = sorted(ids_modelo - ids_en_resultado)
sobran = sorted(ids_en_resultado - ids_modelo)
motivos17 = []
for i in faltan:
    cods = codigos_por_id.get(i)
    if cods is None:
        mot = "El ID no existe en MASTER_ID (el join Cliente->ID de la hoja Clientes lo pierde)"
    elif cods.issubset(set(CLIENTES_EXCLUIR)):
        mot = "Todos sus códigos en MASTER_ID son genéricos (CLIENTES_EXCLUIR) y el filtro de seguridad los quita"
    else:
        mot = "OTRO (revisar)"
    motivos17.append(mot)
det17 = modelo.loc[modelo["ID_Cliente_Unico"].isin(faltan), ["ID_Cliente_Unico", "Metal_3V", "GrupoEstructural_3V", "Estado_3V", "Valor", "n_compras"]].copy()
det17["Motivo"] = det17["ID_Cliente_Unico"].map(dict(zip(faltan, motivos17)))
n_unicos_clientes_sheet = rf3["ID_Cliente_Unico"].nunique()
chk("Diferencia Resumen_Estado vs IDs únicos de la hoja Clientes", "273.399 - 273.382 = 17",
    f"{UNIVERSO:,} - {n_unicos_clientes_sheet:,} = {UNIVERSO - n_unicos_clientes_sheet}",
    UNIVERSO - n_unicos_clientes_sheet == 17 and len(faltan) == 17 and not sobran,
    f"IDs del modelo sin fila en Clientes: {len(faltan)}; motivos: {pd.Series(motivos17).value_counts().to_dict()}; IDs en Clientes fuera del modelo: {len(sobran)}")
chk("Reparto por Metal de los 17 vs diferencia contra Excel v3 actual (8/5/4)", "Diamante 8 / Oro 5 / Plata 4",
    det17["Metal_3V"].value_counts().to_dict(), True, "informativo")

# ---------------------------------------------------------------- Provenzal
sc = leer("SELECT Cliente, Segmento, Venta12M, Venta24M, Recencia, UltimaCompra FROM Ventas_Comerssia.dbo.Segmentacion_Clientes")
n_sc = len(sc)
assert sc["Cliente"].is_unique
sc = sc.merge(master, on="Cliente", how="left")
sin_id = sc[sc["ID_Cliente_Unico"].isna()]
chk("Códigos de Provenzal sin ID en MASTER_ID", "informativo", f"{len(sin_id):,} códigos (Venta12M={sin_id['Venta12M'].sum():,.0f}); quedan fuera del cruce",
    True, f"de {n_sc:,} códigos; segmentos: {sin_id['Segmento'].value_counts().to_dict()}")
sc = sc[sc["ID_Cliente_Unico"].notna()].copy()
sc["ID_Cliente_Unico"] = sc["ID_Cliente_Unico"].astype(master["ID_Cliente_Unico"].dtype)

# Códigos genéricos: se quitan solo de los IDs que además tienen algún código real
# (mismo criterio del modelo, que los excluye por Cliente); un ID solo-genérico se queda.
es_gen = sc["Cliente"].isin(CLIENTES_EXCLUIR)
tiene_real = sc.assign(_real=~es_gen).groupby("ID_Cliente_Unico")["_real"].transform("any")
quitar = es_gen & tiene_real
codigos_quitados = sc.loc[quitar, ["Cliente", "ID_Cliente_Unico", "Segmento", "Venta12M"]]
chk("Códigos genéricos quitados de IDs con códigos reales (Provenzal)", "informativo", f"{len(codigos_quitados)} códigos", True,
    codigos_quitados.to_dict("records") if len(codigos_quitados) else "ninguno")
sc = sc[~quitar].copy()
sc["_rango"] = sc["Segmento"].map(RANGO_SEG)
assert sc["_rango"].notna().all()
sc["Venta12M"] = sc["Venta12M"].fillna(0.0)  # NULL = sin ventas en el periodo
sc["Venta24M"] = sc["Venta24M"].fillna(0.0)
inv_rango = {v: k for k, v in RANGO_SEG.items()}
prov = sc.groupby("ID_Cliente_Unico").agg(
    _rango=("_rango", "max"),
    n_codigos=("Cliente", "nunique"),
    conflicto_segmento=("Segmento", lambda s: s.nunique() > 1),
    Venta12M=("Venta12M", "sum"),
    Venta24M=("Venta24M", "sum"),
    UltimaCompra_Prov=("UltimaCompra", "max"),
).reset_index()
prov["Segmento_Provenzal"] = prov["_rango"].map(inv_rango)
prov = prov.drop(columns="_rango")
print(f"Provenzal: {len(prov):,} IDs | con conflicto de segmento: {int(prov['conflicto_segmento'].sum()):,}")

# ---------------------------------------------------------------- Motivo de salida del modelo
ids_universo = set(prov["ID_Cliente_Unico"]) | ids_modelo
fuera_ids = sorted(ids_universo - ids_modelo)

ids_canal = set()
cl_sql = ", ".join(f"'{c}'" for c in CLIENTES_EXCLUIR)
sk_sql = ", ".join(f"'{c}'" for c in SKUS_EXCLUIR)
ca_sql = ", ".join(f"'{c}'" for c in CANALES_EXCLUIR)
for anio in range(FECHA_INICIO.year, FECHA_CORTE.year + 1):
    q = f"""
    SELECT DISTINCT ID_Cliente_Unico FROM Ventas_Comerssia.dbo.Ventas_Unificadas
    WHERE ID_Cliente_Unico IS NOT NULL AND YEAR(Fecha) = {anio}
      AND Fecha >= '{FECHA_INICIO:%Y-%m-%d}' AND Fecha <= '{FECHA_CORTE:%Y-%m-%d}'
      AND Cliente NOT IN ({cl_sql}) AND Sku NOT IN ({sk_sql}) AND Canal IN ({ca_sql})
    """
    ids_canal |= set(leer(q)["ID_Cliente_Unico"])
print(f"IDs con ventas en Shopify/ChatCenter Web (ventana 2019+): {len(ids_canal):,}")

ids_neg = set(excl_neg["ID_Cliente_Unico"])
solo_genericos = {i for i, cods in codigos_por_id.items() if cods.issubset(set(CLIENTES_EXCLUIR))}


def motivo(i):
    if i in ids_modelo:
        return ""
    if i in solo_genericos:
        return "Cliente excluido"
    if i in ids_neg:
        return "Margen negativo"
    if i in ids_canal:
        return "Solo Shopify-ChatCenter Web"
    return "Otro"


# ---------------------------------------------------------------- Cruce_Maestro
maestro = pd.DataFrame({"ID_Cliente_Unico": sorted(ids_universo)})
maestro = maestro.merge(prov, on="ID_Cliente_Unico", how="left").merge(modelo, on="ID_Cliente_Unico", how="left")
maestro["Segmento_Provenzal"] = maestro["Segmento_Provenzal"].fillna(SIN_REGISTRO)
maestro["En_Modelo"] = np.where(maestro["ID_Cliente_Unico"].isin(ids_modelo), "Sí", "No")
maestro["Motivo_Fuera"] = maestro["ID_Cliente_Unico"].map(motivo)
maestro = maestro.rename(columns={"n_codigos": "n_codigos_Provenzal"})
COLS = ["ID_Cliente_Unico", "Segmento_Provenzal", "En_Modelo", "Motivo_Fuera",
        "Metal_3V", "GrupoEstructural_3V", "Estado_3V", "Metal_4V", "GrupoEstructural_4V", "Estado_4V",
        "Valor", "Frecuencia", "n_compras", "Venta12M", "Venta24M",
        "n_codigos_Provenzal", "conflicto_segmento"]
maestro = maestro[COLS]
assert maestro["ID_Cliente_Unico"].is_unique
chk("Cruce_Maestro: un registro por ID (Provenzal ∪ modelo)", len(ids_universo), len(maestro), len(maestro) == len(ids_universo))
chk("Cruce_Maestro: En_Modelo=Sí = universo del modelo", UNIVERSO, int((maestro["En_Modelo"] == "Sí").sum()), int((maestro["En_Modelo"] == "Sí").sum()) == UNIVERSO)
mot_cnt = maestro.loc[maestro["En_Modelo"] == "No", "Motivo_Fuera"].value_counts().to_dict()
chk("Motivo_Fuera de los IDs fuera del modelo", "informativo", mot_cnt, True, f"total fuera = {sum(mot_cnt.values()):,}")
otros = maestro[(maestro["Motivo_Fuera"] == "Otro")]
chk("'Otro': desglose por segmento Provenzal", "informativo", otros["Segmento_Provenzal"].value_counts().to_dict(), True,
    f"con UltimaCompra Prov. anterior a 2019 o sin compra: {int((otros['ID_Cliente_Unico'].isin(prov.loc[(prov['UltimaCompra_Prov'].isna()) | (prov['UltimaCompra_Prov'] < FECHA_INICIO), 'ID_Cliente_Unico'])).sum()):,}")

# ---------------------------------------------------------------- matrices
mx = maestro.copy()
for c in ["Metal_3V", "GrupoEstructural_3V", "Estado_3V", "Metal_4V", "GrupoEstructural_4V", "Estado_4V"]:
    mx[c] = mx[c].fillna(FUERA)
SEG_ROWS = ORDEN_SEG_PROV + [SIN_REGISTRO]
metales3 = ["Diamante", "Oro", "Plata"]
metales4 = ["Diamante", "Oro", "Plata", "Bronce"]
grupos3 = [f"{m} / {f}" for m in metales3 for f in ORDEN_FREC]
grupos4 = [f"{m} / {f}" for m in metales4 for f in ORDEN_FREC]
matrices_log = []


def matriz(df, fila, col, orden_f, orden_c, universo_esperado, nombre):
    ct = pd.crosstab(df[fila], df[col]).reindex(index=orden_f, columns=orden_c, fill_value=0)
    extra_f = set(df[fila].unique()) - set(orden_f)
    extra_c = set(df[col].unique()) - set(orden_c)
    assert not extra_f and not extra_c, f"{nombre}: categorías fuera de orden {extra_f} {extra_c}"
    total = int(ct.values.sum())
    assert total == universo_esperado, f"{nombre}: suma {total} != universo {universo_esperado}"
    matrices_log.append({"Matriz": nombre, "Suma de la matriz": total, "Universo esperado": universo_esperado, "Estado": "OK"})
    cnt = ct.copy()
    cnt["Total"] = cnt.sum(axis=1)
    cnt.loc["Total"] = cnt.sum(axis=0)
    pf = ct.div(ct.sum(axis=1).replace(0, np.nan), axis=0).fillna(0) * 100
    pf["Total"] = pf.sum(axis=1).round(6)
    pc = ct.div(ct.sum(axis=0).replace(0, np.nan), axis=1).fillna(0) * 100
    pc.loc["Total"] = pc.sum(axis=0).round(6)
    return cnt, pf.round(2), pc.round(2), ct


def escribir_matriz(writer, hoja, titulo, cnt, pf, pc, extra=None):
    r = 0
    for nombre, t, fmt in (("Conteo", cnt, None), ("% por fila (cada fila suma 100)", pf, None), ("% por columna (cada columna suma 100)", pc, None)):
        pd.DataFrame([[f"{titulo} — {nombre}"]]).to_excel(writer, sheet_name=hoja, startrow=r, header=False, index=False)
        t.to_excel(writer, sheet_name=hoja, startrow=r + 1)
        r += len(t) + 4
    if extra is not None:
        for nombre, t in extra:
            pd.DataFrame([[nombre]]).to_excel(writer, sheet_name=hoja, startrow=r, header=False, index=False)
            t.to_excel(writer, sheet_name=hoja, startrow=r + 1)
            r += len(t) + 4


hojas = {}
UNIV_TODO = len(mx)
for tag, met, gru in (("3V", metales3, grupos3), ("4V", metales4, grupos4)):
    hojas[f"Prov_x_Metal{tag}"] = (f"Provenzal × Metal_{tag}", matriz(mx, "Segmento_Provenzal", f"Metal_{tag}", SEG_ROWS, met + [FUERA], UNIV_TODO, f"Provenzal × Metal_{tag}"), None)
    hojas[f"Prov_x_Grupo{tag}"] = (f"Provenzal × GrupoEstructural_{tag}", matriz(mx, "Segmento_Provenzal", f"GrupoEstructural_{tag}", SEG_ROWS, gru + [FUERA], UNIV_TODO, f"Provenzal × GrupoEstructural_{tag}"), None)
    hojas[f"Prov_x_Estado{tag}"] = (f"Provenzal × Estado_{tag}", matriz(mx, "Segmento_Provenzal", f"Estado_{tag}", SEG_ROWS, ORDEN_ESTADO + [FUERA], UNIV_TODO, f"Provenzal × Estado_{tag}"), None)

en = mx[mx["En_Modelo"] == "Sí"]
assert len(en) == UNIVERSO
m34 = matriz(en, "Metal_3V", "Metal_4V", metales3, metales4, UNIVERSO, "Metal_3V × Metal_4V")
g34 = matriz(en, "GrupoEstructural_3V", "GrupoEstructural_4V", grupos3, grupos4, UNIVERSO, "GrupoEstructural_3V × GrupoEstructural_4V")


def cambio(ct, nombre):
    cambian = 0
    filas = []
    for f in ct.index:
        fila_tot = int(ct.loc[f].sum())
        igual = int(ct.loc[f, f]) if f in ct.columns else 0
        cambian += fila_tot - igual
        filas.append({"Segmento 3V": f, "Clientes": fila_tot, "Conservan el mismo nombre en 4V": igual,
                      "Cambian de segmento": fila_tot - igual, "% que cambia": round((fila_tot - igual) / fila_tot * 100, 2) if fila_tot else 0.0})
    tot = int(ct.values.sum())
    filas.append({"Segmento 3V": "TOTAL", "Clientes": tot, "Conservan el mismo nombre en 4V": tot - cambian,
                  "Cambian de segmento": cambian, "% que cambia": round(cambian / tot * 100, 2)})
    return pd.DataFrame(filas).set_index("Segmento 3V"), cambian, tot


cm, cambian_m, tot_m = cambio(m34[3], "Metal")
cg, cambian_g, tot_g = cambio(g34[3], "Grupo")
chk("% de clientes que cambia de Metal (3V→4V)", "informativo", f"{cambian_m:,} de {tot_m:,} ({cambian_m / tot_m * 100:.2f}%)", True, "cambia = el nombre del Metal es distinto")
chk("% de clientes que cambia de GrupoEstructural (3V→4V)", "informativo", f"{cambian_g:,} de {tot_g:,} ({cambian_g / tot_g * 100:.2f}%)", True)
hojas["Metal3V_x_Metal4V"] = ("Metal_3V × Metal_4V", m34, [("Clientes que cambian de segmento (por Metal_3V de origen)", cm)])
hojas["Grupo3V_x_Grupo4V"] = ("GrupoEstructural_3V × GrupoEstructural_4V", g34, [("Clientes que cambian de segmento (por GrupoEstructural_3V de origen)", cg)])

# ---------------------------------------------------------------- perfil por segmento (universo del modelo)
def stats(g, nombre_modelo, estado_col_list):
    out = []
    for seg, d in g:
        row = {"Modelo": nombre_modelo, "Segmento": seg, "n": len(d), "% del universo": round(len(d) / UNIVERSO * 100, 2)}
        for v, pref in (("Valor", "Valor"), ("Frecuencia", "Frec")):
            row.update({f"{pref}_min": d[v].min(), f"{pref}_media": d[v].mean(), f"{pref}_mediana": d[v].median(), f"{pref}_max": d[v].max()})
        row["Venta12M_mediana"] = d["Venta12M"].median()
        row["Venta24M_mediana"] = d["Venta24M"].median()
        row["n_con_dato_Provenzal"] = int(d["Venta12M"].notna().sum())
        for c in estado_col_list:
            row[f"%Activo ({c})"] = round((d[c] == "Activo").mean() * 100, 2)
        out.append(row)
    return out


en = en.copy()
filas = []
orden_prov = SEG_ROWS
filas += stats([(s, en[en["Segmento_Provenzal"] == s]) for s in orden_prov if (en["Segmento_Provenzal"] == s).any()], "Provenzal", ["Estado_3V", "Estado_4V"])
filas += stats([(s, en[en["Metal_3V"] == s]) for s in metales3], "3V", ["Estado_3V"])
filas += stats([(s, en[en["Metal_4V"] == s]) for s in metales4], "4V", ["Estado_4V"])
perfil = pd.DataFrame(filas)
for n_ in ["Valor", "Frec"]:
    pass
chk("Perfil: cada modelo suma el universo", UNIVERSO, {m: int(perfil.loc[perfil['Modelo'] == m, 'n'].sum()) for m in ['Provenzal', '3V', '4V']},
    all(int(perfil.loc[perfil['Modelo'] == m, 'n'].sum()) == UNIVERSO for m in ['Provenzal', '3V', '4V']))

# ---------------------------------------------------------------- foco Diamante/Alta y Oro/Alta
focos = [("Diamante / Alta", "3V"), ("Oro / Alta", "3V"), ("Diamante / Alta", "4V"), ("Oro / Alta", "4V")]
comp = {}
for g_, tag in focos:
    sub = mx[mx[f"GrupoEstructural_{tag}"] == g_]
    s = sub["Segmento_Provenzal"].value_counts().reindex(SEG_ROWS, fill_value=0)
    comp[f"{g_} ({tag}) n"] = s
    comp[f"{g_} ({tag}) %"] = (s / s.sum() * 100).round(2) if s.sum() else s
comp = pd.DataFrame(comp)
comp.loc["Total"] = [comp[c].sum() if c.endswith(" n") else round(comp[c].sum(), 2) for c in comp.columns]
dist = {}
for seg in ["Diamante", "Oro"]:
    sub = mx[mx["Segmento_Provenzal"] == seg]
    for tag, gru in (("3V", grupos3), ("4V", grupos4)):
        s = sub[f"GrupoEstructural_{tag}"].value_counts()
        dist[(seg, tag)] = s
dist_tabs = []
for tag, gru in (("3V", grupos3), ("4V", grupos4)):
    t = pd.DataFrame(index=gru + [FUERA])
    for seg in ["Diamante", "Oro"]:
        s = dist[(seg, tag)].reindex(gru + [FUERA], fill_value=0)
        t[f"{seg} Provenzal (n)"] = s
        t[f"{seg} Provenzal (%)"] = (s / s.sum() * 100).round(2)
    t.loc["Total"] = [t[c].sum() if c.endswith("(n)") else round(t[c].sum(), 2) for c in t.columns]
    dist_tabs.append((f"Distribución de los Diamante y Oro actuales de Provenzal en los grupos {tag}", t))
for seg in ["Diamante", "Oro"]:
    n_prov = int((mx["Segmento_Provenzal"] == seg).sum())
    for tag in ["3V", "4V"]:
        assert int(dist[(seg, tag)].sum()) == n_prov
matrices_log.append({"Matriz": "Foco: distribución Diamante/Oro Provenzal en grupos 3V y 4V", "Suma de la matriz": "ver hoja", "Universo esperado": "n de cada segmento Provenzal", "Estado": "OK"})

# ---------------------------------------------------------------- escritura
with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as w:
    pd.DataFrame(checks).to_excel(w, sheet_name="Chequeos", index=False)
    pd.DataFrame(matrices_log).to_excel(w, sheet_name="Chequeos", index=False, startrow=len(checks) + 3)
    det17.to_excel(w, sheet_name="Chequeo_17_IDs", index=False)
    for hoja, (titulo, (cnt, pf, pc, _), extra) in hojas.items():
        escribir_matriz(w, hoja, titulo, cnt, pf, pc, extra)
    perfil.to_excel(w, sheet_name="Perfil_Segmentos", index=False)
    pd.DataFrame([["Composición por segmento Provenzal de los grupos foco (universo: Cruce_Maestro)"]]).to_excel(w, sheet_name="Foco_Alta", startrow=0, header=False, index=False)
    comp.to_excel(w, sheet_name="Foco_Alta", startrow=1)
    r = len(comp) + 4
    for nombre, t in dist_tabs:
        pd.DataFrame([[nombre]]).to_excel(w, sheet_name="Foco_Alta", startrow=r, header=False, index=False)
        t.to_excel(w, sheet_name="Foco_Alta", startrow=r + 1)
        r += len(t) + 4
    maestro.to_excel(w, sheet_name="Cruce_Maestro", index=False)

print(f"Escrito {OUT_XLSX}")
estados = pd.DataFrame(checks)["Estado"].value_counts().to_dict()
print("Chequeos:", estados)
