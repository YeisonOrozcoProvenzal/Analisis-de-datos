# -*- coding: utf-8 -*-
"""Corre segmentacion_valor_frecuencia_v3.ipynb DOS veces sobre la MISMA extraccion de
datos (misma foto de la base): K_VALOR=3 y K_VALOR=4 (K_FRECUENCIA=3 en ambas).

No modifica el notebook v3: ejecuta sus celdas de codigo tal cual, aplicando en memoria
solo (a) los parametros de la corrida y (b) los nombres de archivos de salida (_v3_3V /
_v3_4V) para no pisar archivos anteriores. Deja los resultados intermedios en
resultados_2019_con_unico/_work/ para el analisis posterior."""
import copy
import json
import os
import pathlib
import re
import sys
import time

os.environ.setdefault("MPLBACKEND", "Agg")
HERE = pathlib.Path(__file__).resolve().parent
os.chdir(HERE)

NB_PATH = HERE / "segmentacion_valor_frecuencia_v3.ipynb"
WORK = HERE / "resultados_2019_con_unico" / "_work"
WORK.mkdir(parents=True, exist_ok=True)

nb = json.load(open(NB_PATH, encoding="utf-8"))
codigo = []
for c in nb["cells"]:
    if c["cell_type"] == "code":
        s = c["source"]
        codigo.append("".join(s) if isinstance(s, list) else s)


def idx_de(marker):
    hits = [i for i, s in enumerate(codigo) if marker in s]
    assert len(hits) == 1, f"marcador {marker!r}: {len(hits)} coincidencias"
    return hits[0]


i_params = idx_de("fecha_corte = pd.to_datetime(")
i_extraccion = idx_de("df_ventas = pd.concat(partes")
i_kvalor = idx_de("K_VALOR = 3")
# Parte "base" = hasta la primera celda que deja armado df_visitas (extraccion + margen + devoluciones)
i_fin_base = next(i for i, s_ in enumerate(codigo) if "df_visitas = " in s_)
assert i_extraccion < i_fin_base < i_kvalor
print(f"celdas de codigo: {len(codigo)} | params={i_params} extraccion={i_extraccion} fin_base={i_fin_base} K_VALOR={i_kvalor}")


def preparar(src, tag, kv):
    s = src
    if "SEPARAR_UNICO = " in s and "FECHA_CORTE_INICIO" in s:
        s = re.sub(r"SEPARAR_UNICO = (True|False)", "SEPARAR_UNICO = True", s, count=1)
        assert "SEPARAR_UNICO = True" in s
    s = re.sub(r"K_VALOR = \d+", f"K_VALOR = {kv}", s, count=1) if "K_VALOR = 3" in s else s
    s = s.replace("_v2.png", f"_v3_{tag}.png")
    s = s.replace("segmentacion_valor_frecuencia_v3.xlsx", f"segmentacion_valor_frecuencia_v3_{tag}.xlsx")
    s = s.replace("segmentacion_valor_frecuencia_v2.xlsx", f"segmentacion_valor_frecuencia_v3_{tag}.xlsx")
    return s


def correr(celdas, ns, tag, kv):
    for i, src in celdas:
        t0 = time.time()
        try:
            exec(compile(preparar(src, tag, kv), f"<celda {i}>", "exec"), ns)
        except Exception:
            print(f"!!! FALLO en la celda de codigo #{i} (corrida {tag})", flush=True)
            raise
        print(f"[{tag}] celda #{i} ok ({time.time() - t0:.0f}s)", flush=True)


base = {"__name__": "__main__"}
print(">>> BASE (extraccion unica compartida por 3V y 4V)", flush=True)
correr([(i, codigo[i]) for i in range(0, i_fin_base + 1)], base, "base", 3)

# La foto de la base queda congelada en disco para auditoria
base["df_ventas"].to_pickle(WORK / "df_ventas_snapshot.pkl")

restantes = [(i, codigo[i]) for i in range(i_fin_base + 1, len(codigo))]
resultados = {}
for tag, kv in (("3V", 3), ("4V", 4)):
    print(f">>> CORRIDA {tag}: K_VALOR={kv}, K_FRECUENCIA=3", flush=True)
    ns = dict(base)  # df_ventas / df_visitas compartidos por referencia (las celdas siguientes no los mutan)
    n_filas = len(base["df_ventas"])
    correr(restantes, ns, tag, kv)
    assert len(base["df_ventas"]) == n_filas
    assert ns["K_VALOR"] == kv and ns["K_FRECUENCIA"] == 3 and ns["SEPARAR_UNICO"] is True
    ns["clientes"].to_pickle(WORK / f"clientes_{tag}.pkl")
    ns["resultado_final"].to_pickle(WORK / f"resultado_final_{tag}.pkl")
    ns["umbrales"].to_pickle(WORK / f"umbrales_{tag}.pkl")
    ns["clientes_excluidos_margen_negativo"].to_pickle(WORK / f"excluidos_margen_negativo_{tag}.pkl")
    resultados[tag] = {"orden_metales": ns["orden_metales"], "n": len(ns["clientes"])}
    print(f"<<< {tag}: {len(ns['clientes']):,} clientes | metales {ns['orden_metales']}", flush=True)
    del ns

json.dump(resultados, open(WORK / "resumen_corridas.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("LISTO", resultados)
