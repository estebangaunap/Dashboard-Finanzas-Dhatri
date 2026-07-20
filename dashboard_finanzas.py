#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dhatri — Dashboard de Finanzas (Egresos a pagar + Proyección Ingresos vs Egresos)
==================================================================================
Lee:
  - finanzas_dhatri.xlsx (export de "Finanzas Dhatri" Google Sheet)
      hoja "Junio 26"  -> listado de egresos a pagar del mes
      hoja "Pagos Mayo" -> para detectar pendientes de mayo no incluidos en junio
  - Notion (REGISTRO DE INGRESOS & EGRESOS) -> para:
      * marcar como "Pagado" automaticamente los items de junio que ya
        tengan un movimiento de Egreso registrado con monto similar
      * calcular proyeccion Ingresos vs Egresos (promedio ult. 3 meses + 10% crecimiento
        en Clases y Terapias)

Salida: dashboard_finanzas_dhatri.html
"""

import os
import sys
import urllib3
import requests
import openpyxl
from collections import defaultdict
from calendar import monthrange
from datetime import datetime

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from config import NOTION_TOKEN, DB_ID, DEPT

XLSX_PATH = "finanzas_dhatri.xlsx"
OUT_PATH  = "dashboard_finanzas_dhatri.html"

MES_ACTUAL_NOMBRE = "Julio 2026"
ANIO_ACTUAL, MES_ACTUAL = 2026, 7

CATEGORIAS_VALIDAS = {
    "Arriendo", "Sistemas", "Directivo", "Deudas", "Remuneraciones",
    "Cuentas & Gastos Comunes", "Marketing", "Donacion", "Emporio",
    "Arreglos", "Inversiones",
}

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

# Items del listado de pagos que en realidad se pagan en varias transferencias
# a lo largo del mes (no de una sola vez), y que en la BD de Ingresos&Egresos
# quedan registrados como Egreso / depto "Administración" a nombre de la persona.
CONTADORES = [
    {
        "label": "Pago Vish & Gita",
        "nombre_sheet": "Pago Vish & Gita",
        "detalles_notion": {"Diego Maldonado"},
        "dept_filter": "Administración",
    },
    {
        "label": "Pago Esteban",
        "nombre_sheet": "Pago Esteban",
        "detalles_notion": {"Esteban Gauna", "Esteban Andres Gauna Palavecino"},
        "dept_filter": "Administración",
    },
    {
        "label": "Pago Nicolas Ibarra",
        "nombre_sheet": "Pago Nicolas Ibarra",
        "detalles_notion": {"Nicolas Ibarra"},
        "dept_filter": None,
    },
    {
        "label": "Pago Elizabeth",
        "nombre_sheet": "Pago Elizabeth",
        "detalles_notion": {"Elizabeth Ibacache"},
        "dept_filter": None,
    },
    {
        "label": "Pago Amanda",
        "nombre_sheet": "Pago Amanda",
        "detalles_notion": {"Amanda Serrano", "Amanda"},
        "dept_filter": None,
    },
    # Marketing: suma todos los egresos del dept Marketing vs presupuesto mensual
    {
        "label": "Marketing",
        "nombre_sheet": "Pago Marketing",
        "detalles_notion": set(),
        "dept_filter": "Marketing",
        "match_all_in_dept": True,
        "presupuesto": 200_000,
    },
]

# Áreas cuyas líneas se agrupan en una sola fila con detalle colapsable en la tabla
GRUPOS_EN_TABLA = {"Operacional", "Emporio"}

# Override de categoría para ítems mal clasificados en el xlsx
CATEGORIA_OVERRIDE = {
    "Wix.com":            "Sistemas",
    "Wix.com 1212841827": "Sistemas",
}

DEPT_DISPLAY = {
    "clases": "Clases", "ayurveda": "Ayurveda", "terapias": "Terapias",
    "admin": "Administración", "educacion": "Educación", "emporio": "Emporio",
    "marketing": "Marketing", "dia_dhatri": "Día Dhatri", "recepcion": "Recepción",
    "tecnologia": "Tecnología", "arriendo": "Arriendo", "finanzas": "Finanzas",
    "operacional": "Operacional",
}

# Mes siguiente (para la proyeccion de costos estimados / meta)
MES_PROX_NOMBRE = "Agosto 2026"
ANIO_PROX, MES_PROX = 2026, 8

# Mapeo Nombre del item de la planilla -> Area (mismos nombres que usan los
# depts de Notion, para poder comparar "basal" vs "real" en la misma grilla,
# y para que coincida con las areas del cuadro de reducciones de costos).
AREA_MAP = {
    "Garantía arriendo": "Arriendo",
    "Boletas": "Tecnología",
    "Pago Vish & Gita": "Administración",
    "Arriendo Las condes": "Arriendo",
    "Arriendo Provi": "Arriendo",
    "Pago Esteban": "Administración",
    "Zoho-workplace": "Tecnología",
    "Www.albato.com": "Tecnología",
    "Canva": "Tecnología",
    "Pago Deuda Paty": "Deudas",
    "Wix.com 1212841827": "Tecnología",
    "Wix.com": "Tecnología",
    "Celulares Dhatri": "Operacional",
    "Pago Amanda": "Educación",
    "Pago Elizabeth": "Recepción",
    "Zoom.com 888-799-966": "Tecnología",
    "Pago Internet": "Operacional",
    "Pago Marketing": "Marketing",
    "Pago Nicolas Ibarra": "Recepción",
    "Pago Constanza Muñoz": "Clases",
    "Pago Melisa Guajardo": "Clases",
    "Pago Constanza Torres": "Clases",
    "Pago Catalina Suckel": "Clases",
    "Pago Agua Dhatri": "Operacional",
    "Pago Katherine Villar": "Clases",
    "Pago Coni Aguayo": "Clases",
    "Pago Luz La Gloria": "Operacional",
    "Pago Paulina Osoroio": "Clases",
    "Pago Nayely": "Recepción",
    "Pago Cesar Lillo": "Clases",
    "Pago Consuelo Silva": "Clases",
    "Credito Esteban": "Deudas",
    "Pago SII Dhatri SPA": "Operacional",
    "Pago Contador Diego Cubillos": "Operacional",
    "Pago Contador Diego":          "Operacional",
    "Tummee yoga platform": "Tecnología",
    "INtereses Credito Esteban": "Deudas",
}
AREA_DEFAULT = "Otros"

# Items extra de costos para la proyeccion del mes siguiente (no estan en el
# listado del sheet porque son nuevos / puntuales, o estaban ausentes del
# sheet pero se confirmaron en el registro de Ingresos & Egresos de Notion)
COSTOS_EXTRA_MES_PROX = [
    {"nombre": "Aire acondicionado", "area": "Operacional", "monto": 50_000},
    {"nombre": "Pago deuda Emporio", "area": "Emporio", "monto": 100_000},
    # Suscripcion mensual no incluida en el sheet — encontrada en Notion
    # ("Notion labs, inc.", depto Tecnologia), monto del ultimo cobro (jun-26)
    {"nombre": "Notion labs, inc.", "area": "Tecnología", "monto": 21_429},
]

# Items del listado de Junio que NO se repiten en Julio (pagos puntuales,
# duplicados, o deudas/creditos que se manejan en la sección de Deudas)
ITEMS_NO_RECURRENTES = {
    "Garantía arriendo",
    "Wix.com 1212841827",        # duplicado de "Wix.com" (26.448) — se deja solo este
    "Credito Esteban",          # se maneja en la sección de Deudas
    "INtereses Credito Esteban", # idem
    "Pago SII Dhatri SPA",
}

# Ajustes puntuales de monto para Julio (y base de la proyección Agosto)
OVERRIDES_MES_PROX = {
    "Arriendo Las condes": 2_500_000,  # Claudia Berkhoff: bajó de 3M a 2.5M
    "Arriendo Provi":        250_000,  # Fijo hasta octubre
}

# ─────────────────────────────────────────────
# Pagos de personal Julio 2026 (Planilla Pagos Personal Julio 26)
# Reglas:
#   Consuelo, Amanda, Diego Contador, Nayely → Valor Bruto
#   Resto → Valor Neto
#   Nicolás Ibarra: Valor Neto total − (135.000 × 0,7) por ajuste de clases
# ─────────────────────────────────────────────
PAGOS_PERSONAL_JULIO = [
    # Consuelo y Cesar excluidos: pagos periodo Junio pagados con atraso en Julio
    # → aparecen en "adicionales" sumados; los de Julio se cargarán cuando llegue la planilla
    {"nombre": "Pago Constanza Aguayo",     "categoria": "Remuneraciones", "dia": 5,  "monto": 114_413},
    {"nombre": "Pago Constanza Muñoz",      "categoria": "Remuneraciones", "dia": 5,  "monto": 51_528},
    {"nombre": "Pago Amanda",               "categoria": "Remuneraciones", "dia": 15, "monto": 290_000},
    # Neto (199.163+96.615+162.720=458.498) − 135.000×0,7 (ajuste clases) = 363.998
    {"nombre": "Pago Nicolas Ibarra",       "categoria": "Remuneraciones", "dia": 5,  "monto": 363_998},
    {"nombre": "Pago Katherine Villar",     "categoria": "Remuneraciones", "dia": 5,  "monto": 40_680},
    {"nombre": "Pago Nayely",               "categoria": "Remuneraciones", "dia": 5,  "monto": 88_000},
    {"nombre": "Pago Catalina Suckel",      "categoria": "Remuneraciones", "dia": 5,  "monto": 101_700},
    {"nombre": "Pago Elizabeth",            "categoria": "Remuneraciones", "dia": 5,  "monto": 437_310},
    {"nombre": "Pago Constanza Torres",     "categoria": "Remuneraciones", "dia": 5,  "monto": 114_413},
    {"nombre": "Pago Contador Diego",       "categoria": "Directivo",      "dia": 30, "monto": 79_285},
]

# Nombres de personal en la hoja de Junio que se reemplazan con los valores de Julio
NOMBRES_PERSONALES_JUNIO = {
    "Pago Cesar Lillo", "Pago Consuelo Silva", "Pago Constanza Aguayo",  # Cesar y Consuelo: pago Junio atrasado → adicionales
    "Pago Constanza Muñoz", "Pago Amanda", "Pago Nicolas Ibarra",
    "Pago Katherine Villar", "Pago Nayely", "Pago Catalina Suckel",
    "Pago Elizabeth", "Pago Constanza Torres", "Pago Contador Diego Cubillos",
    "Pago Melisa Guajardo", "Pago Paulina Osoroio",
}

# ─────────────────────────────────────────────
# Deudas pendientes (consolidado manual, fuera de los flujos mensuales)
# ─────────────────────────────────────────────
DEUDAS_PENDIENTES = [
    {"nombre": "Garantía arriendo", "categoria": "Arriendo", "monto": 3_000_000, "desde": "01/05/2026"},
    {"nombre": "Credito Bco Estado Dhatri Vish 1", "categoria": "Deudas", "monto": 54_521, "desde": "05/05/2026"},
    {"nombre": "Credito Bco Estado Dhatri Vish 2", "categoria": "Deudas", "monto": 224_665, "desde": "05/05/2026"},
    {"nombre": "Credito Bco Estado Dhatri Vish 3", "categoria": "Deudas", "monto": 38_851, "desde": "05/05/2026"},
    {"nombre": "Deuda provi", "categoria": "Deudas", "monto": 50_000, "desde": "30/05/2026"},
    {"nombre": "Deuda Emporio", "categoria": "Emporio", "monto": 400_000, "desde": None},
]

SII_HISTORICO = [
    ("Octubre 2024", 986_192), ("Noviembre 2024", 911_501), ("Diciembre 2024", 413_807),
    ("Enero 2025", 450_805), ("Febrero 2025", 597_166), ("Abril 2025", 539_426),
    ("Mayo 2025", 676_925), ("Junio 2025", 51_734),
    ("Febrero 2026", 824_917), ("Marzo 2026", 374_949), ("Abril 2026", 489_097), ("Mayo 2026", 508_759),
]

INVERSIONES_DB_URL = "https://www.notion.so/20157760cde1806b9458cde9d8efbf38?v=20157760cde180a090b4000c3d547b4d"

# Snapshot de la base Notion "Inversiones - Creditos - Deudas"
# (consultada via notion-query-data-sources; actualizar manualmente o
# reemplazar por una consulta en vivo si se conecta la API de Notion AI)
INVERSIONES_NOTION = [
    {"nombre": "Lorena",   "tipo": "Préstamos", "estado": "Inversión recibida", "valor": 1_000_000, "fecha_inv": "2025-05-14", "valor_devolucion": 1_300_000, "fecha_dev": "2026-07-15", "obs": "30%"},
    {"nombre": "Adrian",   "tipo": "Préstamos", "estado": "Inversión recibida", "valor": 1_000_000, "fecha_inv": "2025-05-12", "valor_devolucion": 1_300_000, "fecha_dev": "2026-05-12", "obs": "30%"},
    {"nombre": "Giovanni", "tipo": "Préstamos", "estado": "Inversión recibida", "valor": 3_000_000, "fecha_inv": "2025-01-30", "valor_devolucion": 4_500_000, "fecha_dev": "2027-01-30", "obs": "50%, una sola cuota al 2027"},
    {"nombre": "Lola",     "tipo": "Préstamos", "estado": "Postergada",        "valor": 2_000_000, "fecha_inv": "2024-12-16", "valor_devolucion": 2_500_000, "fecha_dev": "2025-12-17", "obs": "25% de interés"},
    {"nombre": "Nitai",    "tipo": "Préstamos", "estado": "Por postergar",     "valor": 2_750_000, "fecha_inv": "2024-11-01", "valor_devolucion": 3_575_000, "fecha_dev": "2027-11-01", "obs": "2 cuotas 1.787.500, 30%"},
    {"nombre": "Mamá Vish","tipo": "Préstamos", "estado": "Por postergar",     "valor": 600_000,   "fecha_inv": "2024-10-15", "valor_devolucion": 600_000,   "fecha_dev": None,         "obs": "sin interés, préstamo por urgencia"},
    {"nombre": "Ainara",   "tipo": "Préstamos", "estado": "Postergada",        "valor": 2_000_000, "fecha_inv": "2024-10-14", "valor_devolucion": 3_000_000, "fecha_dev": "2025-10-15", "obs": "2 cuotas, 1.5M c/u, 50%"},
    {"nombre": "Esteban",  "tipo": "Inversión", "estado": "Inversión recibida", "valor": 12_500_000,"fecha_inv": "2024-09-01", "valor_devolucion": None,      "fecha_dev": None,         "obs": None},
    {"nombre": "Paty",     "tipo": "Préstamos", "estado": "Postergada",        "valor": 12_500_000,"fecha_inv": "2024-09-01", "valor_devolucion": 15_000_000,"fecha_dev": "2025-10-15", "obs": "cambio al 20% (desde el 50% previamente establecido)"},
    {"nombre": "Paty",     "tipo": "Préstamos", "estado": "Inversión recibida", "valor": 3_000_000, "fecha_inv": "2026-06-26", "valor_devolucion": None,      "fecha_dev": "2027-01-01", "obs": "Nuevo préstamo — devolución a partir de enero 2027"},
]


# ─────────────────────────────────────────────
# 1. Leer listado de egresos desde el Excel
# ─────────────────────────────────────────────

def leer_egresos(ws):
    """Lee columnas A-E de una hoja: Nombre, Categoria, Monto, Fecha(dia), Estado."""
    items = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        nombre, categoria, monto, fecha, estado = row[0], row[1], row[2], row[3], row[4]
        if nombre is None or monto is None or fecha is None:
            continue
        if categoria not in CATEGORIAS_VALIDAS:
            continue
        items.append({
            "nombre": str(nombre).strip(),
            "categoria": categoria,
            "monto": float(monto),
            "dia": int(fecha),
            "pagado_manual": (estado == "Pagado"),
        })
    return items


def cargar_excel():
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    egresos_junio = leer_egresos(wb["Junio 26"])
    egresos_mayo  = leer_egresos(wb["Pagos Mayo"])
    return egresos_junio, egresos_mayo


# ─────────────────────────────────────────────
# 2. Notion: lectura de movimientos
# ─────────────────────────────────────────────

_notion = requests.Session()
_notion.verify = False
_notion.headers.update({
    "Authorization": f"Bearer {NOTION_TOKEN}",
    "Content-Type": "application/json",
    "Notion-Version": "2022-06-28",
})


def _build_dept_reverse_map():
    mapping = {}
    for key, url in DEPT.items():
        raw = url.rstrip("/").split("/")[-1].replace("-", "")
        if len(raw) >= 32:
            page_id = f"{raw[:8]}-{raw[8:12]}-{raw[12:16]}-{raw[16:20]}-{raw[20:]}"
            mapping[page_id] = DEPT_DISPLAY.get(key, key.capitalize())
    return mapping


DEPT_ID_TO_NAME = _build_dept_reverse_map()


def query_notion_month(year, month):
    """Devuelve todos los registros (Ingreso/Egreso) del mes desde Notion."""
    _, last_day = monthrange(year, month)
    date_from = f"{year}-{month:02d}-01"
    date_to   = f"{year}-{month:02d}-{last_day:02d}"

    records = []
    url = f"https://api.notion.com/v1/databases/{DB_ID}/query"
    payload = {
        "page_size": 100,
        "filter": {
            "and": [
                {"property": "Fecha", "date": {"on_or_after":  date_from}},
                {"property": "Fecha", "date": {"on_or_before": date_to}},
            ]
        },
        "sorts": [{"property": "Fecha", "direction": "ascending"}],
    }

    while True:
        r = _notion.post(url, json=payload)
        r.raise_for_status()
        data = r.json()

        for page in data["results"]:
            props = page["properties"]
            monto = props.get("Monto", {}).get("number") or 0

            mov_sel    = props.get("Movimiento", {}).get("select") or {}
            movimiento = mov_sel.get("name", "")
            if movimiento not in ("Ingreso", "Egreso"):
                continue

            fecha_obj = props.get("Fecha", {}).get("date") or {}
            fecha = fecha_obj.get("start", "")[:10]

            rel = props.get("✊🏽 Departamentos", {}).get("relation", [])
            dept_name = "Sin departamento"
            if rel:
                dept_page_id = rel[0].get("id", "")
                dept_name = DEPT_ID_TO_NAME.get(dept_page_id, "Sin departamento")

            title   = props.get("Detalle", {}).get("title", [])
            detalle = title[0]["plain_text"] if title else ""

            records.append({
                "fecha": fecha, "monto": monto,
                "movimiento": movimiento, "dept": dept_name, "detalle": detalle,
            })

        if not data.get("has_more"):
            break
        payload["start_cursor"] = data["next_cursor"]

    return records


# ─────────────────────────────────────────────
# 3. Matching de egresos del sheet contra Notion
# ─────────────────────────────────────────────

def marcar_pagados(items, registros_notion, tolerancia=500):
    """Marca pagado_auto=True si encuentra un Egreso en Notion con monto similar.

    Los items que correspondan a un "contador" (pagos fraccionados, ej. Vish & Gita
    o Esteban) se excluyen de este matching 1:1 por monto, porque se pagan en
    varias transferencias distintas durante el mes — su avance se calcula aparte.
    """
    nombres_contador = {c["nombre_sheet"] for c in CONTADORES}
    egresos_notion = [r for r in registros_notion if r["movimiento"] == "Egreso"]
    usados = set()
    for it in items:
        it["pagado_auto"] = False
        it["match_notion"] = None
        it["es_contador"] = it["nombre"] in nombres_contador
        if it["pagado_manual"] or it["es_contador"]:
            continue
        for i, r in enumerate(egresos_notion):
            if i in usados:
                continue
            if abs(r["monto"] - it["monto"]) <= tolerancia:
                it["pagado_auto"] = True
                it["match_notion"] = r
                usados.add(i)
                break
    return items


def calcular_contadores(items, registros_notion):
    """Para cada contador, compara el total planificado vs lo pagado en Notion.
    - dept_filter=None : busca en todos los egresos del mes
    - match_all_in_dept: suma todos los egresos del dept (ej. Marketing)
    - presupuesto: override del monto planificado (ej. budget Marketing 200k)
    """
    egresos_notion = [r for r in registros_notion if r["movimiento"] == "Egreso"]
    resultados = []
    for c in CONTADORES:
        dept = c.get("dept_filter")
        pool = [r for r in egresos_notion if dept is None or r["dept"] == dept]
        planificado_sheet = sum(i["monto"] for i in items if i["nombre"] == c["nombre_sheet"])
        planificado = c.get("presupuesto") or planificado_sheet
        if c.get("match_all_in_dept"):
            pagos = pool
        else:
            pagos = [r for r in pool if r["detalle"] in c["detalles_notion"]]
        pagado = sum(r["monto"] for r in pagos)
        resultados.append({
            "label": c["label"],
            "nombre_sheet": c["nombre_sheet"],
            "planificado": planificado,
            "pagado": pagado,
            "pendiente": planificado - pagado,
            "pagos": sorted(pagos, key=lambda r: r["fecha"]),
        })
    return resultados


# ─────────────────────────────────────────────
# 4. Formato
# ─────────────────────────────────────────────

def fmt(v):
    return f"${v:,.0f}".replace(",", ".")


# ─────────────────────────────────────────────
# 5. HTML
# ─────────────────────────────────────────────

def _build_egresos_julio(egresos_junio):
    """Construye la lista de egresos de Julio:
    - items no personales y no puntuales de Junio con overrides aplicados
    - pagos de personal actualizados (PAGOS_PERSONAL_JULIO)
    - campo 'grupo' para agrupar Operacional y Emporio en la tabla
    - CATEGORIA_OVERRIDE corrige categorías mal asignadas en el xlsx
    """
    items = []
    for it in egresos_junio:
        if it["nombre"] in ITEMS_NO_RECURRENTES:
            continue
        if it["nombre"] in NOMBRES_PERSONALES_JUNIO:
            continue
        monto = OVERRIDES_MES_PROX.get(it["nombre"], it["monto"])
        area  = _area_de(it)
        cat   = CATEGORIA_OVERRIDE.get(it["nombre"], it["categoria"])
        grupo = area if area in GRUPOS_EN_TABLA else None
        items.append({**it, "monto": monto, "categoria": cat,
                      "pagado_manual": False, "grupo": grupo})
    for p in PAGOS_PERSONAL_JULIO:
        area  = _area_de(p)
        grupo = area if area in GRUPOS_EN_TABLA else None
        items.append({**p, "pagado_manual": False, "grupo": grupo})
    return items


def _find_extras_notion(egresos_julio, registros_notion):
    """Devuelve egresos de Notion del mes no emparejados con el listado establecido.
    Agrupa Mercado Pago y Consuelo Silva en una sola línea cada uno."""
    matched_keys = set()
    for it in egresos_julio:
        m = it.get("match_notion")
        if m:
            matched_keys.add((m["fecha"], m["monto"], m["detalle"]))

    raw = []
    for r in registros_notion:
        if r["movimiento"] != "Egreso":
            continue
        if (r["fecha"], r["monto"], r["detalle"]) in matched_keys:
            continue
        es_contador = False
        for c in CONTADORES:
            dept = c.get("dept_filter")
            match_all = c.get("match_all_in_dept", False)
            if match_all and dept and r["dept"] == dept:
                es_contador = True
                break
            if r["detalle"] in c["detalles_notion"]:
                if dept is None or r["dept"] == dept:
                    es_contador = True
                    break
        if es_contador:
            continue
        # Pagos operacionales canalizados vía Diego Maldonado (internet, celulares,
        # deuda Paty) → dept distinto a Administración; se excluyen de adicionales
        # porque corresponden a gastos ya contemplados en el listado establecido
        if r["detalle"] == "Diego Maldonado" and r.get("dept") != "Administración":
            continue
        raw.append(r)

    # Agrupar entradas con el mismo tipo en una sola línea sumada
    # Keyword → label visible (se compara con r["detalle"].lower())
    AGRUPAR_KEYWORDS = {
        "mercado":       "Mercado Crédito",   # "Pago de cuotas de Mercado Crédito"
        "consuelo silva": "Consuelo Silva",
        "cesar lillo":   "Cesar Lillo",
    }
    grupos: dict = {}
    individuales = []
    for r in raw:
        clave = next((label for kw, label in AGRUPAR_KEYWORDS.items()
                      if kw in r["detalle"].lower()), None)
        if clave:
            if clave not in grupos:
                grupos[clave] = {"fecha": r["fecha"], "monto": 0,
                                 "detalle": clave, "dept": r["dept"], "n": 0}
            grupos[clave]["monto"] += r["monto"]
            grupos[clave]["n"]     += 1
        else:
            individuales.append(r)

    extras = individuales[:]
    for g in grupos.values():
        extras.append({**g, "detalle": f"{g['detalle']} ({g['n']} pagos)"})
    return sorted(extras, key=lambda x: x["fecha"])


def build_tabla_egresos(items, titulo, extras=None, extra_note=""):
    # Separar items agrupados de individuales
    grupos: dict = {}
    individuales = []
    for it in items:
        g = it.get("grupo")
        if g:
            grupos.setdefault(g, []).append(it)
        else:
            individuales.append(it)

    rows = ""

    def _row(it, pagado_override=None):
        pagado = pagado_override if pagado_override is not None else (
            it["pagado_manual"] or it.get("pagado_auto", False))
        if it.get("es_contador"):
            if it.get("contador_completo"):
                badge  = '<span class="badge badge-ok">Pagado ✓</span>'
                pagado = True
            else:
                badge  = '<span class="badge badge-contador">Ver contador ↓</span>'
                pagado = False
        elif pagado:
            badge = '<span class="badge badge-ok">Pagado</span>'
            if it.get("pagado_auto"):
                m = it["match_notion"]
                badge += f'<div class="badge-sub">Notion: {m["fecha"]} · {fmt(m["monto"])}</div>'
        else:
            badge = '<span class="badge badge-pend">Pendiente</span>'
        return pagado, f"""
        <tr class="{'row-ok' if pagado else 'row-pend'}">
          <td>{it['dia']:02d}/{MES_ACTUAL:02d}</td>
          <td>{it['nombre']}</td>
          <td>{it['categoria']}</td>
          <td style="text-align:right;font-weight:600;">{fmt(it['monto'])}</td>
          <td>{badge}</td>
        </tr>"""

    pagado_total_items = 0
    for it in sorted(individuales, key=lambda x: x["dia"]):
        pag, r = _row(it)
        rows += r
        if pag:
            pagado_total_items += it["monto"]

    # Filas agrupadas (Operacional, Emporio): una fila resumen + detalle colapsable
    for grupo_nombre, g_items in sorted(grupos.items()):
        total_g   = sum(i["monto"] for i in g_items)
        pagados_g = sum(i["monto"] for i in g_items if i["pagado_manual"] or i.get("pagado_auto"))
        pend_g    = total_g - pagados_g
        pct_g     = pagados_g / total_g * 100 if total_g else 0
        dia_min   = min(i["dia"] for i in g_items)
        cat_g     = g_items[0]["categoria"]

        sub_rows = ""
        for si in sorted(g_items, key=lambda x: x["nombre"]):
            si_pag = si["pagado_manual"] or si.get("pagado_auto", False)
            sub_badge = (
                f'<span class="badge badge-ok" style="font-size:9px;">✓</span>' if si_pag
                else '<span class="badge badge-pend" style="font-size:9px;">Pend</span>'
            )
            sub_rows += f"""<tr style="background:#f9f9f9;font-size:12px;">
              <td style="padding-left:20px;">↳ {si['nombre']}</td>
              <td style="text-align:right;">{fmt(si['monto'])}</td>
              <td>{sub_badge}</td></tr>"""

        resumen_badge = (
            '<span class="badge badge-ok">Pagado</span>' if pend_g <= 0
            else f'<span class="badge badge-pend">{fmt(pagados_g)}/{fmt(total_g)}</span>'
        )
        rows += f"""
        <tr class="{'row-ok' if pend_g <= 0 else 'row-pend'}">
          <td>{dia_min:02d}/{MES_ACTUAL:02d}</td>
          <td colspan="2">
            <details style="display:inline;">
              <summary style="cursor:pointer;font-weight:600;">
                {grupo_nombre} ({len(g_items)} ítems)
              </summary>
              <table style="width:100%;margin-top:4px;">
                <tr><th style="text-align:left;font-size:11px;">Concepto</th>
                    <th style="text-align:right;font-size:11px;">Monto</th>
                    <th></th></tr>
                {sub_rows}
              </table>
            </details>
          </td>
          <td style="text-align:right;font-weight:600;">{fmt(total_g)}</td>
          <td>{resumen_badge}</td>
        </tr>"""
        pagado_total_items += pagados_g

    extras = extras or []
    pagado_extras = 0
    if extras:
        rows += f"""<tr><td colspan="5" style="background:#f0eeff;font-weight:700;font-size:11px;
          padding:6px 10px;color:#534AB7;">
          EGRESOS ADICIONALES DETECTADOS EN NOTION (no estaban en listado establecido)
        </td></tr>"""
        for r in sorted(extras, key=lambda x: x["fecha"]):
            rows += f"""
            <tr style="background:#f8f6ff;">
              <td>{r['fecha']}</td>
              <td>{r['detalle'] or '(sin detalle)'}</td>
              <td>{r['dept']}</td>
              <td style="text-align:right;font-weight:600;">{fmt(r['monto'])}</td>
              <td><span class="badge" style="background:#eee;color:#534AB7;">Solo Notion</span></td>
            </tr>"""

    total_extras = sum(r["monto"] for r in extras)
    total      = sum(i["monto"] for i in items) + total_extras
    pagado_tot = pagado_total_items + pagado_extras + total_extras
    pend_tot   = total - pagado_tot
    pct        = (pagado_tot / total * 100) if total else 0

    return f"""
    <div class="card">
      <h3>{titulo}</h3>
      {extra_note}
      <div class="progress-wrap">
        <div class="progress-bar"><div class="progress-fill" style="width:{pct:.1f}%;"></div></div>
        <div class="progress-lbl">{fmt(pagado_tot)} pagado de {fmt(total)} &nbsp;·&nbsp; {pct:.1f}%</div>
      </div>
      <div class="kpi-row">
        <div class="kpi"><div class="kpi-lbl">Total Egresos</div><div class="kpi-val">{fmt(total)}</div></div>
        <div class="kpi kpi-ok"><div class="kpi-lbl">Pagado</div><div class="kpi-val">{fmt(pagado_tot)}</div></div>
        <div class="kpi kpi-pend"><div class="kpi-lbl">Pendiente</div><div class="kpi-val">{fmt(pend_tot)}</div></div>
      </div>
      <table class="tbl">
        <tr><th>Fecha</th><th>Concepto</th><th>Categoría</th><th style="text-align:right;">Monto</th><th>Estado</th></tr>
        {rows}
      </table>
    </div>"""


def build_contadores(contadores):
    cards = ""
    for c in contadores:
        pct = (c["pagado"] / c["planificado"] * 100) if c["planificado"] else 0
        pct_show = min(pct, 100)
        pagos_html = "".join(
            f"<li>{p['fecha']} — {fmt(p['monto'])}</li>" for p in c["pagos"]
        ) or "<li>Sin transferencias registradas en Notion este mes</li>"
        color = "var(--success)" if c["pendiente"] <= 0 else "var(--plight)"
        cards += f"""
        <div class="contador">
          <div class="contador-head">
            <span>{c['label']}</span>
            <span>{fmt(c['pagado'])} / {fmt(c['planificado'])}</span>
          </div>
          <div class="progress-bar"><div class="progress-fill" style="width:{pct_show:.1f}%;background:{color};"></div></div>
          <div class="progress-lbl">
            {pct:.1f}% pagado &nbsp;·&nbsp; Pendiente: {fmt(max(c['pendiente'], 0))}
          </div>
          <details>
            <summary>Transferencias registradas ({len(c['pagos'])})</summary>
            <ul>{pagos_html}</ul>
          </details>
        </div>"""

    return f"""
    <div class="card">
      <h3>Contadores de pagos fraccionados</h3>
      <p class="note">
        Vish &amp; Gita y Esteban se pagan en varias transferencias durante el mes
        (registradas en Ingresos &amp; Egresos como Egreso / Administración a nombre
        de Diego Maldonado y Esteban Gauna). Aquí se compara el total planificado
        del mes vs lo efectivamente transferido a la fecha.
      </p>
      {cards}
    </div>"""


def build_pie_chart(items):
    por_categoria = defaultdict(float)
    for it in items:
        por_categoria[it["categoria"]] += it["monto"]

    labels = sorted(por_categoria, key=lambda k: -por_categoria[k])
    valores = [por_categoria[l] for l in labels]
    total = sum(valores)

    rows = ""
    for l, v in zip(labels, valores):
        pct = v / total * 100 if total else 0
        rows += f"""
        <tr>
          <td>{l}</td>
          <td style="text-align:right;">{fmt(v)}</td>
          <td style="text-align:right;">{pct:.1f}%</td>
        </tr>"""

    return f"""
    <div class="card">
      <h3>Distribución de Egresos por Categoría — {MES_ACTUAL_NOMBRE}</h3>
      <div style="display:flex;flex-wrap:wrap;gap:24px;align-items:center;">
        <div style="position:relative;height:260px;width:260px;flex:0 0 260px;">
          <canvas id="pieEgresos"></canvas>
        </div>
        <table class="tbl" style="flex:1;min-width:240px;">
          <tr><th>Categoría</th><th style="text-align:right;">Monto</th><th style="text-align:right;">%</th></tr>
          {rows}
        </table>
      </div>
    </div>
    <script>
    new Chart(document.getElementById('pieEgresos'), {{
      type: 'doughnut',
      data: {{
        labels: {labels!r},
        datasets: [{{
          data: {valores!r},
          backgroundColor: ['#2e5c3e','#4a7c59','#c9a96e','#1B4F91','#534AB7','#E67E22','#0F6E56','#c0392b','#6b8c78','#a5d6a7','#888'],
        }}]
      }},
      options: {{ plugins: {{ legend: {{ position: 'bottom', labels: {{ font: {{ size: 10 }} }} }} }} }}
    }});
    </script>"""


# ─────────────────────────────────────────────
# Metas de ingresos del mes (definidas por usuario)
# ─────────────────────────────────────────────
METAS_INGRESOS = {
    "Clases":     4_000_000,
    "Terapias":   1_500_000,
    "Educación":  1_500_000,
    "Arriendo":     800_000,
    "Finanzas":   2_000_000,
}

# ─────────────────────────────────────────────
# Posibles reducciones de costos (desde reunión)
# ─────────────────────────────────────────────
REDUCCIONES = [
    # (area, medida, ahorro_estimado, estado)
    ("Marketing",    "Reducir presupuesto mensual (350 → 200 mil)",          150_000,  "Decidido"),
    ("Tecnología",   "Sacar Manychat",                                        40_000,  "Decidido"),
    ("Tecnología",   "Evaluar cancelar 1 cuenta Zoom",                        15_400,  "Decidido"),
    ("Operacional",  "Reducir uso luz (parafinas)",                             None,  "Decidido"),
    ("Operacional",  "Eliminar jardinero",                                      None,  "Decidido"),
    ("Emporio",      "Vender muebles (ingreso puntual, no recurrente)",         None,  "Decidido"),
    ("Arriendo",     "Reducir precio de arriendo (por aprobar)",              500_000,  "En evaluación"),
    ("Clases",       "Sacar clase Nico 19:00 hrs",                            48_000,  "En evaluación"),
    ("Clases",       "Sacar clase Cata Suckel 18:30 hrs",                     48_000,  "Decidido"),
    ("Clases",       "Sacar clase Hatha 18:00 hrs",                           60_000,  "Decidido"),
    ("Clases",       "Evaluar clase Coni Muñoz miérc. 13 hrs → Gita",         60_000,  "Decidido"),
    ("Clases",       "Consuelo: sáb/lun a $15.000 en vez de $17.000",          None,  "Decidido"),
    ("Clases",       "Viernes: clase 1h30 yoga suave + sono (pago único 20-22 mil)", None, "Decidido"),
    ("Recepción",    "Turno Nayely viernes (-60 mil) / Eli cubre más tarde",   60_000,  "En evaluación"),
    ("Recepción",    "Revisar turno martes Nico",                               None,  "En evaluación"),
    ("Recepción",    "Disminuir horas Eli",                                     None,  "En evaluación"),
    ("Educación",    "Automatización boletas → eliminar pago Amanda 40 mil",   40_000,  "Decidido"),
]
# Reducciones ya decididas (se restan del costo estimado) vs en evaluacion
# (se muestran como potencial adicional, sin restar todavia)
REDUCCIONES_DECIDIDAS_POR_AREA = defaultdict(int)
REDUCCIONES_EVALUACION_POR_AREA = defaultdict(int)
for _area, _medida, _ahorro, _estado in REDUCCIONES:
    if not _ahorro:
        continue
    if _estado == "Decidido":
        REDUCCIONES_DECIDIDAS_POR_AREA[_area] += _ahorro
    else:
        REDUCCIONES_EVALUACION_POR_AREA[_area] += _ahorro


def _area_de(item):
    return AREA_MAP.get(item["nombre"], AREA_DEFAULT)


def build_proyeccion_julio(egresos_mes_actual, registros_notion_mes_actual):
    """Sección 1: Meta de ingresos vs costos estimados de Julio (mismo listado
    de Junio sin pagos puntuales, + items nuevos), agrupados por area, con las
    reducciones de costos aplicadas para ver el escenario ajustado."""
    items_por_area = defaultdict(list)
    for it in egresos_mes_actual:
        if it["nombre"] in ITEMS_NO_RECURRENTES:
            continue
        monto = OVERRIDES_MES_PROX.get(it["nombre"], it["monto"])
        items_por_area[_area_de(it)].append((it["nombre"], monto))
    for extra in COSTOS_EXTRA_MES_PROX:
        items_por_area[extra["area"]].append((extra["nombre"], extra["monto"]))

    costo_base_area = {area: sum(m for _, m in items) for area, items in items_por_area.items()}

    total_meta = sum(METAS_INGRESOS.values())
    total_costo_base = sum(costo_base_area.values())
    total_decidida = sum(REDUCCIONES_DECIDIDAS_POR_AREA.values())
    total_evaluacion = sum(REDUCCIONES_EVALUACION_POR_AREA.values())
    total_costo_ajustado = total_costo_base - total_decidida

    neto_base = total_meta - total_costo_base
    neto_ajustado = total_meta - total_costo_ajustado

    rows = ""
    for area in sorted(costo_base_area, key=lambda a: -costo_base_area[a]):
        base = costo_base_area[area]
        decidida = REDUCCIONES_DECIDIDAS_POR_AREA.get(area, 0)
        evaluacion = REDUCCIONES_EVALUACION_POR_AREA.get(area, 0)
        ajustado = base - decidida
        rows += f"""
        <tr>
          <td style="font-weight:600;">{area}</td>
          <td style="text-align:right;">{fmt(base)}</td>
          <td style="text-align:right;color:{'#0F6E56' if decidida else 'var(--muted)'};">
            {'-' + fmt(decidida) if decidida else '—'}
          </td>
          <td style="text-align:right;font-weight:700;">{fmt(ajustado)}</td>
          <td style="text-align:right;color:{'#E67E22' if evaluacion else 'var(--muted)'};">
            {'-' + fmt(evaluacion) if evaluacion else '—'}
          </td>
        </tr>"""

    color_base = '#0F6E56' if neto_base >= 0 else '#c0392b'
    color_ajus = '#0F6E56' if neto_ajustado >= 0 else '#c0392b'

    # Listado detallado: item + monto, agrupado por area, para poder ajustar/corregir
    detalle_html = ""
    for area in sorted(items_por_area, key=lambda a: -costo_base_area[a]):
        items_rows = "".join(
            f"<tr><td>{nombre}</td><td style='text-align:right;'>{fmt(monto)}</td></tr>"
            for nombre, monto in sorted(items_por_area[area], key=lambda x: -x[1])
        )
        detalle_html += f"""
        <details class="area-detalle">
          <summary>{area} — <b>{fmt(costo_base_area[area])}</b> ({len(items_por_area[area])} ítems)</summary>
          <table class="tbl" style="margin-top:6px;">
            <tr><th>Concepto</th><th style="text-align:right;">Monto</th></tr>
            {items_rows}
          </table>
        </details>"""

    # Detalle de ingresos: transacciones reales de Notion del mes en curso,
    # agrupadas por area (mismo formato que el detalle de egresos)
    ingresos_por_area = defaultdict(list)
    for r in registros_notion_mes_actual:
        if r["movimiento"] == "Ingreso":
            ingresos_por_area[r["dept"]].append((r["detalle"] or "(sin detalle)", r["monto"], r["fecha"]))

    ingreso_real_area = {area: sum(m for _, m, _ in items) for area, items in ingresos_por_area.items()}

    detalle_ingresos_html = ""
    for area in sorted(ingresos_por_area, key=lambda a: -ingreso_real_area[a]):
        meta_area = METAS_INGRESOS.get(area)
        meta_txt = f" / meta {fmt(meta_area)}" if meta_area else " (sin meta asignada)"
        items_rows = "".join(
            f"<tr><td>{fecha}</td><td>{detalle}</td><td style='text-align:right;'>{fmt(monto)}</td></tr>"
            for detalle, monto, fecha in sorted(ingresos_por_area[area], key=lambda x: x[2])
        )
        detalle_ingresos_html += f"""
        <details class="area-detalle">
          <summary>{area} — <b>{fmt(ingreso_real_area[area])}</b>{meta_txt} ({len(ingresos_por_area[area])} ítems)</summary>
          <table class="tbl" style="margin-top:6px;">
            <tr><th>Fecha</th><th>Concepto</th><th style="text-align:right;">Monto</th></tr>
            {items_rows}
          </table>
        </details>"""

    return f"""
    <div class="card">
      <h3>Proyección {MES_PROX_NOMBRE} — Meta de Ingresos vs Costos Estimados</h3>
      <p class="note">
        Costos estimados = listado de pagos de {MES_ACTUAL_NOMBRE} (excluyendo pagos puntuales como
        la garantía de arriendo) + ítems nuevos (aire acondicionado $50.000, deuda Emporio $100.000),
        agrupados por área. "Reducción Decidida" ya se resta del costo ajustado; "En Evaluación"
        es ahorro potencial adicional que todavía no se aplica (ej. baja de arriendo $500.000 por aprobar).
      </p>
      <div class="kpi-row">
        <div class="kpi kpi-ok"><div class="kpi-lbl">Meta Ingresos</div><div class="kpi-val">{fmt(total_meta)}</div></div>
        <div class="kpi kpi-pend"><div class="kpi-lbl">Costo Estimado Base</div><div class="kpi-val">{fmt(total_costo_base)}</div></div>
        <div class="kpi"><div class="kpi-lbl">Neto Base</div><div class="kpi-val" style="color:{color_base};">{fmt(neto_base)}</div></div>
        <div class="kpi kpi-pend"><div class="kpi-lbl">Costo Ajustado (decidido)</div><div class="kpi-val">{fmt(total_costo_ajustado)}</div></div>
        <div class="kpi"><div class="kpi-lbl">Neto Ajustado</div><div class="kpi-val" style="color:{color_ajus};">{fmt(neto_ajustado)}</div></div>
      </div>
      <table class="tbl">
        <tr><th>Área</th><th style="text-align:right;">Costo Base</th><th style="text-align:right;">Reducción Decidida</th><th style="text-align:right;">Costo Ajustado</th><th style="text-align:right;">En Evaluación (adicional)</th></tr>
        {rows}
        <tr style="background:#f0f4f0;font-weight:700;">
          <td>Total</td>
          <td style="text-align:right;">{fmt(total_costo_base)}</td>
          <td style="text-align:right;color:#0F6E56;">-{fmt(total_decidida)}</td>
          <td style="text-align:right;">{fmt(total_costo_ajustado)}</td>
          <td style="text-align:right;color:#E67E22;">-{fmt(total_evaluacion)}</td>
        </tr>
      </table>

      <h4>Detalle de costos por área (para ajustar y corregir)</h4>
      {detalle_html}

      <h4>Detalle de ingresos por área — {MES_ACTUAL_NOMBRE} (real, Notion)</h4>
      {detalle_ingresos_html}
    </div>"""


def build_basal_vs_real(egresos_mes_actual, registros_notion):
    """Compara el listado planificado (basal, agrupado por area) contra los
    egresos reales registrados en Notion este mes (agrupados por su propio
    departamento), para ir viendo como se ajusta la realidad al plan."""
    basal_area = defaultdict(float)
    for it in egresos_mes_actual:
        basal_area[_area_de(it)] += it["monto"]

    real_area = defaultdict(float)
    for r in registros_notion:
        if r["movimiento"] == "Egreso":
            real_area[r["dept"]] += r["monto"]

    areas = sorted(set(basal_area) | set(real_area), key=lambda a: -basal_area.get(a, 0))
    rows = ""
    for area in areas:
        base = basal_area.get(area, 0)
        real = real_area.get(area, 0)
        diff = real - base
        diff_color = "#c0392b" if diff > 0 else "#0F6E56"
        rows += f"""
        <tr>
          <td style="font-weight:600;">{area}</td>
          <td style="text-align:right;">{fmt(base)}</td>
          <td style="text-align:right;">{fmt(real)}</td>
          <td style="text-align:right;color:{diff_color};font-weight:600;">{'+' if diff>=0 else ''}{fmt(diff)}</td>
        </tr>"""

    total_base = sum(basal_area.values())
    total_real = sum(real_area.values())

    return f"""
    <div class="card">
      <h3>Basal vs Real — {MES_ACTUAL_NOMBRE}</h3>
      <p class="note">
        Basal = listado planificado de pagos del mes (sheet), agrupado por área.
        Real = egresos efectivamente registrados en la base de Ingresos &amp; Egresos (Notion),
        por departamento. Permite ver dónde la realidad se está desviando del plan.
      </p>
      <div class="kpi-row">
        <div class="kpi kpi-pend"><div class="kpi-lbl">Total Basal</div><div class="kpi-val">{fmt(total_base)}</div></div>
        <div class="kpi"><div class="kpi-lbl">Total Real</div><div class="kpi-val">{fmt(total_real)}</div></div>
      </div>
      <table class="tbl">
        <tr><th>Área</th><th style="text-align:right;">Basal</th><th style="text-align:right;">Real</th><th style="text-align:right;">Diferencia</th></tr>
        {rows}
      </table>
    </div>"""


def build_deudas_inversiones():
    """Sección 5: deudas pendientes consolidadas + tabla de creditos/inversiones
    desde la base Notion "Inversiones - Creditos - Deudas"."""
    rows_deudas = ""
    total_deudas = 0
    for d in DEUDAS_PENDIENTES:
        total_deudas += d["monto"]
        desde = f"pendiente desde {d['desde']}" if d["desde"] else "sin fecha"
        rows_deudas += f"""
        <tr>
          <td>{d['nombre']}</td>
          <td>{d['categoria']}</td>
          <td style="text-align:right;font-weight:600;">{fmt(d['monto'])}</td>
          <td style="font-size:12px;color:var(--muted);">{desde}</td>
        </tr>"""

    total_sii = sum(v for _, v in SII_HISTORICO)
    rows_sii = "".join(
        f"<tr><td>{mes}</td><td style='text-align:right;'>{fmt(v)}</td></tr>"
        for mes, v in SII_HISTORICO
    )

    rows_inv = ""
    total_inv_recibido = 0
    total_a_devolver = 0
    for inv in INVERSIONES_NOTION:
        total_inv_recibido += inv["valor"] or 0
        total_a_devolver += inv["valor_devolucion"] or 0
        estado_color = {
            "Inversión recibida": "#1B4F91", "Pagando": "#E67E22",
            "Postergada": "#c0392b", "Por postergar": "#c0392b", "Pago completado": "#0F6E56",
        }.get(inv["estado"], "#888")
        rows_inv += f"""
        <tr>
          <td>{inv['nombre']}</td>
          <td>{inv['tipo'] or '—'}</td>
          <td style="text-align:right;">{fmt(inv['valor']) if inv['valor'] else '—'}</td>
          <td style="text-align:right;">{fmt(inv['valor_devolucion']) if inv['valor_devolucion'] else '—'}</td>
          <td>{inv['fecha_dev'] or '—'}</td>
          <td><span class="badge" style="background:#eee;color:{estado_color};">{inv['estado']}</span></td>
          <td style="font-size:11px;color:var(--muted);">{inv['obs'] or ''}</td>
        </tr>"""

    return f"""
    <div class="card">
      <h3>Deudas Pendientes</h3>
      <p class="note">Deudas detectadas en Mayo que no quedaron incluidas en el listado mensual, más otras consolidadas.</p>
      <div class="kpi-row">
        <div class="kpi kpi-pend"><div class="kpi-lbl">Total Deudas</div><div class="kpi-val">{fmt(total_deudas)}</div></div>
        <div class="kpi kpi-pend"><div class="kpi-lbl">SII Histórico</div><div class="kpi-val">{fmt(total_sii)}</div></div>
      </div>
      <table class="tbl">
        <tr><th>Concepto</th><th>Categoría</th><th style="text-align:right;">Monto</th><th>Estado</th></tr>
        {rows_deudas}
      </table>
      <h4>SII — histórico mensual</h4>
      <table class="tbl">
        <tr><th>Mes</th><th style="text-align:right;">Monto</th></tr>
        {rows_sii}
        <tr style="background:#f0f4f0;font-weight:700;"><td>Total</td><td style="text-align:right;">{fmt(total_sii)}</td></tr>
      </table>
    </div>
    <div class="card">
      <h3>Inversiones, Créditos y Préstamos</h3>
      <p class="note">
        Fuente: <a href="{INVERSIONES_DB_URL}" target="_blank">base Notion "Inversiones - Créditos - Deudas"</a>.
        Incluye el nuevo préstamo de Paty ($3.000.000, devolución desde enero 2027).
      </p>
      <div class="kpi-row">
        <div class="kpi"><div class="kpi-lbl">Total Recibido</div><div class="kpi-val">{fmt(total_inv_recibido)}</div></div>
        <div class="kpi kpi-pend"><div class="kpi-lbl">Total a Devolver</div><div class="kpi-val">{fmt(total_a_devolver)}</div></div>
      </div>
      <table class="tbl">
        <tr><th>Nombre</th><th>Tipo</th><th style="text-align:right;">Valor</th><th style="text-align:right;">V. Devolución</th><th>Fecha Devolución</th><th>Estado</th><th>Obs.</th></tr>
        {rows_inv}
      </table>
    </div>"""


def build_reducciones():
    por_area = defaultdict(list)
    for r in REDUCCIONES:
        por_area[r[0]].append(r)

    ahorro_decidido = sum(r[2] for r in REDUCCIONES if r[3] == "Decidido" and r[2])
    ahorro_posible  = sum(r[2] for r in REDUCCIONES if r[2])

    rows = ""
    for area in ["Marketing","Tecnología","Operacional","Emporio","Clases","Recepción","Formación"]:
        items_area = por_area.get(area, [])
        if not items_area:
            continue
        rowspan = len(items_area)
        for i, (_, medida, ahorro, estado) in enumerate(items_area):
            color_estado = "#0F6E56" if estado == "Decidido" else "#E67E22"
            ahorro_txt = fmt(ahorro) if ahorro else "Por definir"
            td_area = f'<td rowspan="{rowspan}" style="font-weight:700;vertical-align:top;padding-top:8px;">{area}</td>' if i == 0 else ""
            rows += f"""
            <tr class="{'row-ok' if estado=='Decidido' else 'row-pend'}">
              {td_area}
              <td>{medida}</td>
              <td style="text-align:right;font-weight:600;color:{color_estado};">{ahorro_txt}</td>
              <td><span class="badge {'badge-ok' if estado=='Decidido' else 'badge-pend'}">{estado}</span></td>
            </tr>"""

    return f"""
    <div class="card">
      <h3>Posibles Reducciones de Costos</h3>
      <p class="note">
        Medidas identificadas para reducir egresos.
        Ahorros confirmados: <b>{fmt(ahorro_decidido)}/mes</b> &nbsp;·&nbsp;
        Total potencial (incluye en evaluación): <b>{fmt(ahorro_posible)}/mes</b>
        (más items "por definir" en operacional, clases y recepción — cifra imagen: 570 mil sin Eli/Nico, 760 mil con).
      </p>
      <table class="tbl">
        <tr><th>Área</th><th>Medida</th><th style="text-align:right;">Ahorro Est.</th><th>Estado</th></tr>
        {rows}
        <tr style="background:#f0f4f0;font-weight:700;">
          <td colspan="2">Total ahorros cuantificados (decididos)</td>
          <td style="text-align:right;">{fmt(ahorro_decidido)}</td>
          <td></td>
        </tr>
        <tr style="background:#e8f5ec;font-weight:700;">
          <td colspan="2">Total potencial (todos los cuantificados)</td>
          <td style="text-align:right;">{fmt(ahorro_posible)}</td>
          <td></td>
        </tr>
      </table>
    </div>"""


HTML_HEAD = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Dhatri — Dashboard Finanzas</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
:root {
  --bg: #f0f4f0; --card: #ffffff; --primary: #2e5c3e; --plight: #4a7c59;
  --accent: #c9a96e; --text: #1a2e24; --muted: #6b8c78; --border: #dde8dc;
  --danger: #c0392b; --success: #0F6E56; --warn: #E67E22; --shadow: 0 2px 14px rgba(0,0,0,.07);
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--text)}
header{background:linear-gradient(135deg,var(--primary) 0%,#1a3a28 100%);color:#fff;padding:20px 32px;
  display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 10px rgba(0,0,0,.25);}
.logo{font-size:22px;font-weight:700;letter-spacing:.5px}
.logo span{color:var(--accent)}
.meta{text-align:right;font-size:12px;opacity:.85}
.container{max-width:1100px;margin:0 auto;padding:24px}
.card{background:var(--card);border-radius:12px;padding:22px;box-shadow:var(--shadow);margin-bottom:22px}
.card h3{font-size:16px;margin-bottom:10px;color:var(--primary)}
.card h4{font-size:13px;margin:18px 0 8px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px}
.note{font-size:12px;color:var(--muted);margin-bottom:14px;line-height:1.5}
.kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin:14px 0}
.kpi{background:#f7faf7;border-radius:10px;padding:14px 16px;border-top:4px solid var(--primary)}
.kpi-ok{border-top-color:var(--success)}
.kpi-pend{border-top-color:var(--warn)}
.kpi-lbl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--muted)}
.kpi-val{font-size:24px;font-weight:800;margin-top:4px}
.progress-wrap{margin:10px 0 18px}
.progress-bar{background:#e0e0e0;border-radius:6px;height:14px;overflow:hidden}
.progress-fill{background:linear-gradient(90deg,var(--plight),var(--success));height:100%;}
.progress-lbl{font-size:12px;color:var(--muted);margin-top:6px}
.tbl{width:100%;border-collapse:collapse;font-size:13px;margin-top:6px}
.tbl th{background:var(--primary);color:#fff;padding:8px 10px;text-align:left;font-weight:600;font-size:12px}
.tbl td{padding:7px 10px;border-bottom:1px solid var(--border)}
.row-ok{background:#f3faf6}
.row-pend{background:#fffaf2}
.badge{font-size:10px;font-weight:700;padding:3px 8px;border-radius:20px;text-transform:uppercase;display:inline-block}
.badge-ok{background:#e3f7ed;color:var(--success)}
.badge-pend{background:#fdf0e3;color:var(--warn)}
.badge-growth{background:#e8f5ec;color:var(--primary);margin-left:4px}
.badge-contador{background:#eee8fb;color:#534AB7}
.badge-sub{font-size:10px;color:var(--muted);margin-top:2px}
.alert{background:#fdf0e3;border-left:4px solid var(--warn);padding:12px 16px;border-radius:6px;font-size:13px;margin-bottom:14px}
.alert b{color:var(--warn)}
.contador{border:1px solid var(--border);border-radius:10px;padding:12px 14px;margin-bottom:12px}
.contador-head{display:flex;justify-content:space-between;font-weight:700;font-size:13px;margin-bottom:6px}
.contador details{margin-top:8px;font-size:12px;color:var(--muted)}
.contador summary{cursor:pointer;font-weight:600}
.contador ul{margin:6px 0 0 18px}
.area-detalle{border:1px solid var(--border);border-radius:8px;padding:8px 12px;margin-bottom:8px}
.area-detalle summary{cursor:pointer;font-size:13px;font-weight:600;color:var(--primary)}
</style>
</head>
<body>
<header>
  <div class="logo">DHATRI <span>· Finanzas</span></div>
  <div class="meta"><div>%FECHA%</div><div>%MES_ACTUAL%</div></div>
</header>
<div class="container">
"""

HTML_TAIL = """
</div>
</body>
</html>
"""


def main():
    print("Leyendo Excel (base Junio para items recurrentes)...")
    egresos_junio, _egresos_mayo = cargar_excel()

    # Construir lista de Julio: items recurrentes de Junio + personal actualizado
    egresos_julio = _build_egresos_julio(egresos_junio)

    print("Consultando Notion (Julio 2026)...")
    julio_notion = query_notion_month(ANIO_ACTUAL, MES_ACTUAL)

    egresos_julio = marcar_pagados(egresos_julio, julio_notion)
    contadores    = calcular_contadores(egresos_julio, julio_notion)

    # Si un contador está completamente pagado → marcarlo en la tabla principal
    ctr_by_sheet = {c["nombre_sheet"]: c for c in contadores}
    for it in egresos_julio:
        if it.get("es_contador"):
            ctr = ctr_by_sheet.get(it["nombre"])
            if ctr and ctr["pendiente"] <= 0:
                it["contador_completo"] = True

    extras_notion = _find_extras_notion(egresos_julio, julio_notion)

    html = HTML_HEAD.replace("%FECHA%", datetime.now().strftime("%d/%m/%Y %H:%M"))
    html = html.replace("%MES_ACTUAL%", MES_ACTUAL_NOMBRE)

    # 1. Proyeccion mes siguiente: meta de ingresos vs costos estimados
    html += build_proyeccion_julio(egresos_julio, julio_notion)
    html += build_reducciones()
    # 2. Egresos a pagar del mes en curso: basal vs real + detalle por item
    html += build_basal_vs_real(egresos_julio, julio_notion)
    html += build_tabla_egresos(
        egresos_julio,
        f"Egresos a Pagar — {MES_ACTUAL_NOMBRE}",
        extras=extras_notion,
    )
    # 3. Contadores de pagos fraccionados
    html += build_contadores(contadores)
    # 4. Distribucion de egresos por categoria
    html += build_pie_chart(egresos_julio)
    # 5. Deudas e inversiones
    html += build_deudas_inversiones()
    html += HTML_TAIL

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"OK -> {OUT_PATH}")


if __name__ == "__main__":
    main()
