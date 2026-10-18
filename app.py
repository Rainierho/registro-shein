# -*- coding: utf-8 -*-
"""
Registro SHEIN — app Streamlit
Backends: Google Sheets (produccion en Streamlit Cloud) o Excel local (desarrollo).
"""
import datetime as dt
from zoneinfo import ZoneInfo
import io
import re
import os

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Registro SHEIN", page_icon="💍", layout="wide")

# --------------------------------------------------------------------------
# Configuracion
# --------------------------------------------------------------------------
TABS = {
    "INVENTARIO": ["DESCRIPCION", "NOMBRE", "CANTIDAD", "PRECIO_COMPRA", "PRECIO_VENTA", "LOTE",
                   "FOTO", "OCULTAR"],
    "VENTAS": ["NOMBRE", "MONTO", "FECHA", "LOTE"],
    "DEUDAS": ["DEUDOR", "MONTO", "BOLOS", "NOTA"],
    "CAJA": ["UBICACION", "MONTO", "MONEDA", "TASA"],
    "ENCARGOS": ["PERSONA", "TIPO", "DESCRIPCION", "COSTO", "COBRAR", "ABONADO", "MONEDA",
                 "FECHA", "ULT_ABONO", "NOTA"],
    "CONFIG": ["CLAVE", "VALOR"],
    "GASTOS": ["FECHA", "CATEGORIA", "CONCEPTO", "LOTE", "MONTO", "MONEDA", "UBICACION"],
    "MOVIMIENTOS": ["FECHA", "UBICACION", "MONTO", "MONEDA", "SALDO", "ORIGEN", "DETALLE"],
    "CONTACTOS": ["PERSONA", "TELEFONO"],
    "LOTES": ["LOTE", "ESTADO", "LLEGADA"],
    "VENTAS_HIST": ["NOMBRE", "MONTO", "FECHA", "LOTE"],
    "INVENTARIO_HIST": ["NOMBRE", "SKU", "PROPIETARIO", "TALLA", "CANTIDAD", "PRECIO_COMPRA",
                        "PRECIO_VENTA", "TIPO", "CATEGORIA", "LOTE", "FECHA_LOTE"],
}
# pestañas de solo lectura: nunca se escriben desde la app
# nunca se escriben desde la app
SOLO_LECTURA = ("VENTAS_HIST", "INVENTARIO_HIST", "DEUDAS")
NO_CREAR = ("VENTAS_HIST", "INVENTARIO_HIST")   # no crear la pestaña si no existe
LOTE_CORTE = 10          # lotes 1-10 son del archivo viejo; en el actual son arrastre
NUMERICAS = {
    "INVENTARIO": ["CANTIDAD", "PRECIO_COMPRA", "PRECIO_VENTA", "LOTE"],
    "VENTAS": ["MONTO", "LOTE"],
    "DEUDAS": ["MONTO", "BOLOS"],
    "CAJA": ["MONTO", "TASA"],
    "ENCARGOS": ["COSTO", "COBRAR", "ABONADO"],
    "GASTOS": ["MONTO", "LOTE"],
    "MOVIMIENTOS": ["MONTO", "SALDO"],
    "LOTES": ["LOTE"],
    "VENTAS_HIST": ["MONTO", "LOTE"],
    "INVENTARIO_HIST": ["CANTIDAD", "PRECIO_COMPRA", "PRECIO_VENTA", "LOTE"],
}
# columnas de plata: siempre float, para que el editor acepte decimales
DINERO = {"MONTO", "PRECIO_COMPRA", "PRECIO_VENTA", "BOLOS", "TASA", "COSTO", "COBRAR",
          "ABONADO", "SALDO"}
FECHAS = {"VENTAS": ["FECHA"], "ENCARGOS": ["FECHA", "ULT_ABONO"], "GASTOS": ["FECHA"],
          "MOVIMIENTOS": ["FECHA"], "LOTES": ["LLEGADA"], "VENTAS_HIST": ["FECHA"],
          "INVENTARIO_HIST": ["FECHA_LOTE"]}
CLAVES = {"INVENTARIO": "NOMBRE", "VENTAS": "NOMBRE", "DEUDAS": "DEUDOR", "CAJA": "UBICACION",
          "ENCARGOS": "PERSONA", "CONFIG": "CLAVE", "GASTOS": "CATEGORIA",
          "MOVIMIENTOS": "UBICACION", "CONTACTOS": "PERSONA", "LOTES": "LOTE",
          "VENTAS_HIST": "NOMBRE",
          "INVENTARIO_HIST": "NOMBRE"}
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


RUTA_SECRETS = os.path.join(".streamlit", "secrets.toml")
FALLA_SECRETS = None


def secreto(clave, default=None):
    global FALLA_SECRETS
    try:
        return st.secrets[clave]
    except KeyError:
        return default
    except Exception as e:
        if os.path.exists(RUTA_SECRETS):
            FALLA_SECRETS = str(e)
        return default


BACKEND = secreto("backend", "excel")
EXCEL_PATH = secreto("excel_path", "REGISTRO_SHEIN.xlsx")


# --------------------------------------------------------------------------
# Capa de datos
# --------------------------------------------------------------------------
EPOCA_SHEETS = dt.datetime(1899, 12, 30)
ALIAS = {"E": "DESCRIPCION", "UBICACIÓN": "UBICACION", "DEUDOR ": "DEUDOR"}


def _a_numero(serie):
    """Numeros nativos, o texto con coma o punto decimal ('1,77', '1.234,56', '$5')."""
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce")

    def limpiar(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        if isinstance(v, (int, float)):
            return float(v)
        t = str(v).strip()
        if t in ("", "-", "nan", "None"):
            return None
        t = "".join(ch for ch in t if ch.isdigit() or ch in ".,-")
        if "," in t and "." in t:            # el ultimo separador es el decimal
            if t.rfind(",") > t.rfind("."):
                t = t.replace(".", "").replace(",", ".")
            else:
                t = t.replace(",", "")
        elif "," in t:
            t = t.replace(",", ".")
        try:
            return float(t)
        except ValueError:
            return None

    return pd.to_numeric(serie.map(limpiar), errors="coerce")


def _a_fecha(serie):
    """Seriales de Sheets (dias desde 1899-12-30) o texto en formato dd/mm/aaaa."""
    if pd.api.types.is_numeric_dtype(serie):
        return EPOCA_SHEETS + pd.to_timedelta(pd.to_numeric(serie, errors="coerce"), unit="D")

    def conv(v):
        # cada valor según su propio formato: pandas, si se le pasa la columna entera,
        # adivina el formato por la primera fila y puede invertir día y mes en todas
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return pd.NaT
        if isinstance(v, (pd.Timestamp, dt.datetime, dt.date)):
            return pd.Timestamp(v)
        if isinstance(v, (int, float)):
            return pd.Timestamp(EPOCA_SHEETS + dt.timedelta(days=float(v)))
        s = str(v).strip()
        if not s or s.lower() in ("nan", "nat", "none"):
            return pd.NaT
        if re.match(r"^\d{4}-\d{1,2}-\d{1,2}", s):            # 2026-07-01 [13:34]
            return pd.to_datetime(s, format="ISO8601", errors="coerce")
        return pd.to_datetime(s, dayfirst=True, errors="coerce")  # 01/07/2026

    return pd.to_datetime(serie.map(conv), errors="coerce")


def _matriz_a_df(valores):
    """Matriz cruda de Sheets -> DataFrame, rellenando filas cortas."""
    if not valores:
        return pd.DataFrame()
    ancho = max(len(f) for f in valores)
    filas = [list(f) + [""] * (ancho - len(f)) for f in valores]
    cabecera = [ALIAS.get(str(c).strip().upper(), str(c).strip().upper()) for c in filas[0]]
    cabecera = [c if c else f"COL_{i}" for i, c in enumerate(cabecera)]
    df = pd.DataFrame(filas[1:], columns=cabecera)
    return df.loc[:, ~df.columns.duplicated()]


def _normalizar(df, tab):
    cols = TABS[tab]
    for c in cols:
        if c not in df.columns:
            df[c] = None
    df = df[cols].copy()
    for c in NUMERICAS.get(tab, []):
        df[c] = _a_numero(df[c])
        if c in DINERO:
            df[c] = df[c].astype("float64")
    for c in FECHAS.get(tab, []):
        df[c] = _a_fecha(df[c])
    for c in cols:
        if c not in NUMERICAS.get(tab, []) and c not in FECHAS.get(tab, []):
            df[c] = df[c].fillna("").astype(str).str.strip().replace({"nan": "", "None": ""})
    clave = CLAVES[tab]
    df = df[df[clave].fillna("").astype(str).str.strip() != ""]
    return df.reset_index(drop=True)


def _a_celdas(df):
    """DataFrame -> matriz de valores serializables para Sheets."""
    salida = [list(df.columns)]
    for _, fila in df.iterrows():
        linea = []
        for v in fila:
            if pd.isna(v):
                linea.append("")
            elif isinstance(v, (pd.Timestamp, dt.datetime, dt.date)):
                ts = pd.Timestamp(v)
                con_hora = ts.hour or ts.minute
                linea.append(ts.strftime("%Y-%m-%d %H:%M" if con_hora else "%Y-%m-%d"))
            elif isinstance(v, (int, float)):
                linea.append(float(v))
            else:
                linea.append(str(v))
        salida.append(linea)
    return salida


@st.cache_resource(show_spinner=False)
def _libro():
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES
    )
    return gspread.authorize(creds).open_by_key(secreto("spreadsheet_id"))


def _hoja(tab, crear=True):
    import gspread

    libro = _libro()
    try:
        return libro.worksheet(tab)
    except gspread.WorksheetNotFound:
        if not crear:
            return None
        ws = libro.add_worksheet(tab, rows=2000, cols=max(12, len(TABS[tab])))
        ws.update(values=[TABS[tab]], range_name="A1")
        return ws


@st.cache_data(ttl=120, show_spinner="Cargando datos...")
def leer(tab):
    if BACKEND == "gsheets":
        ws = _hoja(tab, crear=tab not in NO_CREAR)
        valores = ws.get_values(value_render_option="UNFORMATTED_VALUE") if ws else []
        df = _matriz_a_df(valores)
        if df.empty:
            df = pd.DataFrame(columns=TABS[tab])
    else:
        try:
            df = pd.read_excel(EXCEL_PATH, sheet_name=tab)
        except Exception:
            df = pd.DataFrame(columns=TABS[tab])
        df.columns = [str(c).strip().upper() for c in df.columns]
    return _normalizar(df, tab)


def guardar(tab, df):
    if tab in SOLO_LECTURA:
        raise ValueError(f"{tab} es solo lectura")
    df = _normalizar(df, tab)
    if BACKEND == "gsheets":
        ws = _hoja(tab)
        ws.clear()
        ws.update(values=_a_celdas(df), range_name="A1", value_input_option="RAW")
    else:
        hojas = {t: leer(t) for t in TABS}
        hojas[tab] = df
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as w:
            for t, d in hojas.items():
                d.to_excel(w, sheet_name=t, index=False)
    st.cache_data.clear()


def anexar(tab, filas):
    """Añade filas al final sin reescribir la pestaña. Si la hoja tiene un encabezado
    viejo (columnas de menos), cae a reescritura completa, que pone el encabezado nuevo."""
    if tab in SOLO_LECTURA:
        raise ValueError(f"{tab} es solo lectura")
    nuevas = _normalizar(pd.DataFrame(filas), tab)
    if nuevas.empty:
        return
    if BACKEND == "gsheets":
        ws = _hoja(tab)
        cabecera = [str(c).strip().upper() for c in ws.row_values(1)]
        if cabecera == TABS[tab]:
            ws.append_rows(_a_celdas(nuevas)[1:], value_input_option="RAW")
            st.cache_data.clear()
            return
    df = pd.concat([leer(tab), nuevas], ignore_index=True)
    guardar(tab, df)


def agregar(tab, fila: dict):
    anexar(tab, [fila])


@st.cache_data(ttl=3600, show_spinner=False)
def tasa_bcv():
    """Dólar oficial BCV. Devuelve (tasa, vigencia, fuente) o (None, None, motivo)."""
    import requests
    try:
        r = requests.get("https://ve.dolarapi.com/v1/dolares/oficial", timeout=8)
        r.raise_for_status()
        d = r.json()
        valor = _num(d.get("promedio"), None)
        if valor and valor > 0:
            vig = str(d.get("fechaActualizacion") or "")[:10]
            return valor, vig, "BCV vía dolarapi"
    except Exception as e:
        motivo = f"dolarapi: {type(e).__name__}"
    else:
        motivo = "dolarapi: respuesta sin tasa"
    try:
        r = requests.get("https://www.bcv.org.ve/", timeout=8)
        m = re.search(r'id="dolar".*?<strong>\s*([\d.,]+)\s*</strong>', r.text, re.S)
        valor = leer_monto(m.group(1)) if m else None
        if valor and valor > 0:
            return valor, "", "BCV"
    except Exception as e:
        motivo += f" · bcv.org.ve: {type(e).__name__}"
    return None, None, motivo


@st.cache_data(ttl=900, show_spinner=False)
def tasa_usdt():
    """Bs por USDT en Binance P2P: mediana de los 10 primeros anuncios de compra."""
    import statistics
    import requests
    try:
        r = requests.post(
            "https://p2p.binance.com/bapi/c2c/v2/friendly/c2c/adv/search",
            json={"asset": "USDT", "fiat": "VES", "tradeType": "BUY", "page": 1, "rows": 10,
                  "payTypes": [], "publisherType": None},
            headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        precios = [float(a["adv"]["price"]) for a in r.json().get("data", [])]
        return statistics.median(precios) if precios else None
    except Exception:
        return None


RESPALDO_TABS = ["INVENTARIO", "VENTAS", "ENCARGOS", "CAJA", "GASTOS", "MOVIMIENTOS", "LOTES",
                 "CONTACTOS", "CONFIG"]
RESPALDO_DIAS = 30


def hacer_respaldo():
    """Copia las pestañas operativas a una pestaña con la fecha de hoy en la hoja de
    respaldos, y borra las de más de RESPALDO_DIAS días. Devuelve texto de resultado."""
    import gspread
    from google.oauth2.service_account import Credentials

    id_resp = secreto("backup_spreadsheet_id")
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES)
    libro_r = gspread.authorize(creds).open_by_key(id_resp)

    matriz = [[f"Respaldo de Registro SHEIN · {ahora():%d/%m/%Y %H:%M}"], []]
    for tab in RESPALDO_TABS:
        df = leer(tab)
        matriz.append([f"### {tab}"])
        matriz.extend(_a_celdas(df))
        matriz.append([])
    ancho = max(len(f) for f in matriz)
    matriz = [f + [""] * (ancho - len(f)) for f in matriz]

    nombre = f"{hoy():%Y-%m-%d}"
    try:
        ws = libro_r.worksheet(nombre)
        ws.clear()
        ws.resize(rows=len(matriz) + 10, cols=max(ancho, 1))
    except gspread.WorksheetNotFound:
        ws = libro_r.add_worksheet(nombre, rows=len(matriz) + 10, cols=max(ancho, 1))
    ws.update(values=matriz, range_name="A1", value_input_option="RAW")

    fechadas = sorted(w.title for w in libro_r.worksheets()
                      if re.fullmatch(r"\d{4}-\d{2}-\d{2}", w.title))
    borradas = 0
    for viejo in fechadas[:-RESPALDO_DIAS]:
        libro_r.del_worksheet(libro_r.worksheet(viejo))
        borradas += 1
    # la hoja no puede quedar sin pestañas; si quedó la "Hoja 1" vacía inicial, se borra
    for w in libro_r.worksheets():
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", w.title) and len(libro_r.worksheets()) > 1 \
                and not any(w.get_values("A1:B2")):
            libro_r.del_worksheet(w)
    set_config(RESPALDO_FECHA=f"{hoy():%d/%m/%Y}")
    total = sum(len(leer(t_)) for t_ in RESPALDO_TABS)
    return f"Respaldo {nombre}: {total} filas" + (f", {borradas} antiguos borrados" if borradas
                                                   else "")


def set_lote(lote, estado, llegada=None):
    """Marca un lote como EN CAMINO o RECIBIDO (con su fecha de llegada)."""
    df = leer("LOTES")
    df = df[pd.to_numeric(df["LOTE"], errors="coerce") != float(lote)]
    fila = {"LOTE": float(lote), "ESTADO": estado,
            "LLEGADA": pd.Timestamp(llegada) if llegada else None}
    df = pd.concat([df, pd.DataFrame([fila])], ignore_index=True).sort_values("LOTE")
    guardar("LOTES", df)


def set_config(**valores):
    """Actualiza claves de CONFIG sin borrar las demás."""
    cfg_n = leer("CONFIG")
    claves = {k.upper() for k in valores}
    cfg_n = cfg_n[~cfg_n["CLAVE"].astype(str).str.upper().isin(claves)]
    nuevas = pd.DataFrame([{"CLAVE": k.upper(), "VALOR": str(v)} for k, v in valores.items()])
    guardar("CONFIG", pd.concat([cfg_n, nuevas], ignore_index=True))


# --------------------------------------------------------------------------
# Logica de negocio
# --------------------------------------------------------------------------
def inventario_calculado(inv, ven):
    inv = inv.copy()
    vendidas = ven.groupby("NOMBRE").size()
    ingresos = ven.groupby("NOMBRE")["MONTO"].sum()
    inv["VENDIDAS"] = inv["NOMBRE"].map(vendidas).fillna(0).astype(int)
    inv["EXISTENCIA"] = inv["CANTIDAD"].fillna(0) - inv["VENDIDAS"]
    inv["INVERSION"] = inv["CANTIDAD"].fillna(0) * inv["PRECIO_COMPRA"].fillna(0)
    inv["COSTO_STOCK"] = inv["EXISTENCIA"].clip(lower=0) * inv["PRECIO_COMPRA"].fillna(0)
    inv["VALOR_STOCK"] = inv["EXISTENCIA"].clip(lower=0) * inv["PRECIO_VENTA"].fillna(0)
    inv["GANANCIA_POTENCIAL"] = inv["VALOR_STOCK"] - inv["COSTO_STOCK"]
    inv["INGRESO_REAL"] = inv["NOMBRE"].map(ingresos).fillna(0)
    inv["GANANCIA_REAL"] = inv["INGRESO_REAL"] - inv["VENDIDAS"] * inv["PRECIO_COMPRA"].fillna(0)
    inv["MARGEN_UNIT"] = inv["PRECIO_VENTA"].fillna(0) - inv["PRECIO_COMPRA"].fillna(0)
    return inv


def ventas_calculadas(ven, inv):
    costos = inv.drop_duplicates("NOMBRE").set_index("NOMBRE")["PRECIO_COMPRA"]
    ven = ven.copy()
    ven["COSTO"] = ven["NOMBRE"].map(costos)
    ven["EN_INVENTARIO"] = ven["COSTO"].notna()
    ven["GANANCIA"] = ven["MONTO"].fillna(0) - ven["COSTO"].fillna(0)
    ven["MES"] = ven["FECHA"].dt.to_period("M").astype(str)
    return ven


def moneda_norm(serie):
    m = serie.fillna("").astype(str).str.strip().str.upper()
    return m.map(lambda x: "Bs" if x in ("BS", "BS.", "BOLIVARES", "BOLÍVARES", "VES")
                 else "$")


def encargos_calculados(enc, tasa=None):
    enc = enc.copy()
    if enc.empty:
        for c in ("GANANCIA", "COSTO_CONOCIDO", "ESTA_PAGADO", "SALDO", "GAN_COBRADA",
                  "COBRAR_USD", "COSTO_USD", "ABONADO_USD", "SALDO_USD", "GANANCIA_USD",
                  "GAN_COBRADA_USD", "SIN_TASA"):
            enc[c] = pd.Series(dtype="object")
        return enc
    enc["TIPO"] = enc["TIPO"].astype(str).str.strip().str.upper().replace(
        {"": "SIN CLASIFICAR", "NAN": "SIN CLASIFICAR", "P": "PRÉSTAMO",
         "PRESTAMO": "PRÉSTAMO", "E": "ENCARGO"})
    enc["ABONADO"] = enc["ABONADO"].fillna(0).clip(lower=0)
    enc["SALDO"] = (enc["COBRAR"].fillna(0) - enc["ABONADO"]).clip(lower=0)
    enc["ESTA_PAGADO"] = enc["SALDO"] <= 0.004
    enc["COSTO_CONOCIDO"] = enc["COSTO"].notna()
    enc["GANANCIA"] = enc["COBRAR"].fillna(0) - enc["COSTO"]
    # la ganancia se realiza al final: primero se repone el costo, el resto es ganancia
    enc["GAN_COBRADA"] = (enc["ABONADO"] - enc["COSTO"].fillna(0)).clip(lower=0)
    enc.loc[~enc["COSTO_CONOCIDO"], "GAN_COBRADA"] = pd.NA
    # moneda: lo que está en Bs se lleva a dólares con la tasa del día
    enc["MONEDA"] = moneda_norm(enc["MONEDA"])
    factor = enc["MONEDA"].map(lambda m: 1.0 if m == "$" else (1 / tasa if tasa else float("nan")))
    for c in ("COBRAR", "COSTO", "ABONADO", "SALDO", "GANANCIA", "GAN_COBRADA"):
        enc[c + "_USD"] = pd.to_numeric(enc[c], errors="coerce") * factor
    enc["SIN_TASA"] = (enc["MONEDA"] == "Bs") & (not tasa)
    return enc


def gastos_calculados(g, tasa=None):
    g = g.copy()
    g["MONEDA"] = moneda_norm(g["MONEDA"]) if len(g) else g["MONEDA"]
    factor = g["MONEDA"].map(lambda m: 1.0 if m == "$" else (1 / tasa if tasa else float("nan")))
    g["MONTO_USD"] = pd.to_numeric(g["MONTO"], errors="coerce").fillna(0) * \
        pd.to_numeric(factor, errors="coerce")
    return g


CATEGORIAS_GASTO = ["ENVÍO", "EMPAQUE", "DELIVERY", "OTRO"]


def caja_calculada(caja, tasa=None):
    caja = caja.copy()
    caja["MONEDA"] = moneda_norm(caja["MONEDA"]) if len(caja) else caja["MONEDA"]
    factores = []
    for _, r in caja.iterrows():
        if r["MONEDA"] != "Bs":
            factores.append(1.0)
        else:
            t = tasa or _num(r["TASA"]) or None
            factores.append(1 / t if t else float("nan"))
    caja["MONTO_USD"] = caja["MONTO"].fillna(0) * pd.Series(factores, index=caja.index,
                                                           dtype="float64")
    return caja


ZONA = ZoneInfo("America/Caracas")


def ahora():
    """Hora de Venezuela (el servidor de Streamlit Cloud corre en UTC)."""
    return dt.datetime.now(ZONA).replace(tzinfo=None)


def hoy():
    return ahora().date()


def _num(v, default=0.0):
    """Número seguro: NaN, None, vacío o texto inválido -> default."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return default
    return default if pd.isna(f) else f


def _ent(v, default=0):
    return int(round(_num(v, default)))


def money(x):
    x = _num(x, None)
    if x is None:
        return "—"
    return f"-${-x:,.2f}" if x < 0 else f"${x:,.2f}"


def leer_monto(texto):
    """Texto del usuario -> float. Acepta 10,32 · 10.32 · 1.234,56 · $5. Vacío = 0.
    Devuelve None si no es un número."""
    t = str(texto or "").strip().replace("$", "").replace(" ", "").replace("Bs", "")
    if t == "":
        return 0.0
    if not all(ch.isdigit() or ch in ".,-" for ch in t):
        return None
    if "," in t and "." in t:
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def fmt_monto(v):
    """Valor inicial de un campo de dinero, con coma decimal."""
    v = _num(v)
    return "" if v == 0 else f"{v:.2f}".replace(".", ",")


def campo_dinero(contenedor, etiqueta, valor=0.0, key=None, ayuda=None):
    """Campo de texto para montos: no depende del teclado decimal del navegador."""
    return contenedor.text_input(etiqueta, value=fmt_monto(valor), placeholder="0,00",
                                 key=key, help=ayuda)


def plata(x, moneda="$"):
    """Monto con su moneda: $1,234.00 o Bs 1,234.00."""
    x = _num(x, None)
    if x is None:
        return "—"
    signo, x = ("-", -x) if x < 0 else ("", x)
    return f"{signo}${x:,.2f}" if moneda != "Bs" else f"{signo}Bs {x:,.2f}"


def avisar(texto):
    """Mensaje que se muestra después del rerun (si no, la recarga lo borra)."""
    st.session_state["aviso"] = texto


MSG_COBRO_DEFECTO = ("Hola {nombre}! 👋 Te escribo por tu cuenta pendiente: te queda un "
                     "saldo de {saldo}. Hasta ahora has abonado {abonado}. ¿Cuándo podrías "
                     "hacer un abono? ¡Gracias!")


def wa_numero(tel):
    """0414-123.45.67 / 4141234567 / +58 414 1234567 -> 584141234567. Vacío si no sirve."""
    d = "".join(ch for ch in str(tel or "") if ch.isdigit())
    if d.startswith("58") and len(d) >= 12:
        return d
    if d.startswith("0") and len(d) == 11:
        return "58" + d[1:]
    if len(d) == 10 and d.startswith("4"):
        return "58" + d
    return d if len(d) >= 11 else ""


def wa_link(tel, mensaje):
    from urllib.parse import quote
    num = wa_numero(tel)
    return f"https://wa.me/{num}?text={quote(mensaje)}" if num else \
        f"https://wa.me/?text={quote(mensaje)}"


def montos_por_moneda(serie_monto, serie_moneda):
    """'$1,132.00 y Bs 3,800.00' a partir de montos en monedas mezcladas."""
    df = pd.DataFrame({"M": pd.to_numeric(serie_monto, errors="coerce").fillna(0),
                       "MON": serie_moneda.values})
    partes = [plata(v, m) for m, v in df.groupby("MON")["M"].sum().items() if abs(v) > 0.004]
    return " y ".join(partes) if partes else plata(0)


NO_MOVER = "— No mover la caja —"


def convertir(monto, de, a, tasa):
    if de == a:
        return monto
    if not tasa:
        return None
    return monto * tasa if (de == "$" and a == "Bs") else monto / tasa


def selector_caja(cont, etiqueta, key, mover_por_defecto=True):
    """Ubicación de caja donde entra o sale la plata. Recuerda la última usada."""
    ubis = sorted(x for x in CAJA["UBICACION"].unique() if x)
    ops = [NO_MOVER] + ubis
    ult = st.session_state.get("ult_caja")
    if not mover_por_defecto:
        idx = 0
    elif ult in ops:
        idx = ops.index(ult)
    elif "EFECTIVO" in ops:
        idx = ops.index("EFECTIVO")
    else:
        idx = 1 if ubis else 0
    v = cont.selectbox(etiqueta, ops, index=idx, key=key)
    return None if v == NO_MOVER else v


def plan_caja(ubic, monto, moneda="$"):
    """Valida un movimiento antes de guardar nada.
    Devuelve (delta en la moneda de la ubicación, moneda de la ubicación, error)."""
    if not ubic or abs(monto) < 0.0001:
        return 0.0, None, None
    fila = CAJA[CAJA["UBICACION"] == ubic]
    mon_u = fila["MONEDA"].iloc[0] if len(fila) else "$"
    delta = convertir(monto, moneda, mon_u, TASA)
    if delta is None:
        return None, mon_u, (f"{ubic} está en bolívares y no hay tasa del día. "
                             "Cárgala en la barra lateral.")
    actual = _num(fila["MONTO"].sum())
    if actual + delta < -0.004:
        return None, mon_u, (f"No alcanza: {ubic} tiene {plata(actual, mon_u)} y este "
                             f"movimiento saca {plata(-delta, mon_u)}.")
    return delta, mon_u, None


def aplicar_caja(ubic, delta, origen="", detalle=""):
    """Mueve el saldo de una ubicación y lo deja en la bitácora.
    Devuelve texto para el mensaje de éxito."""
    if not ubic or abs(delta) < 0.0001:
        return ""
    df = leer("CAJA")
    df["MONTO"] = pd.to_numeric(df["MONTO"], errors="coerce").fillna(0).astype(float)
    mascara = df["UBICACION"] == ubic
    mon_u = moneda_norm(df.loc[mascara, "MONEDA"]).iloc[0] if mascara.any() else "$"
    df.loc[mascara, "MONTO"] = df.loc[mascara, "MONTO"] + delta
    guardar("CAJA", df)
    registrar_mov(ubic, delta, mon_u, _num(df.loc[mascara, "MONTO"].sum()), origen, detalle)
    st.session_state["ult_caja"] = ubic
    verbo = "Entraron" if delta > 0 else "Salieron"
    return f" {verbo} {plata(abs(delta), mon_u)} {'a' if delta > 0 else 'de'} {ubic}."


def registrar_mov(ubic, delta, moneda, saldo, origen, detalle=""):
    try:
        anexar("MOVIMIENTOS", [{"FECHA": pd.Timestamp(ahora()).floor("min"),
                                "UBICACION": ubic, "MONTO": delta, "MONEDA": moneda,
                                "SALDO": saldo, "ORIGEN": origen or "Otro",
                                "DETALLE": detalle}])
    except Exception as e:      # la bitácora nunca debe tumbar el guardado principal
        st.session_state["aviso_bitacora"] = f"No se pudo anotar en la bitácora: {e}"


# --------------------------------------------------------------------------
# Carga
# --------------------------------------------------------------------------
inv_raw = leer("INVENTARIO")
ven_raw = leer("VENTAS")
enc = leer("ENCARGOS")
caja = leer("CAJA")
cfg = leer("CONFIG")
_cfg = dict(zip(cfg["CLAVE"].astype(str).str.upper(), cfg["VALOR"]))
TASA = leer_monto(_cfg.get("TASA")) or None
TASA_FECHA = str(_cfg.get("TASA_FECHA") or "")
TASA_FUENTE = str(_cfg.get("TASA_FUENTE") or "")

# tasa automática: una vez al día, salvo que ya la cargaste a mano hoy
if (secreto("tasa_automatica", True) and TASA_FECHA != f"{hoy():%d/%m/%Y}"
        and not st.session_state.get("tasa_intentada")):
    st.session_state["tasa_intentada"] = True
    _t, _vig, _fuente = tasa_bcv()
    if _t:
        try:
            set_config(TASA=_t, TASA_FECHA=f"{hoy():%d/%m/%Y}",
                       TASA_FUENTE=_fuente + (f" (vigente {_vig[8:10]}/{_vig[5:7]})"
                                              if len(_vig) == 10 else ""))
            TASA, TASA_FECHA = _t, f"{hoy():%d/%m/%Y}"
            TASA_FUENTE = str(leer("CONFIG").set_index("CLAVE")["VALOR"].get("TASA_FUENTE", ""))
        except Exception as e:
            st.session_state["tasa_error"] = f"No se pudo guardar la tasa: {e}"
    else:
        st.session_state["tasa_error"] = f"No se pudo traer la tasa BCV ({_fuente})."

# respaldo diario: al primer uso del día, antes de cualquier cambio
RESPALDO_FECHA = str(_cfg.get("RESPALDO_FECHA") or "")
if (BACKEND == "gsheets" and secreto("backup_spreadsheet_id")
        and RESPALDO_FECHA != f"{hoy():%d/%m/%Y}"
        and not st.session_state.get("respaldo_intentado")):
    st.session_state["respaldo_intentado"] = True
    try:
        st.session_state["respaldo_ok"] = hacer_respaldo()
        RESPALDO_FECHA = f"{hoy():%d/%m/%Y}"
    except Exception as e:
        st.session_state["respaldo_error"] = f"{type(e).__name__}: {e}"

LOTES = leer("LOTES")
EN_CAMINO = set(pd.to_numeric(
    LOTES.loc[LOTES["ESTADO"].astype(str).str.upper() == "EN CAMINO", "LOTE"],
    errors="coerce").dropna())
LLEGADAS = dict(zip(pd.to_numeric(LOTES["LOTE"], errors="coerce"), LOTES["LLEGADA"]))
INV = inventario_calculado(inv_raw, ven_raw)
INV["EN_CAMINO"] = INV["LOTE"].isin(EN_CAMINO)
VEN = ventas_calculadas(ven_raw, inv_raw)
ENC = encargos_calculados(enc, TASA)
CAJA = caja_calculada(caja, TASA)
gas = leer("GASTOS")
GAS = gastos_calculados(gas, TASA)
contactos = leer("CONTACTOS")
TELEFONOS = dict(zip(contactos["PERSONA"].astype(str).str.upper(), contactos["TELEFONO"]))
MSG_COBRO = str(_cfg.get("MSG_COBRO") or "") or MSG_COBRO_DEFECTO
PEND = ENC[~ENC["ESTA_PAGADO"]]

# --------------------------------------------------------------------------
# Sidebar
# --------------------------------------------------------------------------
with st.sidebar:
    st.title("💍 Registro SHEIN")
    seccion = st.radio(
        "Secciones",
        ["📊 Dashboard", "🧾 Ventas", "📦 Inventario", "🎁 Por cobrar", "🚚 Gastos",
         "💵 Caja Chica", "📈 Rotación", "🏛️ Histórico"],
        label_visibility="collapsed",
    )
    st.divider()
    st.caption(f"Backend: **{BACKEND}**")
    if FALLA_SECRETS:
        st.error(f"secrets.toml no se pudo leer ({FALLA_SECRETS}). "
                 "Revísalo: si lo generaste en PowerShell, guárdalo en UTF-8 sin BOM.", icon="🔑")
    if BACKEND != "gsheets":
        st.warning("Modo Excel local: en Streamlit Cloud los cambios NO se guardan.", icon="⚠️")
    st.caption("Tasa del día (Bs por $)" + (f" · {TASA_FECHA}" if TASA_FECHA else "")
               + (f"  \n{TASA_FUENTE}" if TASA_FUENTE else ""))
    if st.session_state.get("tasa_error"):
        st.caption("⚠️ " + st.session_state["tasa_error"] + " Cárgala a mano.")
    with st.form("form_tasa", clear_on_submit=False, border=False):
        tasa_txt = st.text_input("Tasa", value=fmt_monto(TASA), placeholder="Ej. 312,45",
                                 label_visibility="collapsed")
        if st.form_submit_button("Guardar tasa", width="stretch"):
            nueva = leer_monto(tasa_txt)
            if not nueva or nueva <= 0:
                st.error("Tasa inválida.")
            else:
                set_config(TASA=nueva, TASA_FECHA=f"{hoy():%d/%m/%Y}", TASA_FUENTE="manual")
                st.rerun()
    if st.button("🔄 Traer tasa BCV", width="stretch"):
        tasa_bcv.clear()
        _t, _vig, _fuente = tasa_bcv()
        if _t:
            set_config(TASA=_t, TASA_FECHA=f"{hoy():%d/%m/%Y}",
                       TASA_FUENTE=_fuente + (f" (vigente {_vig[8:10]}/{_vig[5:7]})"
                                              if len(_vig) == 10 else ""))
            st.session_state.pop("tasa_error", None)
            avisar(f"Tasa BCV actualizada: {_t:,.2f} Bs/$.")
            st.rerun()
        else:
            st.error(f"No se pudo traer ({_fuente}). Cárgala a mano.")
    if not TASA and (CAJA["MONEDA"].eq("Bs").any() or ENC["MONEDA"].eq("Bs").any()
                     or GAS["MONEDA"].eq("Bs").any()):
        st.warning("Hay montos en bolívares y no hay tasa: no se suman a los totales.",
                   icon="💱")
    st.divider()
    if BACKEND == "gsheets":
        if not secreto("backup_spreadsheet_id"):
            st.caption("💾 Respaldo diario sin configurar (falta backup_spreadsheet_id).")
        else:
            st.caption(f"💾 Último respaldo: {RESPALDO_FECHA or 'nunca'}")
            if st.session_state.get("respaldo_error"):
                st.caption("⚠️ El respaldo falló: " + st.session_state["respaldo_error"][:160])
            if st.button("💾 Respaldar ahora", width="stretch"):
                try:
                    avisar(hacer_respaldo())
                    st.session_state.pop("respaldo_error", None)
                except Exception as e:
                    st.session_state["respaldo_error"] = f"{type(e).__name__}: {e}"
                st.rerun()
    if st.button("🔄 Refrescar datos", width="stretch"):
        st.cache_data.clear()
        st.rerun()

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as w:
        INV.to_excel(w, sheet_name="INVENTARIO", index=False)
        VEN.to_excel(w, sheet_name="VENTAS", index=False)
        ENC.to_excel(w, sheet_name="ENCARGOS", index=False)
        GAS.to_excel(w, sheet_name="GASTOS", index=False)
        caja.to_excel(w, sheet_name="CAJA", index=False)
    st.download_button(
        "⬇️ Descargar Excel",
        buffer.getvalue(),
        file_name=f"REGISTRO_SHEIN_{hoy():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )

if "aviso" in st.session_state:
    st.success(st.session_state.pop("aviso"), icon="✅")
if "aviso_bitacora" in st.session_state:
    st.warning(st.session_state.pop("aviso_bitacora"), icon="📒")

# --------------------------------------------------------------------------
# DASHBOARD
# --------------------------------------------------------------------------
if seccion == "📊 Dashboard":
    st.header("Resumen general")

    if VEN.empty:
        st.info("Todavía no hay ventas registradas.")
        st.stop()

    incluir = st.toggle("Incluir el histórico 2025 (lotes 1-10)", value=False,
                        help="Suma las ventas del archivo viejo. Ese período no tiene costo "
                             "por venta, así que el margen se sigue calculando solo con la "
                             "etapa actual.")

    base = VEN.assign(ETAPA="Actual")
    if incluir:
        vh = leer("VENTAS_HIST")
        if vh.empty:
            st.warning("No hay histórico cargado todavía. Corre `importar_v1.py` primero.")
        else:
            vh = vh.assign(ETAPA="Histórico", COSTO=pd.NA, GANANCIA=pd.NA,
                           EN_INVENTARIO=True)
            vh["MES"] = vh["FECHA"].dt.to_period("M").astype(str)
            base = pd.concat([base, vh[base.columns]], ignore_index=True)

    lotes = sorted(set(INV["LOTE"].dropna()) | set(base["LOTE"].dropna()))
    c1, c2 = st.columns([2, 3])
    with c1:
        fechas = pd.concat([base["FECHA"], GAS["FECHA"]]) if len(GAS) else base["FECHA"]
        f_min, f_max = fechas.min(), fechas.max()
        hoy_d = hoy()
        if pd.isna(f_max) or f_max.date() < hoy_d:
            f_max = pd.Timestamp(hoy_d)   # el rango llega a hoy: no deja fuera gastos recientes
        rango = st.date_input(
            "Rango de fechas",
            value=(f_min.date() if pd.notna(f_min) else hoy_d,
                   f_max.date() if pd.notna(f_max) else hoy_d),
            format="DD/MM/YYYY",
            key=f"rango_{incluir}",
        )
    with c2:
        f_lotes = st.multiselect("Lotes", [int(l) for l in lotes], default=[])

    v = base.copy()
    if isinstance(rango, (tuple, list)) and len(rango) == 2:
        v = v[(v["FECHA"].dt.date >= rango[0]) & (v["FECHA"].dt.date <= rango[1])]
    if f_lotes:
        v = v[v["LOTE"].isin(f_lotes)]

    actual = v[v["ETAPA"] == "Actual"]
    ingresos = v["MONTO"].sum()
    costo = pd.to_numeric(actual["COSTO"], errors="coerce").fillna(0).sum()
    ganancia = actual["MONTO"].sum() - costo
    margen = (ganancia / actual["MONTO"].sum() * 100) if len(actual) else 0
    solo_actual = " · etapa actual" if incluir else ""

    g = GAS.copy()
    if isinstance(rango, (tuple, list)) and len(rango) == 2 and len(g):
        g = g[(g["FECHA"].dt.date >= rango[0]) & (g["FECHA"].dt.date <= rango[1])]
    if f_lotes and len(g):
        g = g[g["LOTE"].isin(f_lotes)]
    gastos_p = g["MONTO_USD"].sum() if len(g) else 0.0
    despues = ganancia - gastos_p
    margen_d = (despues / actual["MONTO"].sum() * 100) if len(actual) else 0

    k = st.columns(4)
    k[0].metric("Ingresos", money(ingresos), f"{len(v)} ventas")
    k[1].metric("Costo de lo vendido" + solo_actual, money(costo))
    k[2].metric("Ganancia bruta" + solo_actual, money(ganancia), f"{margen:.1f}% margen")
    k[3].metric("Ticket promedio", money(v["MONTO"].mean() if len(v) else 0))
    k = st.columns(4)
    k[0].metric("Gastos (envíos y otros)", money(gastos_p), f"{len(g)} registros")
    k[1].metric("Ganancia después de gastos", money(despues), f"{margen_d:.1f}% margen")

    k = st.columns(4)
    k[0].metric("Unidades en stock", int(INV["EXISTENCIA"].clip(lower=0).sum()))
    k[1].metric("Stock a costo", money(INV["COSTO_STOCK"].sum()))
    k[2].metric("Ganancia potencial", money(INV["GANANCIA_POTENCIAL"].sum()))
    k[3].metric("Inversión total (histórica)", money(INV["INVERSION"].sum()))

    sin_costo = actual[~actual["EN_INVENTARIO"]]
    if len(sin_costo):
        st.caption(
            f"⚠️ {len(sin_costo)} ventas ({money(sin_costo['MONTO'].sum())}) no tienen "
            "producto en INVENTARIO, se cuentan con costo 0."
        )

    st.divider()
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("Ingresos y ganancia por mes")
        if incluir:
            mensual = (v.groupby(["MES", "ETAPA"])["MONTO"].sum().reset_index()
                       .sort_values("MES"))
            fig = px.bar(mensual, x="MES", y="MONTO", color="ETAPA",
                         labels={"MONTO": "$", "MES": "", "ETAPA": ""})
        else:
            mensual = (
                v.groupby("MES")
                .agg(INGRESOS=("MONTO", "sum"), GANANCIA=("GANANCIA", "sum"))
                .reset_index()
                .sort_values("MES")
            )
            fig = px.bar(
                mensual, x="MES", y=["INGRESOS", "GANANCIA"], barmode="group",
                labels={"value": "$", "MES": "", "variable": ""},
            )
        fig.update_layout(height=340, legend_title="")
        st.plotly_chart(fig, width="stretch")

    with c2:
        st.subheader("Acumulado en el tiempo")
        diario = v.sort_values("FECHA").copy()
        diario["ACUM_INGRESOS"] = diario["MONTO"].cumsum()
        diario["ACUM_GANANCIA"] = pd.to_numeric(
            diario["GANANCIA"], errors="coerce").fillna(0).cumsum()
        fig = px.line(
            diario, x="FECHA", y=["ACUM_INGRESOS", "ACUM_GANANCIA"],
            labels={"value": "$", "FECHA": "", "variable": ""},
        )
        fig.update_layout(height=340, legend_title="")
        st.plotly_chart(fig, width="stretch")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Top 10 productos por ingreso")
        top = (
            v.groupby("NOMBRE")
            .agg(INGRESOS=("MONTO", "sum"), UNIDADES=("MONTO", "size"))
            .reset_index()
            .nlargest(10, "INGRESOS")
            .sort_values("INGRESOS")
        )
        fig = px.bar(top, x="INGRESOS", y="NOMBRE", orientation="h", text="UNIDADES")
        fig.update_layout(height=380, yaxis_title="", xaxis_title="$")
        st.plotly_chart(fig, width="stretch")

    with c2:
        st.subheader("Rentabilidad por lote")
        inv_lote = INV.groupby("LOTE").agg(INVERTIDO=("INVERSION", "sum")).reset_index()
        ven_lote = VEN.groupby("LOTE").agg(RECUPERADO=("MONTO", "sum")).reset_index()
        gas_lote = (GAS[GAS["LOTE"].notna()].groupby("LOTE")
                    .agg(GASTOS=("MONTO_USD", "sum")).reset_index())
        lote = (inv_lote.merge(ven_lote, on="LOTE", how="outer")
                .merge(gas_lote, on="LOTE", how="outer").fillna(0))
        lote = lote[lote["LOTE"].notna()].sort_values("LOTE")
        lote["LOTE"] = lote["LOTE"].astype(int).astype(str)
        fig = px.bar(
            lote, x="LOTE", y=["INVERTIDO", "GASTOS", "RECUPERADO"], barmode="group",
            labels={"value": "$", "variable": ""},
        )
        fig.update_layout(height=380, legend_title="")
        st.plotly_chart(fig, width="stretch")

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Sin existencia (reponer)")
        agotados = INV[INV["EXISTENCIA"] <= 0][["NOMBRE", "VENDIDAS", "INGRESO_REAL", "LOTE"]]
        st.dataframe(
            agotados.sort_values("INGRESO_REAL", ascending=False),
            width="stretch", hide_index=True, height=280,
        )
    with c2:
        st.subheader("Stock disponible")
        stock = INV[INV["EXISTENCIA"] > 0][
            ["NOMBRE", "EXISTENCIA", "PRECIO_VENTA", "VALOR_STOCK", "LOTE"]
        ]
        st.dataframe(
            stock.sort_values("VALOR_STOCK", ascending=False),
            width="stretch", hide_index=True, height=280,
        )

    st.subheader("Proyección: cobrar todo y liquidar el stock")

    caja_hoy = CAJA["MONTO_USD"].sum()
    por_cobrar = PEND["SALDO_USD"].sum() if len(PEND) else 0.0
    con_costo = PEND[PEND["COSTO_CONOCIDO"]] if len(PEND) else PEND
    gan_encargos = (con_costo["GANANCIA_USD"].sum()
                    - con_costo["GAN_COBRADA_USD"].fillna(0).sum()) if len(con_costo) else 0.0
    recuperacion = por_cobrar - gan_encargos
    sin_costo_enc = int((~PEND["COSTO_CONOCIDO"]).sum()) if len(PEND) else 0
    monto_sin_costo = PEND[~PEND["COSTO_CONOCIDO"]]["SALDO_USD"].sum() if len(PEND) else 0.0

    stock_costo = INV["COSTO_STOCK"].sum()
    stock_venta = INV["VALOR_STOCK"].sum()
    potencial = INV["GANANCIA_POTENCIAL"].sum()
    realizada = VEN["GANANCIA"].sum()

    tras_cobrar = caja_hoy + por_cobrar
    final = tras_cobrar + stock_venta

    k = st.columns(3)
    k[0].metric("Caja hoy", money(caja_hoy))
    k[1].metric("Saldo por cobrar", money(por_cobrar), f"{len(PEND)} cuentas abiertas")
    k[2].metric("Efectivo si cobran todo", money(tras_cobrar))

    k = st.columns(3)
    k[0].metric("Stock a precio de venta", money(stock_venta),
                f"{int(INV['EXISTENCIA'].clip(lower=0).sum())} artículos")
    k[1].metric("Efectivo si además se vende todo", money(final))
    k[2].metric("Ganancia por realizar", money(potencial + gan_encargos),
                f"{money(potencial)} stock + {money(gan_encargos)} por cobrar")

    fig = go.Figure(go.Waterfall(
        orientation="v",
        measure=["absolute", "relative", "relative", "total"],
        x=["Caja hoy", "Cobro de lo pendiente", "Venta del stock", "Efectivo final"],
        y=[caja_hoy, por_cobrar, stock_venta, 0],
        text=[money(caja_hoy), money(por_cobrar), money(stock_venta), money(final)],
        textposition="outside",
        connector={"line": {"dash": "dot"}},
    ))
    fig.update_layout(height=340, showlegend=False, yaxis_title="$")
    st.plotly_chart(fig, width="stretch")

    nota = (f"De los {money(por_cobrar)} por cobrar, {money(recuperacion)} es reponer lo que "
            f"gastaste comprando el encargo y {money(gan_encargos)} es ganancia tuya. "
            f"De la venta del stock, {money(stock_costo)} recupera el costo y "
            f"{money(potencial)} es ganancia.")
    sin_tasa_n = int(CAJA["MONTO_USD"].isna().sum()) + int(PEND["SALDO_USD"].isna().sum()
                                                           if len(PEND) else 0)
    if sin_tasa_n:
        nota += (f" {sin_tasa_n} montos en bolívares quedaron fuera porque no hay tasa del "
                 "día cargada.")
    elif TASA:
        nota += f" Bolívares convertidos a {TASA:,.2f} Bs/$."
    if sin_costo_enc:
        nota += (f" Ojo: {sin_costo_enc} encargos ({money(monto_sin_costo)}) no tienen el costo "
                 "anotado, así que esa parte no se puede repartir.")
    st.caption(nota)

    st.divider()
    st.subheader("Resultado acumulado de la etapa")
    gan_enc_cobrada = ENC["GAN_COBRADA_USD"].fillna(0).sum() if len(ENC) else 0.0
    gastos_tot = GAS["MONTO_USD"].sum() if len(GAS) else 0.0
    k = st.columns(4)
    k[0].metric("Ganancia de reventa", money(realizada))
    k[1].metric("Ganancia de encargos cobrada", money(gan_enc_cobrada))
    k[2].metric("Gastos", money(-gastos_tot) if gastos_tot else money(0))
    k[3].metric("Total ganado", money(realizada + gan_enc_cobrada - gastos_tot))

# --------------------------------------------------------------------------
# VENTAS
# --------------------------------------------------------------------------
elif seccion == "🧾 Ventas":
    st.header("Ventas")

    with st.expander("➕ Registrar venta", expanded=True):
        disponibles = INV[INV["EXISTENCIA"] > 0].sort_values("NOMBRE")
        libre = st.checkbox("Producto fuera del inventario")

        if libre:
            with st.form("venta_libre", clear_on_submit=True):
                c = st.columns([3, 1, 1, 1])
                nombre = c[0].text_input("Producto").strip().upper()
                monto_txt = campo_dinero(c[1], "Monto $", 5.0)
                fecha = c[2].date_input("Fecha", value=hoy(), format="DD/MM/YYYY")
                lote = c[3].number_input("Lote", min_value=0, step=1, value=0)
                ubic = selector_caja(st, "¿Dónde entró la plata?", "caja_venta_libre")
                if st.form_submit_button("Guardar venta", type="primary"):
                    monto = leer_monto(monto_txt)
                    delta, _, err = plan_caja(ubic, monto or 0)
                    if not nombre:
                        st.error("Falta el nombre del producto.")
                    elif monto is None or monto <= 0:
                        st.error(f"Monto inválido: «{monto_txt}». Escribe por ejemplo 5 o 5,50.")
                    elif err:
                        st.error(err)
                    else:
                        agregar("VENTAS", {"NOMBRE": nombre, "MONTO": monto,
                                           "FECHA": pd.Timestamp(fecha), "LOTE": lote or None})
                        extra = aplicar_caja(ubic, delta, "Venta", nombre)
                        avisar(f"Venta registrada: {nombre} · {money(monto)}.{extra}")
                        st.rerun()
        elif disponibles.empty:
            st.warning("No hay productos con existencia. Carga inventario o usa producto libre.")
        else:
            etiquetas = {
                f"{r.NOMBRE}  ·  stock {_ent(r.EXISTENCIA)}  ·  {money(r.PRECIO_VENTA)}"
                + ("  ·  🚚 en camino" if r.EN_CAMINO else ""): r.NOMBRE
                for r in disponibles.itertuples()
            }
            sel = st.selectbox("Producto", list(etiquetas.keys()))
            nombre = etiquetas[sel]
            prod = disponibles[disponibles["NOMBRE"] == nombre].iloc[0]
            with st.form("venta", clear_on_submit=False):
                c = st.columns([1, 1, 1, 1])
                monto_txt = campo_dinero(c[0], "Monto $", prod["PRECIO_VENTA"],
                                         key=f"monto_venta_{nombre}")
                fecha = c[1].date_input("Fecha", value=hoy(), format="DD/MM/YYYY")
                lote = c[2].number_input("Lote", min_value=0, step=1,
                                         value=_ent(prod["LOTE"]), key=f"lote_venta_{nombre}")
                cantidad = c[3].number_input("Unidades", min_value=1, step=1, value=1)
                ubic = selector_caja(st, "¿Dónde entró la plata?", "caja_venta")
                enviado = st.form_submit_button("Guardar venta", type="primary")
                monto = leer_monto(monto_txt)
                delta, _, err = plan_caja(ubic, (monto or 0) * int(cantidad))
                if enviado and (monto is None or monto <= 0):
                    st.error(f"Monto inválido: «{monto_txt}». Escribe por ejemplo 5 o 5,50.")
                elif enviado and int(cantidad) > _ent(prod["EXISTENCIA"]):
                    st.error(f"Solo quedan {_ent(prod['EXISTENCIA'])} de {nombre}.")
                elif enviado and err:
                    st.error(err)
                elif enviado:
                    anexar("VENTAS", [{
                        "NOMBRE": nombre, "MONTO": monto,
                        "FECHA": pd.Timestamp(fecha), "LOTE": lote or None,
                    }] * int(cantidad))
                    extra = aplicar_caja(ubic, delta, "Venta", f"{nombre} x{int(cantidad)}")
                    avisar(f"{int(cantidad)} venta(s) de {nombre} · "
                           f"{money(monto * cantidad)}.{extra}")
                    st.rerun()

    st.divider()
    c1, c2, c3 = st.columns([2, 1, 1])
    busq = c1.text_input("Buscar producto").strip().upper()
    lotes_v = sorted(int(x) for x in VEN["LOTE"].dropna().unique())
    f_lote = c2.multiselect("Lote", lotes_v)
    orden = c3.selectbox("Orden", ["Más recientes", "Más antiguas", "Mayor monto"])

    v = VEN.copy()
    if busq:
        v = v[v["NOMBRE"].str.contains(busq, na=False)]
    if f_lote:
        v = v[v["LOTE"].isin(f_lote)]
    v = v.sort_values(
        {"Más recientes": "FECHA", "Más antiguas": "FECHA", "Mayor monto": "MONTO"}[orden],
        ascending=(orden == "Más antiguas"),
    )

    k = st.columns(3)
    k[0].metric("Ventas", len(v))
    k[1].metric("Ingresos", money(v["MONTO"].sum()))
    k[2].metric("Ganancia", money(v["GANANCIA"].sum()))

    st.dataframe(
        v[["FECHA", "NOMBRE", "MONTO", "COSTO", "GANANCIA", "LOTE"]],
        width="stretch", hide_index=True, height=420,
        column_config={
            "FECHA": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
            "MONTO": st.column_config.NumberColumn("Monto", format="$%.2f"),
            "COSTO": st.column_config.NumberColumn("Costo", format="$%.2f"),
            "GANANCIA": st.column_config.NumberColumn("Ganancia", format="$%.2f"),
            "LOTE": st.column_config.NumberColumn("Lote", format="%d"),
        },
    )

    with st.expander("🗑️ Eliminar una venta"):
        st.caption("Para borrar una venta que registraste dos veces o por error. Las más "
                   "recientes salen primero; escribe en la lista para buscar.")
        lista = ven_raw.sort_values("FECHA", ascending=False, na_position="last")
        ops = {}
        for n, r in enumerate(lista.itertuples(), 1):
            fecha = f"{r.FECHA:%d/%m/%Y}" if pd.notna(r.FECHA) else "sin fecha"
            lote = f" · lote {_ent(r.LOTE)}" if pd.notna(r.LOTE) else ""
            ops[f"#{n} · {fecha} · {r.NOMBRE} · {money(r.MONTO)}{lote}"] = r.Index
        if ops:
            elegida = st.selectbox("Venta a eliminar", list(ops.keys()), key="borrar_ven")
            ubic = selector_caja(st, "Devolver la plata desde",
                                 "caja_borrar_ven", mover_por_defecto=False)
            st.caption("Si al registrarla movió la caja, elige la misma ubicación para "
                       "revertirlo. Si fue un duplicado que nunca entró, deja «No mover».")
            conf = st.checkbox("Sí, quiero borrar esta venta", key="conf_borrar_ven")
            if st.button("🗑️ Eliminar venta", disabled=not conf, key="btn_borrar_ven"):
                monto_v = _num(ven_raw.loc[ops[elegida], "MONTO"])
                delta, _, err = plan_caja(ubic, -monto_v)
                if err:
                    st.error(err)
                else:
                    df = leer("VENTAS")
                    guardar("VENTAS", df.drop(index=ops[elegida]))
                    extra = aplicar_caja(ubic, delta, "Reversión", f"venta eliminada: {elegida}")
                    avisar("Venta eliminada. La existencia del producto se recalcula sola."
                           + extra)
                    st.rerun()

    with st.expander("✏️ Editar / eliminar histórico de ventas"):
        st.caption("Edita celdas, agrega filas o marca filas y presiona Supr para borrar.")
        edit = st.data_editor(
            ven_raw, num_rows="dynamic", width="stretch", height=400,
            key="ed_ventas",
            column_config={
                "FECHA": st.column_config.DateColumn("FECHA", format="DD/MM/YYYY"),
                "MONTO": st.column_config.NumberColumn("MONTO", format="%.2f"),
                "LOTE": st.column_config.NumberColumn("LOTE", format="%d"),
            },
        )
        if st.button("💾 Guardar cambios de ventas", type="primary"):
            guardar("VENTAS", edit)
            avisar("Ventas actualizadas.")
            st.rerun()

# --------------------------------------------------------------------------
# INVENTARIO
# --------------------------------------------------------------------------
elif seccion == "📦 Inventario":
    st.header("Inventario")

    with st.expander("➕ Agregar producto"):
        with st.form("nuevo_prod", clear_on_submit=True):
            c = st.columns([3, 2])
            desc = c[0].text_input("Descripción SHEIN")
            nombre = c[1].text_input("Nombre corto (clave única)").strip().upper()
            c = st.columns(4)
            cant = c[0].number_input("Cantidad", min_value=1, step=1, value=1)
            pc_txt = campo_dinero(c[1], "Precio compra $", 1.0)
            pv_txt = campo_dinero(c[2], "Precio venta $", 5.0)
            lote = c[3].number_input("Lote", min_value=0, step=1,
                                     value=_ent(INV["LOTE"].max() if len(INV) else 0))
            foto = st.text_input("Foto (enlace, opcional)", placeholder="https://…",
                                 help="Sale en el catálogo para clientes.")
            no_llega = st.checkbox("Este lote todavía no ha llegado",
                                   help="Sus productos no salen en el catálogo hasta que "
                                        "marques el lote como recibido.")
            ubic = selector_caja(st, "¿De dónde salió el pago?", "caja_nuevo_prod",
                                 mover_por_defecto=False)
            st.caption("Elige la ubicación si estás registrando una compra nueva. "
                       "Deja «No mover» si solo pasas al sistema algo ya pagado.")
            if st.form_submit_button("Agregar", type="primary"):
                pc, pv = leer_monto(pc_txt), leer_monto(pv_txt)
                delta, _, err = plan_caja(ubic, -(pc or 0) * int(cant))
                if pc is None or pv is None:
                    st.error("Revisa los precios: escribe por ejemplo 1,77 o 5.")
                elif err:
                    st.error(err)
                elif not nombre:
                    st.error("El nombre corto es obligatorio.")
                elif nombre in set(inv_raw["NOMBRE"]):
                    st.error("Ese nombre ya existe: usaría el mismo contador de ventas.")
                else:
                    agregar("INVENTARIO", {
                        "DESCRIPCION": desc, "NOMBRE": nombre, "CANTIDAD": cant,
                        "PRECIO_COMPRA": pc, "PRECIO_VENTA": pv, "LOTE": lote,
                        "FOTO": foto.strip(),
                    })
                    if no_llega and lote and lote not in EN_CAMINO:
                        set_lote(lote, "EN CAMINO")
                    extra = aplicar_caja(ubic, delta, "Compra", f"{nombre} x{int(cant)}")
                    avisar(f"{nombre} agregado.{extra}")
                    st.rerun()

    c1, c2, c3 = st.columns([2, 1, 1])
    busq = c1.text_input("Buscar").strip().upper()
    f_lote = c2.multiselect("Lote", sorted(int(x) for x in INV["LOTE"].dropna().unique()))
    estado = c3.selectbox("Estado", ["Todos", "Con existencia", "Agotados", "En camino"])

    i = INV.copy()
    if busq:
        i = i[i["NOMBRE"].str.contains(busq, na=False) | i["DESCRIPCION"].str.upper().str.contains(busq, na=False)]
    if f_lote:
        i = i[i["LOTE"].isin(f_lote)]
    if estado == "Con existencia":
        i = i[i["EXISTENCIA"] > 0]
    elif estado == "Agotados":
        i = i[i["EXISTENCIA"] <= 0]
    elif estado == "En camino":
        i = i[i["EN_CAMINO"]]

    k = st.columns(4)
    k[0].metric("Productos", len(i))
    k[1].metric("Unidades en stock", int(i["EXISTENCIA"].clip(lower=0).sum()))
    k[2].metric("Stock a costo", money(i["COSTO_STOCK"].sum()))
    k[3].metric("Ganancia potencial", money(i["GANANCIA_POTENCIAL"].sum()))

    st.dataframe(
        i[["NOMBRE", "CANTIDAD", "VENDIDAS", "EXISTENCIA", "PRECIO_COMPRA", "PRECIO_VENTA",
           "MARGEN_UNIT", "INGRESO_REAL", "GANANCIA_REAL", "GANANCIA_POTENCIAL", "LOTE",
           "DESCRIPCION"]],
        width="stretch", hide_index=True, height=460,
        column_config={
            "PRECIO_COMPRA": st.column_config.NumberColumn(format="$%.2f"),
            "PRECIO_VENTA": st.column_config.NumberColumn(format="$%.2f"),
            "MARGEN_UNIT": st.column_config.NumberColumn("MARGEN/U", format="$%.2f"),
            "INGRESO_REAL": st.column_config.NumberColumn(format="$%.2f"),
            "GANANCIA_REAL": st.column_config.NumberColumn(format="$%.2f"),
            "GANANCIA_POTENCIAL": st.column_config.NumberColumn("POTENCIAL", format="$%.2f"),
            "LOTE": st.column_config.NumberColumn(format="%d"),
        },
    )

    with st.expander(f"🚚 Lotes en camino ({len(EN_CAMINO)})", expanded=bool(EN_CAMINO)):
        st.caption("Los productos de un lote en camino no salen en el catálogo. Cuando llegue, "
                   "márcalo como recibido y aparecen solos.")
        if EN_CAMINO:
            filas_ec = []
            for l_ in sorted(EN_CAMINO):
                prod_l = INV[INV["LOTE"] == l_]
                filas_ec.append({"LOTE": int(l_), "PRODUCTOS": len(prod_l),
                                 "UNIDADES": int(prod_l["CANTIDAD"].sum()),
                                 "MERCANCÍA": prod_l["INVERSION"].sum()})
            st.dataframe(pd.DataFrame(filas_ec), width="stretch", hide_index=True,
                         column_config={"MERCANCÍA": st.column_config.NumberColumn(
                             format="$%.2f")})
        lotes_inv = sorted({int(x) for x in INV["LOTE"].dropna()}, reverse=True)
        with st.form("form_lote", border=False):
            c = st.columns([1, 2, 1])
            l_sel = c[0].selectbox("Lote", lotes_inv) if lotes_inv else None
            accion = c[1].radio("Estado", ["Llegó", "En camino"], horizontal=True)
            llegada = c[2].date_input("Fecha de llegada", value=hoy(), format="DD/MM/YYYY")
            if st.form_submit_button("Guardar estado del lote") and l_sel is not None:
                if accion == "Llegó":
                    set_lote(l_sel, "RECIBIDO", llegada)
                    avisar(f"Lote {l_sel} recibido el {llegada:%d/%m/%Y}: ya sale en el "
                           "catálogo.")
                else:
                    set_lote(l_sel, "EN CAMINO")
                    avisar(f"Lote {l_sel} en camino: sus productos no salen en el catálogo.")
                st.rerun()

    with st.expander("🛍️ Catálogo para clientes"):
        url_cat = str(_cfg.get("CATALOGO_URL") or "")
        if url_cat:
            st.markdown(f"Tu catálogo: [{url_cat}]({url_cat})")
        visibles = INV[(INV["EXISTENCIA"] > 0) & (INV["PRECIO_VENTA"] > 0) &
                       (INV["OCULTAR"].astype(str).str.upper() != "SI") & (~INV["EN_CAMINO"])]
        con_foto = visibles["FOTO"].astype(str).str.startswith("http").sum()
        st.caption(f"Se ven {len(visibles)} productos (con existencia, no ocultos y de lotes "
                   f"que ya llegaron); {con_foto} con foto. Para agregar fotos u ocultar algo, usa las "
                   "columnas FOTO y OCULTAR en «✏️ Editar inventario».")
        with st.form("cfg_catalogo", border=False):
            c = st.columns(2)
            nom_c = c[0].text_input("Nombre que ven los clientes",
                                    value=str(_cfg.get("CATALOGO_NOMBRE") or "Catálogo"))
            wa_c = c[1].text_input("WhatsApp donde te escriben",
                                   value=str(_cfg.get("CATALOGO_WHATSAPP") or ""),
                                   placeholder="0414-1234567")
            c = st.columns([3, 1])
            url_n = c[0].text_input("Dirección del catálogo (cuando lo publiques)",
                                    value=url_cat, placeholder="https://tu-catalogo.streamlit.app")
            mostrar_usdt = c[1].checkbox(
                "Precio en Bs a tasa USDT",
                value=str(_cfg.get("CATALOGO_PRECIO_USDT") or "").upper() == "SI",
                help="Precio en dólares × tasa USDT de Binance P2P, igual que el bot.")
            t_usdt = tasa_usdt()
            st.caption(f"Tasa USDT ahora: {t_usdt:,.2f} Bs (Binance P2P). Un producto de $30 "
                       f"saldría a Bs {30 * t_usdt:,.2f}." if t_usdt
                       else "No se pudo consultar Binance ahora; el catálogo lo reintenta solo.")
            if st.form_submit_button("Guardar ajustes del catálogo"):
                if wa_c.strip() and not wa_numero(wa_c):
                    st.error("Ese número de WhatsApp no es válido.")
                else:
                    set_config(CATALOGO_NOMBRE=nom_c.strip() or "Catálogo",
                               CATALOGO_WHATSAPP=wa_c.strip(), CATALOGO_URL=url_n.strip(),
                               CATALOGO_PRECIO_USDT="SI" if mostrar_usdt else "NO")
                    avisar("Ajustes del catálogo guardados.")
                    st.rerun()

    with st.expander("✏️ Editar inventario"):
        st.caption("EXISTENCIA y ganancias se calculan solas contando VENTAS por NOMBRE. "
                   "FOTO: enlace a la imagen (https://…). OCULTAR: SI para que no salga en "
                   "el catálogo.")
        edit = st.data_editor(
            inv_raw, num_rows="dynamic", width="stretch", height=400,
            key="ed_inv",
            column_config={
                "CANTIDAD": st.column_config.NumberColumn(format="%d"),
                "PRECIO_COMPRA": st.column_config.NumberColumn(format="%.2f"),
                "PRECIO_VENTA": st.column_config.NumberColumn(format="%.2f"),
                "LOTE": st.column_config.NumberColumn(format="%d"),
                "FOTO": st.column_config.TextColumn(help="Enlace https a la imagen del producto."),
                "OCULTAR": st.column_config.SelectboxColumn(options=["", "SI"]),
            },
        )
        if st.button("💾 Guardar inventario", type="primary"):
            guardar("INVENTARIO", edit)
            avisar("Inventario actualizado.")
            st.rerun()

# --------------------------------------------------------------------------
# ENCARGOS Y PRESTAMOS
# --------------------------------------------------------------------------
elif seccion == "🎁 Por cobrar":
    st.header("Encargos y préstamos")
    st.caption("Encargo: compras algo para alguien y le cobras más de lo que te costó. "
               "Préstamo: sale plata de SHEIN y vuelve igual. Como pagan por partes, "
               "lo que importa es ABONADO contra COBRAR. Los totales van en dólares; "
               "lo que esté en bolívares se convierte con la tasa del día.")

    pend = PEND
    con_c = pend[pend["COSTO_CONOCIDO"]] if len(pend) else pend
    k = st.columns(4)
    k[0].metric("Saldo por cobrar", money(pend["SALDO_USD"].sum() if len(pend) else 0),
                f"{len(pend)} abiertos")
    k[1].metric("Ya abonado", money(ENC["ABONADO_USD"].sum() if len(ENC) else 0),
                f"de {money(ENC['COBRAR_USD'].sum() if len(ENC) else 0)} en total")
    k[2].metric("Ganancia pendiente",
                money(con_c["GANANCIA_USD"].sum() - con_c["GAN_COBRADA_USD"].fillna(0).sum()
                      if len(con_c) else 0))
    k[3].metric("Ganancia ya cobrada",
                money(ENC["GAN_COBRADA_USD"].fillna(0).sum() if len(ENC) else 0))

    if len(pend):
        por_tipo = pend.groupby("TIPO")["SALDO_USD"].agg(["size", "sum"])
        st.caption(" · ".join(f"{tp}: {int(r['size'])} por {money(r['sum'])}"
                              for tp, r in por_tipo.iterrows()))
        if pend["SIN_TASA"].any():
            n_bs = int(pend["SIN_TASA"].sum())
            st.warning(f"{n_bs} cuentas en bolívares no entran a los totales: falta la tasa "
                       "del día en la barra lateral.", icon="💱")
        faltan = pend[(~pend["COSTO_CONOCIDO"]) | (pend["TIPO"] == "SIN CLASIFICAR")]
        if len(faltan):
            st.warning(
                f"{len(faltan)} filas están incompletas: les falta el tipo o el costo. "
                "En un préstamo sin interés, COSTO va igual a COBRAR.", icon="✏️")

    # ---------------- antigüedad ----------------
    if len(pend):
        st.subheader("⏳ Antigüedad de las deudas")
        hoy_ts = pd.Timestamp(hoy())
        per = (pend.groupby("PERSONA")
               .agg(SALDO=("SALDO_USD", "sum"), CUENTAS=("SALDO_USD", "size"),
                    ABIERTA_DESDE=("FECHA", "min"))
               .reset_index())
        # último abono de la persona en cualquiera de sus cuentas, saldadas o no:
        # una cuenta nueva no reinicia el reloj, un pago sí
        ult = ENC.groupby("PERSONA")["ULT_ABONO"].max()
        per["ULTIMO_ABONO"] = per["PERSONA"].map(ult)
        desde = per["ULTIMO_ABONO"].where(per["ULTIMO_ABONO"].notna(), per["ABIERTA_DESDE"])
        per["DIAS"] = (hoy_ts - pd.to_datetime(desde)).dt.days
        # nunca abonó y alguna cuenta abierta no tiene fecha: la antigüedad real se desconoce
        sin_f = pend.groupby("PERSONA")["FECHA"].apply(lambda s: s.isna().any())
        desconocida = per["ULTIMO_ABONO"].isna() & per["PERSONA"].map(sin_f).fillna(False)
        per.loc[desconocida, "DIAS"] = float("nan")

        def semaforo(d):
            if pd.isna(d):
                return "⚪ sin fecha"
            if d <= 30:
                return "🟢 al día"
            if d <= 60:
                return "🟡 atento"
            return "🔴 cobrar"

        per["ESTADO"] = per["DIAS"].map(semaforo)
        per = per.sort_values(["DIAS", "SALDO"], ascending=[False, False], na_position="first")

        def mensaje(persona_m):
            cs = pend[pend["PERSONA"] == persona_m]
            dias = per.loc[per["PERSONA"] == persona_m, "DIAS"].iloc[0]
            try:
                return MSG_COBRO.format(
                    nombre=persona_m.title(),
                    saldo=montos_por_moneda(cs["SALDO"], cs["MONEDA"]),
                    abonado=montos_por_moneda(cs["ABONADO"], cs["MONEDA"]),
                    dias="" if pd.isna(dias) else int(dias))
            except (KeyError, IndexError, ValueError):
                return MSG_COBRO_DEFECTO.format(
                    nombre=persona_m.title(),
                    saldo=montos_por_moneda(cs["SALDO"], cs["MONEDA"]),
                    abonado=montos_por_moneda(cs["ABONADO"], cs["MONEDA"]))

        per["WHATSAPP"] = [wa_link(TELEFONOS.get(p_.upper()), mensaje(p_))
                           for p_ in per["PERSONA"]]
        st.dataframe(
            per[["ESTADO", "PERSONA", "SALDO", "DIAS", "ULTIMO_ABONO", "CUENTAS", "WHATSAPP"]],
            width="stretch", hide_index=True, height=min(38 * (len(per) + 1), 330),
            column_config={
                "SALDO": st.column_config.NumberColumn("Debe ($)", format="$%.2f"),
                "DIAS": st.column_config.NumberColumn("Días sin abonar", format="%d"),
                "ULTIMO_ABONO": st.column_config.DateColumn("Último abono",
                                                            format="DD/MM/YYYY"),
                "WHATSAPP": st.column_config.LinkColumn("Cobrar", display_text="📲 WhatsApp"),
            },
        )
        sin_tel = [p_ for p_ in per["PERSONA"] if not wa_numero(TELEFONOS.get(p_.upper()))]
        if sin_tel:
            st.caption(f"Sin teléfono: {', '.join(sin_tel)}. Su botón abre WhatsApp para "
                       "que elijas el contacto; guarda el número en «📇 Teléfonos y mensaje».")

        with st.expander("📇 Teléfonos y mensaje de cobro"):
            todas = sorted(set(ENC["PERSONA"]) | set(contactos["PERSONA"]))
            base_c = pd.DataFrame({"PERSONA": todas})
            base_c["TELEFONO"] = base_c["PERSONA"].map(
                lambda x: TELEFONOS.get(str(x).upper(), "")).fillna("").astype(str)
            ed_c = st.data_editor(base_c, width="stretch", hide_index=True, key="ed_tel",
                                  disabled=["PERSONA"],
                                  column_config={"TELEFONO": st.column_config.TextColumn(
                                      help="Como lo escribas: 0414-1234567 sirve.")})
            plantilla = st.text_area("Mensaje", value=MSG_COBRO, height=110,
                                     help="Puedes usar {nombre}, {saldo}, {abonado} y {dias}.")
            if st.button("💾 Guardar teléfonos y mensaje", key="btn_tel"):
                malos = [r.PERSONA for r in ed_c.itertuples()
                         if str(r.TELEFONO).strip() and not wa_numero(r.TELEFONO)]
                if malos:
                    st.error("Teléfono no válido para: " + ", ".join(malos))
                else:
                    try:
                        plantilla.format(nombre="", saldo="", abonado="", dias="")
                    except (KeyError, IndexError, ValueError):
                        st.error("El mensaje tiene llaves { } que no reconozco. Usa solo "
                                 "{nombre}, {saldo}, {abonado} y {dias}.")
                    else:
                        guardar("CONTACTOS", ed_c[ed_c["TELEFONO"].astype(str).str.strip() != ""])
                        set_config(MSG_COBRO=plantilla)
                        avisar("Teléfonos y mensaje guardados.")
                        st.rerun()
        st.caption("Días desde su último abono; si nunca abonó, desde su cuenta abierta más "
                   "vieja. Prestarle de nuevo no reinicia el reloj. "
                   "🟢 hasta 30 días · 🟡 31 a 60 · 🔴 más de 60.")
        sin_fecha = pend[pend["FECHA"].isna() & pend["ULT_ABONO"].isna()]
        if len(sin_fecha):
            st.info(f"{len(sin_fecha)} cuentas no tienen fecha, así que no se puede medir su "
                    "antigüedad. Lo ideal es poner la fecha aproximada en «✏️ Editar»; si no "
                    "la recuerdas, empieza a contar desde hoy.", icon="📅")
            if st.button("📅 Empezar a contar desde hoy", key="btn_fecha_hoy"):
                df = leer("ENCARGOS")
                vacias = df["FECHA"].isna() & df.index.isin(sin_fecha.index)
                df.loc[vacias, "FECHA"] = pd.Timestamp(hoy())
                guardar("ENCARGOS", df)
                avisar(f"{int(vacias.sum())} cuentas ahora cuentan desde hoy.")
                st.rerun()

    # ---------------- abono ----------------
    if len(pend):
        st.subheader("Registrar un abono")
        filas = []
        for persona_i, g in pend.groupby("PERSONA"):
            mon = g["MONEDA"].unique()
            if len(mon) == 1:
                etiqueta = f"{persona_i} · debe {plata(g['SALDO'].sum(), mon[0])}"
            else:
                etiqueta = f"{persona_i} · debe ≈ {money(g['SALDO_USD'].sum())} (mixto)"
            if len(g) > 1:
                etiqueta += f" · {len(g)} cuentas"
            filas.append((g["SALDO_USD"].sum(), etiqueta, persona_i))
        filas.sort(key=lambda x: (-(x[0] if pd.notna(x[0]) else 0), x[1]))
        etiquetas = {e: p_ for _, e, p_ in filas}
        sel = st.selectbox("Persona", list(etiquetas.keys()), label_visibility="collapsed",
                           key="abono_persona")
        quien = etiquetas[sel]
        cuentas = pend[pend["PERSONA"] == quien].sort_values("FECHA", na_position="first")
        monedas = cuentas["MONEDA"].unique()

        destino = "TODAS"
        if len(cuentas) > 1:
            ops = {}
            if len(monedas) == 1:
                ops["Repartir entre sus cuentas, de la más antigua primero"] = "TODAS"
            for n, r in enumerate(cuentas.itertuples(), 1):
                fecha = f" · {r.FECHA:%d/%m/%Y}" if pd.notna(r.FECHA) else ""
                ops[f"Cuenta {n}: {r.TIPO.lower()}"
                    + (f" · {r.DESCRIPCION}" if r.DESCRIPCION else "")
                    + fecha + f" · debe {plata(r.SALDO, r.MONEDA)}"] = r.Index
            destino = ops[st.selectbox("Aplicar a", list(ops.keys()), key="abono_destino")]
            if len(monedas) > 1:
                st.caption("Tiene cuentas en dólares y en bolívares: elige a cuál va el abono.")

        mon_ab = monedas[0] if destino == "TODAS" else ENC.loc[destino, "MONEDA"]
        saldo = (_num(cuentas["SALDO"].sum()) if destino == "TODAS"
                 else _num(ENC.loc[destino, "SALDO"]))

        with st.form("abono", clear_on_submit=True):
            c = st.columns([2, 1, 1])
            monto_txt = campo_dinero(c[0], f"Abono ({mon_ab})",
                                     ayuda=f"Debe {plata(saldo, mon_ab)}. Marca 'Saldó todo' "
                                           "si te pagó completo.")
            fecha_ab = c[1].date_input("Fecha del abono", value=hoy(),
                                       format="DD/MM/YYYY")
            todo = c[2].checkbox("Saldó todo")
            ubic = selector_caja(st, "¿Dónde entró el abono?", "caja_abono")
            if st.form_submit_button("Registrar abono", type="primary"):
                monto = leer_monto(monto_txt)
                pago = saldo if todo else monto
                delta, _, err = plan_caja(ubic, pago or 0, mon_ab)
                if pago is None:
                    st.error(f"Monto inválido: «{monto_txt}». Escribe por ejemplo 10 o 10,50.")
                elif pago <= 0:
                    st.error("El abono debe ser mayor que cero.")
                elif pago > saldo + 0.004:
                    st.error(f"Te está abonando {plata(pago, mon_ab)} pero solo debe "
                             f"{plata(saldo, mon_ab)}.")
                elif err:
                    st.error(err)
                else:
                    df = leer("ENCARGOS")
                    df["ABONADO"] = pd.to_numeric(df["ABONADO"], errors="coerce") \
                        .fillna(0).astype(float)
                    f_ab = pd.Timestamp(fecha_ab)
                    tocadas = []
                    if destino == "TODAS":
                        resto = pago
                        for idx in cuentas.index:
                            if resto <= 0.004:
                                break
                            aplica = min(resto, _num(ENC.loc[idx, "SALDO"]))
                            df.loc[idx, "ABONADO"] = _num(df.loc[idx, "ABONADO"]) + aplica
                            tocadas.append(idx)
                            resto -= aplica
                    else:
                        df.loc[destino, "ABONADO"] = _num(df.loc[destino, "ABONADO"]) + pago
                        tocadas.append(destino)
                    for idx in tocadas:
                        previo = df.loc[idx, "ULT_ABONO"]
                        if pd.isna(previo) or f_ab > pd.Timestamp(previo):
                            df.loc[idx, "ULT_ABONO"] = f_ab
                    guardar("ENCARGOS", df)
                    extra = aplicar_caja(ubic, delta, "Abono", quien)
                    resta = saldo - pago
                    avisar(f"Abono de {plata(pago, mon_ab)} de {quien}. "
                           + ("Queda al día." if resta <= 0.004
                              else f"Le queda debiendo {plata(resta, mon_ab)}.")
                           + (extra or " No se movió la caja."))
                    st.rerun()

    # ---------------- alta ----------------
    st.divider()
    tipo = st.radio("Registrar algo nuevo", ["Encargo", "Préstamo"], horizontal=True)
    with st.expander(f"➕ Nuevo {tipo.lower()}", expanded=ENC.empty):
        conocidas = sorted(x for x in ENC["PERSONA"].unique() if x) if len(ENC) else []
        c = st.columns([2, 3])
        elegida = c[0].selectbox("Persona", ["➕ Nueva persona"] + conocidas,
                                 key=f"sel_persona_{tipo}")
        if elegida == "➕ Nueva persona":
            persona = c[1].text_input("Nombre de la persona nueva",
                                      key=f"nueva_persona_{tipo}").strip().upper()
        else:
            persona = elegida
            abiertas = ENC[(ENC["PERSONA"] == persona) & (~ENC["ESTA_PAGADO"])]
            if len(abiertas):
                c[1].caption(f"Ya debe {money(abiertas['SALDO_USD'].sum())} "
                             f"en {len(abiertas)} cuenta(s).")

        if tipo == "Encargo":
            with st.form("nuevo_encargo", clear_on_submit=True):
                c = st.columns([3, 1])
                desc = c[0].text_input("Qué pidió")
                mon = c[1].selectbox("Moneda", ["$", "Bs"], key="mon_enc")
                c = st.columns(4)
                costo_txt = campo_dinero(c[0], "Costo (lo que gastaste)")
                cobrar_txt = campo_dinero(c[1], "Cobrar")
                abonado_txt = campo_dinero(c[2], "Ya abonó")
                fecha = c[3].date_input("Fecha", value=hoy(), format="DD/MM/YYYY")
                c = st.columns([2, 3])
                ubic = selector_caja(c[0], "Caja", "caja_enc")
                c[1].caption("Del lugar elegido sale el costo y entra lo que ya abonó.")
                nota = st.text_input("Nota")
                if st.form_submit_button("Guardar encargo", type="primary"):
                    costo = leer_monto(costo_txt)
                    cobrar = leer_monto(cobrar_txt)
                    abonado = leer_monto(abonado_txt)
                    neto = (abonado or 0) - (costo or 0)
                    delta, _, err = plan_caja(ubic, neto, mon)
                    if not persona:
                        st.error("Falta la persona.")
                    elif None in (costo, cobrar, abonado):
                        st.error("Revisa los montos: escribe por ejemplo 40 o 40,50.")
                    elif cobrar <= 0:
                        st.error("Falta el monto a cobrar.")
                    elif abonado > cobrar + 0.004:
                        st.error("Lo abonado no puede pasar de lo que se cobra.")
                    elif err:
                        st.error(err)
                    else:
                        agregar("ENCARGOS", {
                            "PERSONA": persona, "TIPO": "ENCARGO", "DESCRIPCION": desc,
                            "COSTO": costo or None, "COBRAR": cobrar, "ABONADO": abonado,
                            "MONEDA": mon, "FECHA": pd.Timestamp(fecha),
                            "ULT_ABONO": pd.Timestamp(fecha) if abonado else None,
                            "NOTA": nota,
                        })
                        extra = aplicar_caja(ubic, delta, "Encargo", f"{persona}: {desc}")
                        avisar(f"Encargo de {persona}: cobras {plata(cobrar, mon)}"
                               + (f", ganas {plata(cobrar - costo, mon)}" if costo else "")
                               + "." + extra)
                        st.rerun()
        else:
            with st.form("nuevo_prestamo", clear_on_submit=True):
                c = st.columns([1, 1, 1])
                prestado_txt = campo_dinero(c[0], "Monto prestado")
                cobrar_txt = campo_dinero(c[1], "A cobrar",
                                          ayuda="Déjalo vacío si no cobras interés.")
                mon = c[2].selectbox("Moneda", ["$", "Bs"], key="mon_pre")
                c = st.columns([1, 1, 3])
                abonado_txt = campo_dinero(c[0], "Ya abonó")
                fecha = c[1].date_input("Fecha", value=hoy(), format="DD/MM/YYYY")
                nota = c[2].text_input("Nota")
                ubic = selector_caja(st, "¿De dónde salió la plata prestada?", "caja_pre")
                if st.form_submit_button("Guardar préstamo", type="primary"):
                    prestado = leer_monto(prestado_txt)
                    cobrar = leer_monto(cobrar_txt)
                    abonado = leer_monto(abonado_txt)
                    neto = (abonado or 0) - (prestado or 0)
                    delta, _, err = plan_caja(ubic, neto, mon)
                    total = (cobrar if cobrar and cobrar > 0 else prestado) or 0
                    if not persona:
                        st.error("Falta la persona.")
                    elif None in (prestado, cobrar, abonado):
                        st.error("Revisa los montos: escribe por ejemplo 50 o 50,25.")
                    elif prestado <= 0:
                        st.error("Falta el monto prestado.")
                    elif abonado > total + 0.004:
                        st.error("Lo abonado no puede pasar de lo que se cobra.")
                    elif err:
                        st.error(err)
                    else:
                        agregar("ENCARGOS", {
                            "PERSONA": persona, "TIPO": "PRÉSTAMO", "DESCRIPCION": "",
                            "COSTO": prestado, "COBRAR": total, "ABONADO": abonado,
                            "MONEDA": mon, "FECHA": pd.Timestamp(fecha),
                            "ULT_ABONO": pd.Timestamp(fecha) if abonado else None,
                            "NOTA": nota,
                        })
                        extra = aplicar_caja(ubic, delta, "Préstamo", persona)
                        dif = total - prestado
                        avisar(f"Préstamo a {persona} por {plata(prestado, mon)}"
                               + (f", cobras {plata(dif, mon)} de más" if dif else "")
                               + "." + extra)
                        st.rerun()

    if ENC.empty:
        st.info("Todavía no hay nada cargado.")
        st.stop()

    # ---------------- listado ----------------
    st.divider()
    c1, c2, c3 = st.columns([2, 1, 1])
    busq = c1.text_input("Buscar persona").strip().upper()
    f_tipo = c2.multiselect("Tipo", sorted(ENC["TIPO"].unique()))
    estado = c3.selectbox("Estado", ["Con saldo", "Saldados", "Todos"])

    e = ENC.copy()
    if busq:
        e = e[e["PERSONA"].str.contains(busq, na=False)]
    if f_tipo:
        e = e[e["TIPO"].isin(f_tipo)]
    if estado == "Con saldo":
        e = e[~e["ESTA_PAGADO"]]
    elif estado == "Saldados":
        e = e[e["ESTA_PAGADO"]]

    detalle = st.checkbox("Ver cuenta por cuenta", value=False)
    if detalle:
        st.dataframe(
            e[["PERSONA", "TIPO", "DESCRIPCION", "MONEDA", "COSTO", "COBRAR", "ABONADO",
               "SALDO", "GANANCIA", "FECHA", "ULT_ABONO", "NOTA"]]
            .sort_values(["SALDO", "PERSONA"], ascending=False),
            width="stretch", hide_index=True, height=360,
            column_config={
                "COSTO": st.column_config.NumberColumn(format="%.2f"),
                "COBRAR": st.column_config.NumberColumn(format="%.2f"),
                "ABONADO": st.column_config.NumberColumn(format="%.2f"),
                "SALDO": st.column_config.NumberColumn(format="%.2f"),
                "GANANCIA": st.column_config.NumberColumn(format="%.2f"),
                "FECHA": st.column_config.DateColumn(format="DD/MM/YYYY"),
                "ULT_ABONO": st.column_config.DateColumn("ÚLT. ABONO", format="DD/MM/YYYY"),
            },
        )
        st.caption("En esta vista cada monto va en la moneda de su cuenta (columna MONEDA).")
    elif len(e):
        tot = (e.groupby("PERSONA")
               .agg(CUENTAS=("COBRAR_USD", "size"), COBRAR=("COBRAR_USD", "sum"),
                    ABONADO=("ABONADO_USD", "sum"), SALDO=("SALDO_USD", "sum"),
                    GANANCIA=("GANANCIA_USD", "sum"))
               .reset_index().sort_values("SALDO", ascending=False))
        tipos = e.groupby("PERSONA")["TIPO"].agg(lambda s: " + ".join(sorted(set(s))))
        monedas_p = e.groupby("PERSONA")["MONEDA"].agg(lambda s: " + ".join(sorted(set(s))))
        tot["TIPO"] = tot["PERSONA"].map(tipos)
        tot["MONEDA"] = tot["PERSONA"].map(monedas_p)
        tot = tot[["PERSONA", "TIPO", "MONEDA", "CUENTAS", "COBRAR", "ABONADO", "SALDO",
                   "GANANCIA"]]
        st.dataframe(
            tot, width="stretch", hide_index=True, height=360,
            column_config={
                "COBRAR": st.column_config.NumberColumn("COBRAR ($)", format="$%.2f"),
                "ABONADO": st.column_config.NumberColumn("ABONADO ($)", format="$%.2f"),
                "SALDO": st.column_config.NumberColumn("SALDO ($)", format="$%.2f"),
                "GANANCIA": st.column_config.NumberColumn("GANANCIA ($)", format="$%.2f"),
            },
        )
        st.caption(f"{len(tot)} personas · {len(e)} cuentas · saldo "
                   f"{money(tot['SALDO'].sum())}. Todo en dólares; los bolívares se "
                   "convierten con la tasa del día.")
        graf = tot[tot["SALDO"].notna()]
        if len(graf) > 1:
            st.plotly_chart(
                px.bar(graf.sort_values("SALDO"), x="SALDO", y="PERSONA", orientation="h",
                       labels={"SALDO": "$", "PERSONA": ""}),
                width="stretch")
    else:
        st.info("Nada que mostrar con esos filtros.")

    # ---------------- eliminar ----------------
    claves = ["PERSONA", "TIPO", "DESCRIPCION", "COSTO", "COBRAR", "MONEDA", "FECHA"]
    dup = ENC[ENC.duplicated(subset=claves, keep=False)]
    if len(dup):
        nombres = ", ".join(sorted(dup["PERSONA"].unique()))
        st.warning(f"Posibles cuentas duplicadas: {nombres}. Revísalas abajo en "
                   "«Eliminar una cuenta».", icon="⚠️")

    with st.expander("🗑️ Eliminar una cuenta", expanded=len(dup) > 0):
        st.caption("Para borrar algo que se registró dos veces o por error. "
                   "No se puede deshacer, salvo desde el historial de versiones del Sheet.")
        lista = ENC.sort_values(["PERSONA", "FECHA"], na_position="first")
        ops = {}
        for n, r in enumerate(lista.itertuples(), 1):
            marca = "⚠️ " if r.Index in dup.index else ""
            fecha = f" · {r.FECHA:%d/%m/%Y}" if pd.notna(r.FECHA) else " · sin fecha"
            ops[f"{marca}#{n} · {r.PERSONA} · {r.TIPO.lower()}"
                + (f" · {r.DESCRIPCION}" if r.DESCRIPCION else "")
                + f" · cobrar {plata(r.COBRAR, r.MONEDA)} · abonado "
                  f"{plata(r.ABONADO, r.MONEDA)}{fecha}"] = r.Index
        primero = next((i for i, k_ in enumerate(ops) if k_.startswith("⚠️")), 0)
        elegida = st.selectbox("Cuenta a eliminar", list(ops.keys()), index=primero,
                               key="borrar_enc")
        idx_b = ops[elegida]
        if _num(ENC.loc[idx_b, "ABONADO"]) > 0:
            st.info(f"Esta cuenta tiene {plata(ENC.loc[idx_b, 'ABONADO'], ENC.loc[idx_b, 'MONEDA'])} "
                    "abonados. Si la otra copia no los tiene, borra mejor la que está en cero.")
        ubic_b = selector_caja(st, "Revertir en caja", "caja_borrar_enc",
                               mover_por_defecto=False)
        st.caption("Si al registrarla movió la caja, elige la misma ubicación: se devuelve el "
                   "costo y se saca lo abonado. Si fue un duplicado, deja «No mover».")
        conf = st.checkbox("Sí, quiero borrar esta cuenta", key="conf_borrar_enc")
        if st.button("🗑️ Eliminar", disabled=not conf, key="btn_borrar_enc"):
            fila = ENC.loc[idx_b]
            revertir = _num(fila["COSTO"]) - _num(fila["ABONADO"])
            delta, _, err = plan_caja(ubic_b, revertir, fila["MONEDA"])
            if err:
                st.error(err)
            else:
                df = leer("ENCARGOS")
                guardar("ENCARGOS", df.drop(index=idx_b))
                extra = aplicar_caja(ubic_b, delta, "Reversión", f"cuenta eliminada: {elegida}")
                avisar("Cuenta eliminada." + extra)
                st.rerun()

    with st.expander("✏️ Editar"):
        edit = st.data_editor(
            enc, num_rows="dynamic", width="stretch", height=400, key="ed_enc",
            column_config={
                "TIPO": st.column_config.SelectboxColumn(options=["ENCARGO", "PRÉSTAMO"]),
                "COSTO": st.column_config.NumberColumn(format="%.2f"),
                "COBRAR": st.column_config.NumberColumn(format="%.2f"),
                "ABONADO": st.column_config.NumberColumn(format="%.2f"),
                "MONEDA": st.column_config.SelectboxColumn(options=["$", "Bs"]),
                "FECHA": st.column_config.DateColumn(format="DD/MM/YYYY"),
                "ULT_ABONO": st.column_config.DateColumn("ÚLT. ABONO", format="DD/MM/YYYY"),
            },
        )
        st.caption("Editar aquí no mueve la caja: es para corregir datos.")
        if st.button("💾 Guardar", type="primary"):
            guardar("ENCARGOS", edit)
            avisar("Actualizado.")
            st.rerun()

# --------------------------------------------------------------------------
# ROTACION
# --------------------------------------------------------------------------
elif seccion == "📈 Rotación":
    st.header("Rotación: qué recomprar")
    st.caption("Qué se vende rápido y deja margen, qué lote rindió y qué está estancado. "
               "Usa las ventas de la etapa actual que tienen su producto en Inventario.")

    vv = VEN[VEN["EN_INVENTARIO"] & VEN["FECHA"].notna()]
    if vv.empty or INV.empty:
        st.info("Hace falta inventario y ventas registradas para calcular la rotación.")
        st.stop()
    hoy_ts = pd.Timestamp(hoy())

    # ---- por producto
    fechas_p = vv.groupby("NOMBRE")["FECHA"].agg(PRIMERA="min", ULTIMA="max")
    pr = INV.set_index("NOMBRE").join(fechas_p, how="left").reset_index()
    pr["TIPO_PROD"] = pr["NOMBRE"].astype(str).str.split().str[0].str.upper()
    fin = pr["ULTIMA"].where(pr["EXISTENCIA"] <= 0, hoy_ts)
    semanas = ((fin - pr["PRIMERA"]).dt.days / 7).clip(lower=1)
    pr["POR_SEMANA"] = (pr["VENDIDAS"] / semanas).where(pr["VENDIDAS"] > 0, 0.0)
    pr["MARGEN_PCT"] = (pr["MARGEN_UNIT"] / pr["PRECIO_VENTA"] * 100) \
        .where(pr["PRECIO_VENTA"] > 0)
    pr["DIAS_SIN_VENDER"] = (hoy_ts - pr["ULTIMA"]).dt.days

    t1, t2, t3, t4 = st.tabs(["🔁 Recomprar", "🏷️ Por tipo", "📦 Por lote", "🐢 Estancado"])

    with t1:
        st.caption("Agotados, ordenados por lo rápido que se fueron y lo que dejaron. "
                   "Son los mejores candidatos para el próximo pedido.")
        ag = pr[(pr["EXISTENCIA"] <= 0) & (pr["VENDIDAS"] > 0)].copy()
        if ag.empty:
            st.info("No hay productos agotados todavía.")
        else:
            ag["PUNTAJE"] = ag["POR_SEMANA"] * ag["MARGEN_UNIT"].clip(lower=0)
            ag = ag.sort_values("PUNTAJE", ascending=False)
            st.dataframe(
                ag[["NOMBRE", "VENDIDAS", "POR_SEMANA", "PRECIO_COMPRA", "PRECIO_VENTA",
                    "MARGEN_PCT", "GANANCIA_REAL", "ULTIMA", "LOTE"]],
                width="stretch", hide_index=True, height=400,
                column_config={
                    "POR_SEMANA": st.column_config.NumberColumn("Por semana", format="%.1f"),
                    "PRECIO_COMPRA": st.column_config.NumberColumn("Costo", format="$%.2f"),
                    "PRECIO_VENTA": st.column_config.NumberColumn("Precio", format="$%.2f"),
                    "MARGEN_PCT": st.column_config.NumberColumn("Margen", format="%.0f%%"),
                    "GANANCIA_REAL": st.column_config.NumberColumn("Ganó", format="$%.2f"),
                    "ULTIMA": st.column_config.DateColumn("Se agotó", format="DD/MM/YYYY"),
                    "LOTE": st.column_config.NumberColumn(format="%d"),
                })
            st.caption("Por semana: unidades vendidas entre su primera y su última venta. "
                       "El orden combina velocidad y ganancia por unidad.")

    with t2:
        st.caption("Agrupado por la primera palabra del nombre (ANILLO, COLLAR, PULSERA…).")
        tp = (pr.groupby("TIPO_PROD")
              .agg(PRODUCTOS=("NOMBRE", "size"), COMPRADAS=("CANTIDAD", "sum"),
                   VENDIDAS=("VENDIDAS", "sum"), EN_STOCK=("EXISTENCIA", "sum"),
                   GANANCIA=("GANANCIA_REAL", "sum"), POR_SEMANA=("POR_SEMANA", "sum"),
                   INGRESO=("INGRESO_REAL", "sum"))
              .reset_index())
        tp["VENDIDO_PCT"] = tp["VENDIDAS"] / tp["COMPRADAS"] * 100
        tp["MARGEN_PCT"] = (tp["GANANCIA"] / tp["INGRESO"] * 100).where(tp["INGRESO"] > 0)
        tp = tp.sort_values("GANANCIA", ascending=False)
        c1, c2 = st.columns([3, 2])
        with c1:
            st.dataframe(
                tp[["TIPO_PROD", "PRODUCTOS", "COMPRADAS", "VENDIDAS", "VENDIDO_PCT",
                    "EN_STOCK", "MARGEN_PCT", "GANANCIA"]],
                width="stretch", hide_index=True, height=420,
                column_config={
                    "TIPO_PROD": "Tipo",
                    "VENDIDO_PCT": st.column_config.NumberColumn("% vendido", format="%.0f%%"),
                    "MARGEN_PCT": st.column_config.NumberColumn("Margen", format="%.0f%%"),
                    "GANANCIA": st.column_config.NumberColumn("Ganó", format="$%.2f"),
                })
        with c2:
            top = tp.head(10).sort_values("GANANCIA")
            fig = px.bar(top, x="GANANCIA", y="TIPO_PROD", orientation="h",
                         labels={"GANANCIA": "$ ganado", "TIPO_PROD": ""})
            fig.update_layout(height=420)
            st.plotly_chart(fig, width="stretch")

    with t3:
        st.caption("Cuánto tardó cada lote en venderse y cuánto dejó después de sus gastos.")
        uni = INV.groupby("LOTE").agg(COMPRADAS=("CANTIDAD", "sum"),
                                     VENDIDAS=("VENDIDAS", "sum"),
                                     INVERTIDO=("INVERSION", "sum"))
        # cada venta cuenta para el lote de su producto (el lote escrito en la venta
        # a veces difiere), así unidades, ingresos y costo hablan del mismo lote
        lote_prod = INV.drop_duplicates("NOMBRE").set_index("NOMBRE")["LOTE"]
        vl = (vv.assign(LOTE=vv["NOMBRE"].map(lote_prod))
              .groupby("LOTE").agg(PRIMERA=("FECHA", "min"), ULTIMA=("FECHA", "max"),
                                   INGRESOS=("MONTO", "sum"), BRUTA=("GANANCIA", "sum")))
        gl = (GAS[GAS["LOTE"].notna()].groupby("LOTE")["MONTO_USD"].sum()
              .rename("GASTOS") if len(GAS) else pd.Series(dtype=float, name="GASTOS"))
        lt = uni.join(vl, how="left").join(gl, how="left").reset_index()
        lt = lt[lt["LOTE"].notna()]
        lt["GASTOS"] = lt["GASTOS"].fillna(0)
        lt["NETA"] = lt["BRUTA"].fillna(0) - lt["GASTOS"]
        lt["VENDIDO_PCT"] = lt["VENDIDAS"] / lt["COMPRADAS"] * 100
        completo = lt["VENDIDAS"] >= lt["COMPRADAS"]
        lt["LLEGADA"] = pd.to_datetime(lt["LOTE"].map(LLEGADAS), errors="coerce")
        inicio = lt["LLEGADA"].where(lt["LLEGADA"].notna(), lt["PRIMERA"])
        lt["DIAS"] = ((lt["ULTIMA"].where(completo, hoy_ts) - inicio).dt.days)
        lt["ESTADO"] = completo.map({True: "✅ vendido", False: "⏳ en venta"})
        lt.loc[lt["LOTE"].isin(EN_CAMINO), "ESTADO"] = "🚚 en camino"
        lt.loc[lt["LOTE"].isin(EN_CAMINO), "DIAS"] = float("nan")
        lt["LOTE"] = lt["LOTE"].astype(int)
        lt = lt.sort_values("LOTE", ascending=False)
        st.dataframe(
            lt[["LOTE", "ESTADO", "COMPRADAS", "VENDIDO_PCT", "DIAS", "INVERTIDO", "INGRESOS",
                "GASTOS", "NETA"]],
            width="stretch", hide_index=True, height=420,
            column_config={
                "LOTE": st.column_config.NumberColumn(format="%d"),
                "VENDIDO_PCT": st.column_config.NumberColumn("% vendido", format="%.0f%%"),
                "DIAS": st.column_config.NumberColumn("Días", format="%d",
                                                      help="Desde la primera venta: hasta "
                                                           "agotarse, o hasta hoy si sigue."),
                "INVERTIDO": st.column_config.NumberColumn(format="$%.2f"),
                "INGRESOS": st.column_config.NumberColumn(format="$%.2f"),
                "GASTOS": st.column_config.NumberColumn(format="$%.2f"),
                "NETA": st.column_config.NumberColumn("Ganancia neta", format="$%.2f"),
            })
        st.caption("Ganancia neta = lo vendido menos lo que costó esa mercancía, menos los "
                   "gastos asignados al lote (envíos y otros). Los días cuentan desde que el "
                   "lote llegó; si no marcaste la llegada, desde su primera venta.")

    with t4:
        lim = st.slider("Sin vender hace más de (días)", 15, 180, 60, step=15)
        est = pr[(pr["EXISTENCIA"] > 0) & (~pr["EN_CAMINO"]) &
                 ((pr["DIAS_SIN_VENDER"] > lim) | pr["ULTIMA"].isna())].copy()
        est["CAPITAL"] = est["EXISTENCIA"] * est["PRECIO_COMPRA"]
        est = est.sort_values("CAPITAL", ascending=False)
        if est.empty:
            st.success(f"Nada lleva más de {lim} días sin venderse.")
        else:
            st.caption(f"{len(est)} productos con {money(est['CAPITAL'].sum())} de mercancía "
                       "parada. Candidatos a rebaja, combo o regalo con compra.")
            st.dataframe(
                est[["NOMBRE", "EXISTENCIA", "DIAS_SIN_VENDER", "ULTIMA", "PRECIO_VENTA",
                     "CAPITAL", "LOTE"]],
                width="stretch", hide_index=True, height=380,
                column_config={
                    "DIAS_SIN_VENDER": st.column_config.NumberColumn("Días sin vender",
                                                                     format="%d"),
                    "ULTIMA": st.column_config.DateColumn("Última venta", format="DD/MM/YYYY"),
                    "PRECIO_VENTA": st.column_config.NumberColumn("Precio", format="$%.2f"),
                    "CAPITAL": st.column_config.NumberColumn("Mercancía parada",
                                                             format="$%.2f"),
                    "LOTE": st.column_config.NumberColumn(format="%d"),
                })
            st.caption("Sin «última venta» significa que nunca se ha vendido una unidad.")

# --------------------------------------------------------------------------
# GASTOS
# --------------------------------------------------------------------------
elif seccion == "🚚 Gastos":
    st.header("Gastos")
    st.caption("Envíos y otros gastos que se pagan de la caja. Cada uno descuenta la "
               "ubicación de donde salió y entra a la ganancia del Dashboard.")

    total = GAS["MONTO_USD"].sum() if len(GAS) else 0.0
    mes = pd.Timestamp(hoy()).to_period("M")
    del_mes = GAS[GAS["FECHA"].dt.to_period("M") == mes]["MONTO_USD"].sum() if len(GAS) else 0
    envios = GAS[GAS["CATEGORIA"] == "ENVÍO"]["MONTO_USD"].sum() if len(GAS) else 0
    k = st.columns(3)
    k[0].metric("Total en gastos", money(total), f"{len(GAS)} registros")
    k[1].metric("Este mes", money(del_mes))
    k[2].metric("En envíos", money(envios))
    if len(GAS) and GAS["MONTO_USD"].isna().any():
        st.warning("Hay gastos en bolívares que no entran al total: falta la tasa del día.",
                   icon="💱")

    with st.expander("➕ Registrar gasto", expanded=True):
        with st.form("nuevo_gasto", clear_on_submit=True):
            c = st.columns([1, 2, 1])
            cat = c[0].selectbox("Categoría", CATEGORIAS_GASTO)
            concepto = c[1].text_input("Concepto", placeholder="Ej. envío pedido SHEIN")
            ult_lote = _ent(INV["LOTE"].max() if len(INV) else 0)
            lote = c[2].number_input("Lote", min_value=0, step=1, value=ult_lote,
                                     help="0 si no corresponde a un lote.")
            c = st.columns([1, 1, 1, 2])
            monto_txt = campo_dinero(c[0], "Monto")
            mon = c[1].selectbox("Moneda", ["$", "Bs"])
            fecha = c[2].date_input("Fecha", value=hoy(), format="DD/MM/YYYY")
            ubic = selector_caja(c[3], "¿De dónde salió?", "caja_gasto")
            if st.form_submit_button("Guardar gasto", type="primary"):
                monto = leer_monto(monto_txt)
                delta, _, err = plan_caja(ubic, -(monto or 0), mon)
                if monto is None or monto <= 0:
                    st.error(f"Monto inválido: «{monto_txt}». Escribe por ejemplo 12 o 12,50.")
                elif err:
                    st.error(err)
                else:
                    agregar("GASTOS", {
                        "FECHA": pd.Timestamp(fecha), "CATEGORIA": cat, "CONCEPTO": concepto,
                        "LOTE": lote or None, "MONTO": monto, "MONEDA": mon,
                        "UBICACION": ubic or "",
                    })
                    extra = aplicar_caja(ubic, delta, "Gasto", f"{cat.lower()}: {concepto}")
                    avisar(f"Gasto de {plata(monto, mon)} ({cat.lower()}"
                           + (f", lote {lote}" if lote else "") + ")."
                           + (extra or " No se movió la caja."))
                    st.rerun()

    if GAS.empty:
        st.info("Todavía no hay gastos registrados.")
        st.stop()

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Por mes")
        pm = GAS.assign(MES=GAS["FECHA"].dt.to_period("M").astype(str))
        pm = pm.groupby(["MES", "CATEGORIA"])["MONTO_USD"].sum().reset_index()
        fig = px.bar(pm, x="MES", y="MONTO_USD", color="CATEGORIA",
                     labels={"MONTO_USD": "$", "MES": "", "CATEGORIA": ""})
        fig.update_layout(height=320, legend_title="")
        st.plotly_chart(fig, width="stretch")
    with c2:
        st.subheader("Envío sobre lo comprado, por lote")
        env_l = (GAS[(GAS["CATEGORIA"] == "ENVÍO") & GAS["LOTE"].notna()]
                 .groupby("LOTE")["MONTO_USD"].sum())
        compra_l = INV.groupby("LOTE")["INVERSION"].sum()
        if len(env_l):
            tl = pd.DataFrame({"ENVIO": env_l, "COMPRA": compra_l}).dropna(subset=["ENVIO"])
            tl["PCT"] = tl["ENVIO"] / tl["COMPRA"] * 100
            tl = tl.reset_index().rename(columns={"index": "LOTE"})
            tl["LOTE"] = tl["LOTE"].astype(int).astype(str)
            st.dataframe(tl, width="stretch", hide_index=True, height=280,
                         column_config={
                             "ENVIO": st.column_config.NumberColumn("Envío", format="$%.2f"),
                             "COMPRA": st.column_config.NumberColumn("Mercancía",
                                                                     format="$%.2f"),
                             "PCT": st.column_config.NumberColumn("% del lote", format="%.1f%%"),
                         })
            st.caption("Cuánto pesó el envío sobre lo que costó la mercancía de ese lote.")
        else:
            st.caption("Asigna un lote a los envíos para ver cuánto pesan sobre cada compra.")

    st.subheader("Registro")
    st.dataframe(
        GAS.sort_values("FECHA", ascending=False)[
            ["FECHA", "CATEGORIA", "CONCEPTO", "LOTE", "MONTO", "MONEDA", "MONTO_USD",
             "UBICACION"]],
        width="stretch", hide_index=True, height=320,
        column_config={
            "FECHA": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "LOTE": st.column_config.NumberColumn(format="%d"),
            "MONTO": st.column_config.NumberColumn(format="%.2f"),
            "MONTO_USD": st.column_config.NumberColumn("EN $", format="$%.2f"),
        },
    )

    with st.expander("🗑️ Eliminar un gasto"):
        lista = GAS.sort_values("FECHA", ascending=False)
        ops = {}
        for n, r in enumerate(lista.itertuples(), 1):
            f_ = f"{r.FECHA:%d/%m/%Y}" if pd.notna(r.FECHA) else "sin fecha"
            ops[f"#{n} · {f_} · {r.CATEGORIA.lower()} · {r.CONCEPTO or 's/c'} · "
                f"{plata(r.MONTO, r.MONEDA)}"
                + (f" · de {r.UBICACION}" if r.UBICACION else "")] = r.Index
        elegido = st.selectbox("Gasto a eliminar", list(ops.keys()), key="borrar_gasto")
        fila = GAS.loc[ops[elegido]]
        devolver = st.checkbox(f"Devolver la plata a {fila['UBICACION']}",
                               value=bool(fila["UBICACION"]),
                               disabled=not fila["UBICACION"], key="dev_gasto")
        conf = st.checkbox("Sí, quiero borrar este gasto", key="conf_borrar_gasto")
        if st.button("🗑️ Eliminar gasto", disabled=not conf, key="btn_borrar_gasto"):
            ubic_d = fila["UBICACION"] if devolver else None
            if ubic_d and ubic_d not in set(CAJA["UBICACION"]):
                st.error(f"La ubicación {ubic_d} ya no existe en Caja.")
            else:
                delta, _, err = plan_caja(ubic_d, _num(fila["MONTO"]), fila["MONEDA"])
                if err:
                    st.error(err)
                else:
                    df = leer("GASTOS")
                    guardar("GASTOS", df.drop(index=ops[elegido]))
                    extra = aplicar_caja(ubic_d, delta, "Reversión", f"gasto eliminado: {elegido}")
                    avisar("Gasto eliminado." + extra)
                    st.rerun()

    with st.expander("✏️ Editar"):
        edit = st.data_editor(
            gas, num_rows="dynamic", width="stretch", height=360, key="ed_gastos",
            column_config={
                "FECHA": st.column_config.DateColumn(format="DD/MM/YYYY"),
                "CATEGORIA": st.column_config.SelectboxColumn(options=CATEGORIAS_GASTO),
                "LOTE": st.column_config.NumberColumn(format="%d"),
                "MONTO": st.column_config.NumberColumn(format="%.2f"),
                "MONEDA": st.column_config.SelectboxColumn(options=["$", "Bs"]),
            },
        )
        st.caption("Editar aquí no mueve la caja: es para corregir datos.")
        if st.button("💾 Guardar gastos", type="primary"):
            guardar("GASTOS", edit)
            avisar("Gastos actualizados.")
            st.rerun()

# --------------------------------------------------------------------------
# CAJA CHICA
# --------------------------------------------------------------------------
elif seccion == "💵 Caja Chica":
    st.header("Caja Chica")
    total_usd = CAJA["MONTO_USD"].sum()
    k = st.columns(3)
    k[0].metric("Total en dólares", money(total_usd))
    bs = CAJA[CAJA["MONEDA"] == "Bs"]
    k[1].metric("En bolívares", plata(bs["MONTO"].sum(), "Bs"),
                f"≈ {money(bs['MONTO_USD'].sum())}" if len(bs) and TASA else None)
    k[2].metric("Ubicaciones", len(CAJA))
    if CAJA["MONTO_USD"].isna().any():
        st.warning("Hay ubicaciones en bolívares que no entran al total: falta la tasa del "
                   "día en la barra lateral.", icon="💱")

    t_mov, t_cuadre, t_bit = st.tabs(["Movimiento", "🔍 Cuadre", "📒 Bitácora"])

    # ---------------- movimiento y transferencia ----------------
    with t_mov:
        ubis = sorted(x for x in CAJA["UBICACION"].unique() if x)
        c = st.columns([2, 2])
        destino = c[0].selectbox("Ubicación", ubis + ["➕ Nueva ubicación"], key="caja_ubi")
        if destino == "➕ Nueva ubicación":
            nombre = c[1].text_input("Nombre de la ubicación", key="caja_nueva").strip().upper()
            mon_u = c[1].selectbox("Moneda de la ubicación", ["$", "Bs"], key="caja_nueva_mon")
            actual = 0.0
        else:
            nombre = destino
            fila = CAJA[CAJA["UBICACION"] == nombre]
            mon_u = fila["MONEDA"].iloc[0]
            actual = _num(fila["MONTO"].sum())
            eq = (f" (≈ {money(fila['MONTO_USD'].sum())})"
                  if mon_u == "Bs" and pd.notna(fila["MONTO_USD"].sum()) else "")
            c[1].caption(f"Tiene {plata(actual, mon_u)}{eq}.")

        with st.form("mov_caja", clear_on_submit=True):
            c = st.columns([1, 1, 2])
            signo = c[0].radio("Tipo", ["Entrada", "Salida"], horizontal=True)
            monto_txt = campo_dinero(c[1], f"Monto ({mon_u})")
            concepto = c[2].text_input("Concepto (opcional)")
            if st.form_submit_button("Aplicar", type="primary"):
                monto = leer_monto(monto_txt)
                if not nombre:
                    st.error("Falta la ubicación.")
                elif monto is None:
                    st.error("Revisa el monto: escribe por ejemplo 12 o 12,55.")
                elif monto <= 0:
                    st.error("El monto debe ser mayor que cero.")
                else:
                    cambio = monto if signo == "Entrada" else -monto
                    nuevo_saldo = actual + cambio
                    if nuevo_saldo < -0.004:
                        st.error(f"No alcanza: {nombre} tiene {plata(actual, mon_u)} y quieres "
                                 f"sacar {plata(monto, mon_u)}.")
                    else:
                        df = leer("CAJA")
                        df["MONTO"] = pd.to_numeric(df["MONTO"], errors="coerce").fillna(0) \
                            .astype(float)
                        if nombre in set(df["UBICACION"]):
                            df.loc[df["UBICACION"] == nombre, "MONTO"] = nuevo_saldo
                        else:
                            df = pd.concat([df, pd.DataFrame([{
                                "UBICACION": nombre, "MONTO": nuevo_saldo, "MONEDA": mon_u,
                                "TASA": None}])], ignore_index=True)
                        guardar("CAJA", df)
                        registrar_mov(nombre, cambio, mon_u, nuevo_saldo, "Manual",
                                      concepto or signo.lower())
                        st.session_state["ult_caja"] = nombre
                        avisar(f"{signo} de {plata(monto, mon_u)} en {nombre}"
                               + (f" ({concepto})" if concepto else "")
                               + f". Queda {plata(nuevo_saldo, mon_u)}.")
                        st.rerun()

        st.subheader("Transferir entre ubicaciones")
        if len(ubis) >= 2:
            with st.form("transferir", clear_on_submit=True):
                c = st.columns([2, 2, 1])
                de = c[0].selectbox("Desde", ubis, key="tr_de")
                a = c[1].selectbox("Hacia", ubis, index=1, key="tr_a")
                monto_txt = campo_dinero(c[2], "Monto (moneda de origen)")
                if st.form_submit_button("Transferir", type="primary"):
                    monto = leer_monto(monto_txt)
                    mon_de = CAJA.loc[CAJA["UBICACION"] == de, "MONEDA"].iloc[0]
                    d_de, _, err1 = plan_caja(de, -(monto or 0), mon_de)
                    d_a, mon_a, err2 = plan_caja(a, monto or 0, mon_de)
                    if de == a:
                        st.error("Elige dos ubicaciones distintas.")
                    elif monto is None or monto <= 0:
                        st.error("Revisa el monto.")
                    elif err1 or err2:
                        st.error(err1 or err2)
                    else:
                        df = leer("CAJA")
                        df["MONTO"] = pd.to_numeric(df["MONTO"], errors="coerce").fillna(0) \
                            .astype(float)
                        df.loc[df["UBICACION"] == de, "MONTO"] += d_de
                        df.loc[df["UBICACION"] == a, "MONTO"] += d_a
                        guardar("CAJA", df)
                        s_de = _num(df.loc[df["UBICACION"] == de, "MONTO"].sum())
                        s_a = _num(df.loc[df["UBICACION"] == a, "MONTO"].sum())
                        registrar_mov(de, d_de, mon_de, s_de, "Transferencia", f"hacia {a}")
                        registrar_mov(a, d_a, mon_a, s_a, "Transferencia", f"desde {de}")
                        avisar(f"Transferidos {plata(monto, mon_de)} de {de} a {a}"
                               + (f" (llegan {plata(d_a, mon_a)})" if mon_a != mon_de else "")
                               + ".")
                        st.rerun()
        else:
            st.caption("Necesitas al menos dos ubicaciones.")

        with st.expander("✏️ Editar saldos directamente"):
            st.caption("Para corregir un saldo a mano, cambiar la moneda o eliminar una "
                       "ubicación. Las diferencias quedan anotadas en la bitácora como "
                       "«Edición manual».")
            edit = st.data_editor(caja, num_rows="dynamic", width="stretch", key="ed_caja",
                                  column_config={
                                      "MONTO": st.column_config.NumberColumn(format="%.2f"),
                                      "MONEDA": st.column_config.SelectboxColumn(
                                          options=["$", "Bs"]),
                                      "TASA": st.column_config.NumberColumn(
                                          format="%.2f",
                                          help="Opcional. Se usa solo si no hay tasa del día.")})
            if st.button("💾 Guardar caja", type="primary"):
                antes = dict(zip(caja["UBICACION"], caja["MONTO"].fillna(0)))
                guardar("CAJA", edit)
                despues_df = leer("CAJA")
                despues = dict(zip(despues_df["UBICACION"], despues_df["MONTO"].fillna(0)))
                monedas_d = dict(zip(despues_df["UBICACION"], moneda_norm(despues_df["MONEDA"])))
                for u in sorted(set(antes) | set(despues)):
                    dif = _num(despues.get(u, 0)) - _num(antes.get(u, 0))
                    if abs(dif) > 0.004:
                        registrar_mov(u, dif, monedas_d.get(u, "$"), _num(despues.get(u, 0)),
                                      "Edición manual",
                                      "ubicación eliminada" if u not in despues else "")
                avisar("Caja actualizada.")
                st.rerun()

    # ---------------- cuadre ----------------
    with t_cuadre:
        st.caption("Cuenta la plata real de cada ubicación y escríbela. La app te muestra la "
                   "diferencia contra lo que debería haber. Deja vacío lo que no contaste.")
        movs = leer("MOVIMIENTOS")
        ult_cuadre = (movs[movs["ORIGEN"] == "Cuadre"].groupby("UBICACION")["FECHA"].max()
                      if len(movs) else pd.Series(dtype="datetime64[ns]"))
        with st.form("cuadre", clear_on_submit=False):
            contados = {}
            for r in CAJA.sort_values("UBICACION").itertuples():
                c = st.columns([2, 1, 1])
                ult = ult_cuadre.get(r.UBICACION)
                c[0].markdown(f"**{r.UBICACION}**  \nla app dice {plata(r.MONTO, r.MONEDA)}"
                              + (f"  \núltimo cuadre {ult:%d/%m/%Y}" if pd.notna(ult) else ""))
                contados[r.UBICACION] = (campo_dinero(c[1], f"Contado ({r.MONEDA})",
                                                      key=f"cuadre_{r.UBICACION}"),
                                         r.MONEDA, _num(r.MONTO))
            ajustar = st.checkbox("Ajustar los saldos a lo contado", value=True,
                                  help="Si lo desmarcas, solo ves las diferencias.")
            if st.form_submit_button("Revisar cuadre", type="primary"):
                filas, malos = [], []
                for u, (txt_c, mon_c, app_v) in contados.items():
                    if not str(txt_c).strip():
                        continue
                    contado = leer_monto(txt_c)
                    if contado is None:
                        malos.append(u)
                        continue
                    filas.append({"UBICACION": u, "MONEDA": mon_c, "APP": app_v,
                                  "CONTADO": contado, "DIFERENCIA": contado - app_v})
                if malos:
                    st.error("Revisa lo escrito en: " + ", ".join(malos))
                elif not filas:
                    st.info("No escribiste ningún monto contado.")
                else:
                    res = pd.DataFrame(filas)
                    st.session_state["cuadre_res"] = res
                    cuadradas = res["DIFERENCIA"].abs() <= 0.004
                    if ajustar and (~cuadradas).any():
                        df = leer("CAJA")
                        df["MONTO"] = pd.to_numeric(df["MONTO"], errors="coerce").fillna(0) \
                            .astype(float)
                        for r in res.itertuples():
                            df.loc[df["UBICACION"] == r.UBICACION, "MONTO"] = r.CONTADO
                        guardar("CAJA", df)
                    for r in res.itertuples():
                        registrar_mov(r.UBICACION,
                                      r.DIFERENCIA if ajustar else 0.0, r.MONEDA,
                                      r.CONTADO if ajustar else r.APP, "Cuadre",
                                      f"contado {plata(r.CONTADO, r.MONEDA)}, app "
                                      f"{plata(r.APP, r.MONEDA)}, diferencia "
                                      f"{plata(r.DIFERENCIA, r.MONEDA)}"
                                      + ("" if ajustar else " (sin ajustar)"))
                    n_ok = int(cuadradas.sum())
                    avisar(f"Cuadre registrado: {n_ok} de {len(res)} ubicaciones cuadran."
                           + (" Saldos ajustados a lo contado." if ajustar and n_ok < len(res)
                              else ""))
                    st.rerun()
        if "cuadre_res" in st.session_state:
            res = st.session_state["cuadre_res"]
            st.dataframe(
                res.assign(ESTADO=res["DIFERENCIA"].map(
                    lambda d: "✅ cuadra" if abs(d) <= 0.004
                    else ("⬆️ sobra" if d > 0 else "⬇️ falta"))),
                width="stretch", hide_index=True,
                column_config={c: st.column_config.NumberColumn(format="%.2f")
                               for c in ("APP", "CONTADO", "DIFERENCIA")})
            st.caption("Si falta plata, busca en la bitácora: una venta sin registrar o un "
                       "gasto que no anotaste suelen ser la causa.")

    # ---------------- bitácora ----------------
    with t_bit:
        movs = leer("MOVIMIENTOS")
        if movs.empty:
            st.info("Todavía no hay movimientos. Desde ahora, cada entrada y salida de caja "
                    "queda anotada aquí.")
        else:
            c = st.columns([2, 2, 2])
            f_u = c[0].multiselect("Ubicación", sorted(movs["UBICACION"].unique()))
            f_o = c[1].multiselect("Origen", sorted(movs["ORIGEN"].unique()))
            desde_d = c[2].date_input("Desde", value=hoy() - dt.timedelta(days=30),
                                      format="DD/MM/YYYY")
            m = movs[movs["FECHA"].dt.date >= desde_d]
            if f_u:
                m = m[m["UBICACION"].isin(f_u)]
            if f_o:
                m = m[m["ORIGEN"].isin(f_o)]
            st.dataframe(
                m.sort_values("FECHA", ascending=False), width="stretch", hide_index=True,
                height=420,
                column_config={
                    "FECHA": st.column_config.DatetimeColumn(format="DD/MM/YYYY HH:mm"),
                    "MONTO": st.column_config.NumberColumn(format="%+.2f"),
                    "SALDO": st.column_config.NumberColumn("SALDO DESPUÉS", format="%.2f"),
                })
            st.caption(f"{len(m)} movimientos. MONTO positivo entra, negativo sale, en la "
                       "moneda de la ubicación.")

    con_saldo = CAJA[CAJA["MONTO_USD"] > 0]
    if len(con_saldo):
        st.plotly_chart(
            px.pie(con_saldo, values="MONTO_USD", names="UBICACION", hole=0.45,
                   title="Reparto en dólares"),
            width="stretch",
        )


# --------------------------------------------------------------------------
# HISTORICO (archivo viejo, lotes 1-10)
# --------------------------------------------------------------------------
else:
    st.header("Desde el comienzo")

    ven_h = leer("VENTAS_HIST")
    inv_h = leer("INVENTARIO_HIST")

    if ven_h.empty and inv_h.empty:
        st.info("Todavía no hay histórico cargado. Corre `python importar_v1.py "
                "Negocio_Shein.xlsx <ID>` para traer los lotes 1-10.")
        st.stop()

    stock_h = inv_h[inv_h["TIPO"] == "I"].copy()
    stock_h["INVERSION"] = stock_h["CANTIDAD"].fillna(0) * stock_h["PRECIO_COMPRA"].fillna(0)
    pedidos_h = inv_h[inv_h["TIPO"] == "P"].copy()
    pedidos_h["INVERSION"] = pedidos_h["CANTIDAD"].fillna(0) * pedidos_h["PRECIO_COMPRA"].fillna(0)

    capital_inicial = stock_h[stock_h["LOTE"] == 1]["INVERSION"].sum()
    compras_viejas = stock_h["INVERSION"].sum()
    compras_nuevas = INV[INV["LOTE"] > LOTE_CORTE]["INVERSION"].sum()
    arrastre = INV[INV["LOTE"] <= LOTE_CORTE]["INVERSION"].sum()
    compras = compras_viejas + compras_nuevas
    ingresos = ven_h["MONTO"].sum() + VEN["MONTO"].sum()
    gastos_tot = GAS["MONTO_USD"].sum() if len(GAS) else 0.0
    excedente = ingresos - compras - gastos_tot

    k = st.columns(4)
    k[0].metric("Capital inicial (lote 1)", money(capital_inicial))
    k[1].metric("Compras de mercancía", money(compras),
                f"{len(stock_h) + len(INV[INV['LOTE'] > LOTE_CORTE])} líneas")
    k[2].metric("Ingresos acumulados", money(ingresos), f"{len(ven_h) + len(VEN)} ventas")
    k[3].metric("Excedente generado", money(excedente),
                f"{ingresos / capital_inicial:.0f}x el capital inicial"
                if capital_inicial else None)

    st.caption(
        f"Cubre {ven_h['FECHA'].min():%b %Y} a {VEN['FECHA'].max():%b %Y}. "
        f"Las compras excluyen {money(arrastre)} de los lotes 1-{LOTE_CORTE} que reingresaste "
        f"como inventario inicial (ya contados arriba) y {money(pedidos_h['INVERSION'].sum())} "
        "de pedidos por encargo, que eran plata de terceros."
        + (f" El excedente descuenta {money(gastos_tot)} en gastos registrados en la app; "
           "los envíos del archivo viejo no se incluyen." if gastos_tot else "")
    )

    st.divider()
    st.subheader("Ingresos por mes, historia completa")
    a = ven_h[["FECHA", "MONTO"]].assign(ETAPA="Histórico (lotes 1-10)")
    b = VEN[["FECHA", "MONTO"]].assign(ETAPA="Actual")
    todo = pd.concat([a, b], ignore_index=True)
    todo["MES"] = todo["FECHA"].dt.to_period("M").astype(str)
    mensual = todo.groupby(["MES", "ETAPA"])["MONTO"].sum().reset_index().sort_values("MES")
    fig = px.bar(mensual, x="MES", y="MONTO", color="ETAPA",
                 labels={"MONTO": "$", "MES": "", "ETAPA": ""})
    fig.update_layout(height=360, legend_title="")
    st.plotly_chart(fig, width="stretch")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Invertido en mercancía por lote")
        viejo = stock_h.groupby("LOTE")["INVERSION"].sum().reset_index().assign(ETAPA="Histórico")
        nuevo = (INV[INV["LOTE"] > LOTE_CORTE].groupby("LOTE")["INVERSION"].sum()
                 .reset_index().assign(ETAPA="Actual"))
        lotes = pd.concat([viejo, nuevo], ignore_index=True).sort_values("LOTE")
        lotes["LOTE"] = lotes["LOTE"].astype(int).astype(str)
        fig = px.bar(lotes, x="LOTE", y="INVERSION", color="ETAPA",
                     labels={"INVERSION": "$", "ETAPA": ""})
        fig.update_layout(height=340, legend_title="")
        st.plotly_chart(fig, width="stretch")
    with c2:
        st.subheader("Compras viejas por categoría")
        cat = stock_h.groupby("CATEGORIA")["INVERSION"].sum().reset_index()
        cat = cat[cat["INVERSION"] > 0]
        st.plotly_chart(px.pie(cat, values="INVERSION", names="CATEGORIA", hole=0.45),
                        width="stretch")

    st.divider()
    st.subheader("Archivo del período viejo")
    st.caption("Solo lectura. En el archivo viejo la existencia se llevaba a mano, así que "
               "estas filas no cruzan con las ventas ni afectan el stock actual.")

    t1, t2 = st.tabs([f"Ventas ({len(ven_h)})", f"Compras ({len(inv_h)})"])
    with t1:
        st.dataframe(
            ven_h.sort_values("FECHA", ascending=False), width="stretch", hide_index=True,
            height=380,
            column_config={"FECHA": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
                           "MONTO": st.column_config.NumberColumn("Monto", format="$%.2f"),
                           "LOTE": st.column_config.NumberColumn("Lote", format="%d")},
        )
    with t2:
        c = st.columns(3)
        f_tipo = c[0].selectbox("Tipo", ["Todos", "I (stock)", "P (pedido)"])
        f_prop = c[1].multiselect("Propietario",
                                  sorted(x for x in inv_h["PROPIETARIO"].unique() if x))
        f_cat = c[2].multiselect("Categoría",
                                 sorted(x for x in inv_h["CATEGORIA"].unique() if x))
        h = inv_h.copy()
        if f_tipo != "Todos":
            h = h[h["TIPO"] == f_tipo[0]]
        if f_prop:
            h = h[h["PROPIETARIO"].isin(f_prop)]
        if f_cat:
            h = h[h["CATEGORIA"].isin(f_cat)]
        h["INVERSION"] = h["CANTIDAD"].fillna(0) * h["PRECIO_COMPRA"].fillna(0)
        st.caption(f"{len(h)} líneas · {money(h['INVERSION'].sum())} en compras")
        st.dataframe(
            h.sort_values(["LOTE", "NOMBRE"]), width="stretch", hide_index=True, height=340,
            column_config={"PRECIO_COMPRA": st.column_config.NumberColumn(format="$%.2f"),
                           "PRECIO_VENTA": st.column_config.NumberColumn(format="$%.2f"),
                           "INVERSION": st.column_config.NumberColumn(format="$%.2f"),
                           "LOTE": st.column_config.NumberColumn(format="%d"),
                           "FECHA_LOTE": st.column_config.DateColumn(format="DD/MM/YYYY")},
        )
