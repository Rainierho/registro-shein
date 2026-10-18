# -*- coding: utf-8 -*-
"""
Registro SHEIN — app Streamlit
Backends: Google Sheets (produccion en Streamlit Cloud) o Excel local (desarrollo).
"""
import datetime as dt
import io
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
    "INVENTARIO": ["DESCRIPCION", "NOMBRE", "CANTIDAD", "PRECIO_COMPRA", "PRECIO_VENTA", "LOTE"],
    "VENTAS": ["NOMBRE", "MONTO", "FECHA", "LOTE"],
    "DEUDAS": ["DEUDOR", "MONTO", "BOLOS", "NOTA"],
    "CAJA": ["UBICACION", "MONTO", "MONEDA", "TASA"],
    "ENCARGOS": ["PERSONA", "TIPO", "DESCRIPCION", "COSTO", "COBRAR", "ABONADO", "MONEDA",
                 "FECHA", "ULT_ABONO", "NOTA"],
    "CONFIG": ["CLAVE", "VALOR"],
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
    "VENTAS_HIST": ["MONTO", "LOTE"],
    "INVENTARIO_HIST": ["CANTIDAD", "PRECIO_COMPRA", "PRECIO_VENTA", "LOTE"],
}
# columnas de plata: siempre float, para que el editor acepte decimales
DINERO = {"MONTO", "PRECIO_COMPRA", "PRECIO_VENTA", "BOLOS", "TASA", "COSTO", "COBRAR",
          "ABONADO"}
FECHAS = {"VENTAS": ["FECHA"], "ENCARGOS": ["FECHA", "ULT_ABONO"], "VENTAS_HIST": ["FECHA"],
          "INVENTARIO_HIST": ["FECHA_LOTE"]}
CLAVES = {"INVENTARIO": "NOMBRE", "VENTAS": "NOMBRE", "DEUDAS": "DEUDOR", "CAJA": "UBICACION",
          "ENCARGOS": "PERSONA", "CONFIG": "CLAVE", "VENTAS_HIST": "NOMBRE",
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
        if isinstance(v, (int, float)) and not pd.isna(v):
            return EPOCA_SHEETS + dt.timedelta(days=float(v))
        return v

    return pd.to_datetime(serie.map(conv), dayfirst=True, errors="coerce")


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
                linea.append(pd.Timestamp(v).strftime("%Y-%m-%d"))
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


def agregar(tab, fila: dict):
    df = leer(tab)
    df = pd.concat([df, pd.DataFrame([fila])], ignore_index=True)
    guardar(tab, df)


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
    return "—" if x is None else f"${x:,.2f}"


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
    return f"${x:,.2f}" if moneda != "Bs" else f"Bs {x:,.2f}"


def avisar(texto):
    """Mensaje que se muestra después del rerun (si no, la recarga lo borra)."""
    st.session_state["aviso"] = texto


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


def aplicar_caja(ubic, delta):
    """Mueve el saldo de una ubicación. Devuelve texto para el mensaje de éxito."""
    if not ubic or abs(delta) < 0.0001:
        return ""
    df = leer("CAJA")
    df["MONTO"] = pd.to_numeric(df["MONTO"], errors="coerce").fillna(0).astype(float)
    mascara = df["UBICACION"] == ubic
    mon_u = moneda_norm(df.loc[mascara, "MONEDA"]).iloc[0] if mascara.any() else "$"
    df.loc[mascara, "MONTO"] = df.loc[mascara, "MONTO"] + delta
    guardar("CAJA", df)
    st.session_state["ult_caja"] = ubic
    verbo = "Entraron" if delta > 0 else "Salieron"
    return f" {verbo} {plata(abs(delta), mon_u)} {'a' if delta > 0 else 'de'} {ubic}."


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

INV = inventario_calculado(inv_raw, ven_raw)
VEN = ventas_calculadas(ven_raw, inv_raw)
ENC = encargos_calculados(enc, TASA)
CAJA = caja_calculada(caja, TASA)
PEND = ENC[~ENC["ESTA_PAGADO"]]

# --------------------------------------------------------------------------
# Sidebar
# --------------------------------------------------------------------------
with st.sidebar:
    st.title("💍 Registro SHEIN")
    seccion = st.radio(
        "Secciones",
        ["📊 Dashboard", "🧾 Ventas", "📦 Inventario", "🎁 Por cobrar", "💵 Caja Chica",
         "🏛️ Histórico"],
        label_visibility="collapsed",
    )
    st.divider()
    st.caption(f"Backend: **{BACKEND}**")
    if FALLA_SECRETS:
        st.error(f"secrets.toml no se pudo leer ({FALLA_SECRETS}). "
                 "Revísalo: si lo generaste en PowerShell, guárdalo en UTF-8 sin BOM.", icon="🔑")
    if BACKEND != "gsheets":
        st.warning("Modo Excel local: en Streamlit Cloud los cambios NO se guardan.", icon="⚠️")
    st.caption("Tasa del día (Bs por $)" + (f" · {TASA_FECHA}" if TASA_FECHA else ""))
    with st.form("form_tasa", clear_on_submit=False, border=False):
        tasa_txt = st.text_input("Tasa", value=fmt_monto(TASA), placeholder="Ej. 312,45",
                                 label_visibility="collapsed")
        if st.form_submit_button("Guardar tasa", width="stretch"):
            nueva = leer_monto(tasa_txt)
            if not nueva or nueva <= 0:
                st.error("Tasa inválida.")
            else:
                guardar("CONFIG", pd.DataFrame([
                    {"CLAVE": "TASA", "VALOR": f"{nueva}"},
                    {"CLAVE": "TASA_FECHA", "VALOR": f"{dt.date.today():%d/%m/%Y}"}]))
                st.rerun()
    if not TASA and (CAJA["MONEDA"].eq("Bs").any() or ENC["MONEDA"].eq("Bs").any()):
        st.warning("Hay montos en bolívares y no hay tasa: no se suman a los totales.",
                   icon="💱")
    st.divider()
    if st.button("🔄 Refrescar datos", width="stretch"):
        st.cache_data.clear()
        st.rerun()

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as w:
        INV.to_excel(w, sheet_name="INVENTARIO", index=False)
        VEN.to_excel(w, sheet_name="VENTAS", index=False)
        ENC.to_excel(w, sheet_name="ENCARGOS", index=False)
        caja.to_excel(w, sheet_name="CAJA", index=False)
    st.download_button(
        "⬇️ Descargar Excel",
        buffer.getvalue(),
        file_name=f"REGISTRO_SHEIN_{dt.date.today():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )

if "aviso" in st.session_state:
    st.success(st.session_state.pop("aviso"), icon="✅")

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
        f_min, f_max = base["FECHA"].min(), base["FECHA"].max()
        hoy = dt.date.today()
        rango = st.date_input(
            "Rango de fechas",
            value=(f_min.date() if pd.notna(f_min) else hoy,
                   f_max.date() if pd.notna(f_max) else hoy),
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

    k = st.columns(4)
    k[0].metric("Ingresos", money(ingresos), f"{len(v)} ventas")
    k[1].metric("Costo de lo vendido" + solo_actual, money(costo))
    k[2].metric("Ganancia neta" + solo_actual, money(ganancia), f"{margen:.1f}% margen")
    k[3].metric("Ticket promedio", money(v["MONTO"].mean() if len(v) else 0))

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
        lote = inv_lote.merge(ven_lote, on="LOTE", how="outer").fillna(0)
        lote = lote[lote["LOTE"].notna()].sort_values("LOTE")
        lote["LOTE"] = lote["LOTE"].astype(int).astype(str)
        fig = px.bar(
            lote, x="LOTE", y=["INVERTIDO", "RECUPERADO"], barmode="group",
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
    k = st.columns(3)
    k[0].metric("Ganancia de reventa", money(realizada))
    k[1].metric("Ganancia de encargos cobrada", money(gan_enc_cobrada))
    k[2].metric("Total ganado", money(realizada + gan_enc_cobrada))

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
                fecha = c[2].date_input("Fecha", value=dt.date.today(), format="DD/MM/YYYY")
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
                        extra = aplicar_caja(ubic, delta)
                        avisar(f"Venta registrada: {nombre} · {money(monto)}.{extra}")
                        st.rerun()
        elif disponibles.empty:
            st.warning("No hay productos con existencia. Carga inventario o usa producto libre.")
        else:
            etiquetas = {
                f"{r.NOMBRE}  ·  stock {_ent(r.EXISTENCIA)}  ·  {money(r.PRECIO_VENTA)}": r.NOMBRE
                for r in disponibles.itertuples()
            }
            sel = st.selectbox("Producto", list(etiquetas.keys()))
            nombre = etiquetas[sel]
            prod = disponibles[disponibles["NOMBRE"] == nombre].iloc[0]
            with st.form("venta", clear_on_submit=False):
                c = st.columns([1, 1, 1, 1])
                monto_txt = campo_dinero(c[0], "Monto $", prod["PRECIO_VENTA"],
                                         key=f"monto_venta_{nombre}")
                fecha = c[1].date_input("Fecha", value=dt.date.today(), format="DD/MM/YYYY")
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
                    df = leer("VENTAS")
                    nuevas = pd.DataFrame([{
                        "NOMBRE": nombre, "MONTO": monto,
                        "FECHA": pd.Timestamp(fecha), "LOTE": lote or None,
                    }] * int(cantidad))
                    guardar("VENTAS", pd.concat([df, nuevas], ignore_index=True))
                    extra = aplicar_caja(ubic, delta)
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
                    extra = aplicar_caja(ubic, delta)
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
                    })
                    extra = aplicar_caja(ubic, delta)
                    avisar(f"{nombre} agregado.{extra}")
                    st.rerun()

    c1, c2, c3 = st.columns([2, 1, 1])
    busq = c1.text_input("Buscar").strip().upper()
    f_lote = c2.multiselect("Lote", sorted(int(x) for x in INV["LOTE"].dropna().unique()))
    estado = c3.selectbox("Estado", ["Todos", "Con existencia", "Agotados"])

    i = INV.copy()
    if busq:
        i = i[i["NOMBRE"].str.contains(busq, na=False) | i["DESCRIPCION"].str.upper().str.contains(busq, na=False)]
    if f_lote:
        i = i[i["LOTE"].isin(f_lote)]
    if estado == "Con existencia":
        i = i[i["EXISTENCIA"] > 0]
    elif estado == "Agotados":
        i = i[i["EXISTENCIA"] <= 0]

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

    with st.expander("✏️ Editar inventario"):
        st.caption("EXISTENCIA y ganancias se calculan solas contando VENTAS por NOMBRE.")
        edit = st.data_editor(
            inv_raw, num_rows="dynamic", width="stretch", height=400,
            key="ed_inv",
            column_config={
                "CANTIDAD": st.column_config.NumberColumn(format="%d"),
                "PRECIO_COMPRA": st.column_config.NumberColumn(format="%.2f"),
                "PRECIO_VENTA": st.column_config.NumberColumn(format="%.2f"),
                "LOTE": st.column_config.NumberColumn(format="%d"),
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
        hoy = pd.Timestamp(dt.date.today())
        per = (pend.groupby("PERSONA")
               .agg(SALDO=("SALDO_USD", "sum"), CUENTAS=("SALDO_USD", "size"),
                    ABIERTA_DESDE=("FECHA", "min"))
               .reset_index())
        # último abono de la persona en cualquiera de sus cuentas, saldadas o no:
        # una cuenta nueva no reinicia el reloj, un pago sí
        ult = ENC.groupby("PERSONA")["ULT_ABONO"].max()
        per["ULTIMO_ABONO"] = per["PERSONA"].map(ult)
        desde = per["ULTIMO_ABONO"].where(per["ULTIMO_ABONO"].notna(), per["ABIERTA_DESDE"])
        per["DIAS"] = (hoy - pd.to_datetime(desde)).dt.days
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
        st.dataframe(
            per[["ESTADO", "PERSONA", "SALDO", "DIAS", "ULTIMO_ABONO", "CUENTAS"]],
            width="stretch", hide_index=True, height=min(38 * (len(per) + 1), 330),
            column_config={
                "SALDO": st.column_config.NumberColumn("Debe ($)", format="$%.2f"),
                "DIAS": st.column_config.NumberColumn("Días sin abonar", format="%d"),
                "ULTIMO_ABONO": st.column_config.DateColumn("Último abono",
                                                            format="DD/MM/YYYY"),
            },
        )
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
                df.loc[vacias, "FECHA"] = pd.Timestamp(dt.date.today())
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
            fecha_ab = c[1].date_input("Fecha del abono", value=dt.date.today(),
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
                    extra = aplicar_caja(ubic, delta)
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
                fecha = c[3].date_input("Fecha", value=dt.date.today(), format="DD/MM/YYYY")
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
                        extra = aplicar_caja(ubic, delta)
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
                fecha = c[1].date_input("Fecha", value=dt.date.today(), format="DD/MM/YYYY")
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
                        extra = aplicar_caja(ubic, delta)
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
                extra = aplicar_caja(ubic_b, delta)
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

    st.subheader("Movimiento")
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
                    avisar(f"Transferidos {plata(monto, mon_de)} de {de} a {a}"
                           + (f" (llegan {plata(d_a, mon_a)})" if mon_a != mon_de else "")
                           + ".")
                    st.rerun()
    else:
        st.caption("Necesitas al menos dos ubicaciones.")

    st.divider()
    with st.expander("✏️ Editar saldos directamente"):
        st.caption("Para corregir un saldo a mano, cambiar la moneda o eliminar una ubicación.")
        edit = st.data_editor(caja, num_rows="dynamic", width="stretch", key="ed_caja",
                              column_config={
                                  "MONTO": st.column_config.NumberColumn(format="%.2f"),
                                  "MONEDA": st.column_config.SelectboxColumn(
                                      options=["$", "Bs"]),
                                  "TASA": st.column_config.NumberColumn(
                                      format="%.2f",
                                      help="Opcional. Se usa solo si no hay tasa del día.")})
        if st.button("💾 Guardar caja", type="primary"):
            guardar("CAJA", edit)
            avisar("Caja actualizada.")
            st.rerun()

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
    excedente = ingresos - compras

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
