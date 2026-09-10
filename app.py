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
    "ENCARGOS": ["PERSONA", "TIPO", "DESCRIPCION", "COSTO", "COBRAR", "ABONADO", "FECHA",
                 "NOTA"],
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
FECHAS = {"VENTAS": ["FECHA"], "ENCARGOS": ["FECHA"], "VENTAS_HIST": ["FECHA"],
          "INVENTARIO_HIST": ["FECHA_LOTE"]}
CLAVES = {"INVENTARIO": "NOMBRE", "VENTAS": "NOMBRE", "DEUDAS": "DEUDOR", "CAJA": "UBICACION",
          "ENCARGOS": "PERSONA", "VENTAS_HIST": "NOMBRE", "INVENTARIO_HIST": "NOMBRE"}
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


def encargos_calculados(enc):
    enc = enc.copy()
    if enc.empty:
        for c in ("GANANCIA", "COSTO_CONOCIDO", "ESTA_PAGADO", "SALDO", "GAN_COBRADA"):
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
    return enc


def money(x):
    return f"${x:,.2f}"


# --------------------------------------------------------------------------
# Carga
# --------------------------------------------------------------------------
inv_raw = leer("INVENTARIO")
ven_raw = leer("VENTAS")
enc = leer("ENCARGOS")
caja = leer("CAJA")

INV = inventario_calculado(inv_raw, ven_raw)
VEN = ventas_calculadas(ven_raw, inv_raw)
ENC = encargos_calculados(enc)
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
        rango = st.date_input(
            "Rango de fechas",
            value=(base["FECHA"].min().date(), base["FECHA"].max().date()),
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

    caja_hoy = caja["MONTO"].sum()
    por_cobrar = PEND["SALDO"].sum() if len(PEND) else 0.0
    con_costo = PEND[PEND["COSTO_CONOCIDO"]] if len(PEND) else PEND
    gan_encargos = (con_costo["GANANCIA"].sum()
                    - con_costo["GAN_COBRADA"].fillna(0).sum()) if len(con_costo) else 0.0
    recuperacion = por_cobrar - gan_encargos
    sin_costo_enc = int((~PEND["COSTO_CONOCIDO"]).sum()) if len(PEND) else 0
    monto_sin_costo = PEND[~PEND["COSTO_CONOCIDO"]]["SALDO"].sum() if len(PEND) else 0.0

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
    if sin_costo_enc:
        nota += (f" Ojo: {sin_costo_enc} encargos ({money(monto_sin_costo)}) no tienen el costo "
                 "anotado, así que esa parte no se puede repartir.")
    st.caption(nota)

    st.divider()
    st.subheader("Resultado acumulado de la etapa")
    gan_enc_cobrada = ENC["GAN_COBRADA"].fillna(0).sum() if len(ENC) else 0.0
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
                monto = c[1].number_input("Monto $", min_value=0.0, step=1.0, value=5.0)
                fecha = c[2].date_input("Fecha", value=dt.date.today(), format="DD/MM/YYYY")
                lote = c[3].number_input("Lote", min_value=0, step=1, value=0)
                if st.form_submit_button("Guardar venta", type="primary"):
                    if not nombre:
                        st.error("Falta el nombre del producto.")
                    else:
                        agregar("VENTAS", {"NOMBRE": nombre, "MONTO": monto,
                                           "FECHA": pd.Timestamp(fecha), "LOTE": lote or None})
                        st.success(f"Venta registrada: {nombre}")
                        st.rerun()
        elif disponibles.empty:
            st.warning("No hay productos con existencia. Carga inventario o usa producto libre.")
        else:
            etiquetas = {
                f"{r.NOMBRE}  ·  stock {int(r.EXISTENCIA)}  ·  {money(r.PRECIO_VENTA)}": r.NOMBRE
                for r in disponibles.itertuples()
            }
            sel = st.selectbox("Producto", list(etiquetas.keys()))
            nombre = etiquetas[sel]
            prod = disponibles[disponibles["NOMBRE"] == nombre].iloc[0]
            with st.form("venta", clear_on_submit=False):
                c = st.columns([1, 1, 1, 1])
                monto = c[0].number_input("Monto $", min_value=0.0, step=1.0,
                                          value=float(prod["PRECIO_VENTA"] or 0))
                fecha = c[1].date_input("Fecha", value=dt.date.today(), format="DD/MM/YYYY")
                lote = c[2].number_input("Lote", min_value=0, step=1,
                                         value=int(prod["LOTE"] or 0))
                cantidad = c[3].number_input("Unidades", min_value=1, step=1, value=1)
                if st.form_submit_button("Guardar venta", type="primary"):
                    df = leer("VENTAS")
                    nuevas = pd.DataFrame([{
                        "NOMBRE": nombre, "MONTO": monto,
                        "FECHA": pd.Timestamp(fecha), "LOTE": lote or None,
                    }] * int(cantidad))
                    guardar("VENTAS", pd.concat([df, nuevas], ignore_index=True))
                    st.success(f"{int(cantidad)} venta(s) de {nombre} · {money(monto*cantidad)}")
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
            st.success("Ventas actualizadas.")
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
            pc = c[1].number_input("Precio compra $", min_value=0.0, step=0.01, value=1.0)
            pv = c[2].number_input("Precio venta $", min_value=0.0, step=0.5, value=5.0)
            lote = c[3].number_input("Lote", min_value=0, step=1,
                                     value=int(INV["LOTE"].max() or 0))
            if st.form_submit_button("Agregar", type="primary"):
                if not nombre:
                    st.error("El nombre corto es obligatorio.")
                elif nombre in set(inv_raw["NOMBRE"]):
                    st.error("Ese nombre ya existe: usaría el mismo contador de ventas.")
                else:
                    agregar("INVENTARIO", {
                        "DESCRIPCION": desc, "NOMBRE": nombre, "CANTIDAD": cant,
                        "PRECIO_COMPRA": pc, "PRECIO_VENTA": pv, "LOTE": lote,
                    })
                    st.success(f"{nombre} agregado.")
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
            st.success("Inventario actualizado.")
            st.rerun()

# --------------------------------------------------------------------------
# ENCARGOS Y PRESTAMOS
# --------------------------------------------------------------------------
elif seccion == "🎁 Por cobrar":
    st.header("Encargos y préstamos")
    st.caption("Encargo: compras algo para alguien y le cobras más de lo que te costó. "
               "Préstamo: sale plata de SHEIN y vuelve igual. Como pagan por partes, "
               "lo que importa es ABONADO contra COBRAR.")

    pend = PEND
    k = st.columns(4)
    k[0].metric("Saldo por cobrar", money(pend["SALDO"].sum() if len(pend) else 0),
                f"{len(pend)} abiertos")
    k[1].metric("Ya abonado", money(ENC["ABONADO"].sum() if len(ENC) else 0),
                f"de {money(ENC['COBRAR'].sum() if len(ENC) else 0)} en total")
    k[2].metric("Ganancia pendiente",
                money(pend[pend["COSTO_CONOCIDO"]]["GANANCIA"].sum()
                      - pend[pend["COSTO_CONOCIDO"]]["GAN_COBRADA"].fillna(0).sum()
                      if len(pend) else 0))
    k[3].metric("Ganancia ya cobrada",
                money(ENC["GAN_COBRADA"].fillna(0).sum() if len(ENC) else 0))

    if len(pend):
        por_tipo = pend.groupby("TIPO")["SALDO"].agg(["size", "sum"])
        st.caption(" · ".join(f"{tp}: {int(r['size'])} por {money(r['sum'])}"
                              for tp, r in por_tipo.iterrows()))
        faltan = pend[(~pend["COSTO_CONOCIDO"]) | (pend["TIPO"] == "SIN CLASIFICAR")]
        if len(faltan):
            st.warning(
                f"{len(faltan)} filas ({money(faltan['SALDO'].sum())}) están incompletas: "
                "les falta el tipo o el costo. En un préstamo sin interés, COSTO va igual "
                "a COBRAR.", icon="✏️")

    if len(pend):
        st.subheader("Registrar un abono")
        etiquetas = {
            f"{r.PERSONA} · {r.TIPO.lower()}"
            + (f" · {r.DESCRIPCION}" if r.DESCRIPCION else "")
            + f" · debe {money(r.SALDO)}": r.Index
            for r in pend.sort_values("SALDO", ascending=False).itertuples()
        }
        sel = st.selectbox("Cuenta", list(etiquetas.keys()), label_visibility="collapsed")
        idx = etiquetas[sel]
        saldo = float(ENC.loc[idx, "SALDO"])
        with st.form("abono", clear_on_submit=True):
            c = st.columns([2, 1, 2])
            monto = c[0].number_input("Abono $", min_value=0.0, max_value=float(saldo),
                                      step=1.0, value=0.0,
                                      help=f"Debe {money(saldo)}. Marca 'Saldó todo' "
                                           "si te pagó completo.")
            todo = c[1].checkbox("Saldó todo")
            if c[2].form_submit_button("Registrar abono", type="primary"):
                pago = saldo if todo else monto
                if pago <= 0:
                    st.error("El abono debe ser mayor que cero.")
                else:
                    df = leer("ENCARGOS")
                    df.loc[idx, "ABONADO"] = float(df.loc[idx, "ABONADO"] or 0) + pago
                    guardar("ENCARGOS", df)
                    resta = saldo - pago
                    st.success(f"Abono de {money(pago)} registrado. "
                               + ("Cuenta saldada." if resta <= 0.004
                                  else f"Queda debiendo {money(resta)}.")
                               + " Acuérdate de sumarlo en Caja Chica.")
                    st.rerun()

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
            deuda = ENC[(ENC["PERSONA"] == persona) & (~ENC["ESTA_PAGADO"])]["SALDO"].sum()
            if deuda:
                c[1].caption(f"Ya debe {money(deuda)} en {len(ENC[(ENC['PERSONA'] == persona) & (~ENC['ESTA_PAGADO'])])} cuenta(s).")

        if tipo == "Encargo":
            with st.form("nuevo_encargo", clear_on_submit=True):
                desc = st.text_input("Qué pidió")
                c = st.columns(4)
                costo = c[0].number_input("Costo $ (lo que gastaste)", min_value=0.0, step=0.5)
                cobrar = c[1].number_input("Cobrar $", min_value=0.0, step=1.0)
                abonado = c[2].number_input("Ya abonó $", min_value=0.0, step=1.0)
                fecha = c[3].date_input("Fecha", value=dt.date.today(), format="DD/MM/YYYY")
                nota = st.text_input("Nota")
                if st.form_submit_button("Guardar encargo", type="primary"):
                    if not persona:
                        st.error("Falta la persona.")
                    elif cobrar <= 0:
                        st.error("Falta el monto a cobrar.")
                    else:
                        agregar("ENCARGOS", {
                            "PERSONA": persona, "TIPO": "ENCARGO", "DESCRIPCION": desc,
                            "COSTO": costo or None, "COBRAR": cobrar, "ABONADO": abonado,
                            "FECHA": pd.Timestamp(fecha), "NOTA": nota,
                        })
                        st.success(f"Encargo de {persona}: cobras {money(cobrar)}"
                                   + (f", ganas {money(cobrar - costo)}" if costo else ""))
                        st.rerun()
        else:
            with st.form("nuevo_prestamo", clear_on_submit=True):
                c = st.columns(2)
                prestado = c[0].number_input("Monto prestado $", min_value=0.0, step=1.0)
                cobrar = c[1].number_input("A cobrar $", min_value=0.0, step=1.0,
                                           help="Igual al prestado si no cobras interés.")
                c = st.columns([1, 1, 3])
                abonado = c[0].number_input("Ya abonó $", min_value=0.0, step=1.0)
                fecha = c[1].date_input("Fecha", value=dt.date.today(), format="DD/MM/YYYY")
                nota = c[2].text_input("Nota")
                if st.form_submit_button("Guardar préstamo", type="primary"):
                    if not persona:
                        st.error("Falta la persona.")
                    elif prestado <= 0:
                        st.error("Falta el monto prestado.")
                    else:
                        total = cobrar if cobrar > 0 else prestado
                        agregar("ENCARGOS", {
                            "PERSONA": persona, "TIPO": "PRÉSTAMO", "DESCRIPCION": "",
                            "COSTO": prestado, "COBRAR": total, "ABONADO": abonado,
                            "FECHA": pd.Timestamp(fecha), "NOTA": nota,
                        })
                        extra = total - prestado
                        st.success(f"Préstamo a {persona} por {money(prestado)}"
                                   + (f", cobras {money(extra)} de más" if extra else ""))
                        st.rerun()

    if ENC.empty:
        st.info("Todavía no hay nada cargado.")
        st.stop()

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
            e[["PERSONA", "TIPO", "DESCRIPCION", "COSTO", "COBRAR", "ABONADO", "SALDO",
               "GANANCIA", "FECHA", "NOTA"]].sort_values(["SALDO", "PERSONA"],
                                                         ascending=False),
            width="stretch", hide_index=True, height=360,
            column_config={
                "COSTO": st.column_config.NumberColumn(format="$%.2f"),
                "COBRAR": st.column_config.NumberColumn(format="$%.2f"),
                "ABONADO": st.column_config.NumberColumn(format="$%.2f"),
                "SALDO": st.column_config.NumberColumn(format="$%.2f"),
                "GANANCIA": st.column_config.NumberColumn(format="$%.2f"),
                "FECHA": st.column_config.DateColumn(format="DD/MM/YYYY"),
            },
        )
    elif len(e):
        tot = (e.groupby("PERSONA")
               .agg(CUENTAS=("COBRAR", "size"), COBRAR=("COBRAR", "sum"),
                    ABONADO=("ABONADO", "sum"), SALDO=("SALDO", "sum"),
                    GANANCIA=("GANANCIA", "sum"))
               .reset_index().sort_values("SALDO", ascending=False))
        tipos = e.groupby("PERSONA")["TIPO"].agg(lambda s: " + ".join(sorted(set(s))))
        tot["TIPO"] = tot["PERSONA"].map(tipos)
        tot = tot[["PERSONA", "TIPO", "CUENTAS", "COBRAR", "ABONADO", "SALDO", "GANANCIA"]]
        st.dataframe(
            tot, width="stretch", hide_index=True, height=360,
            column_config={
                "COBRAR": st.column_config.NumberColumn(format="$%.2f"),
                "ABONADO": st.column_config.NumberColumn(format="$%.2f"),
                "SALDO": st.column_config.NumberColumn(format="$%.2f"),
                "GANANCIA": st.column_config.NumberColumn(format="$%.2f"),
            },
        )
        st.caption(f"{len(tot)} personas · {len(e)} cuentas · saldo "
                   f"{money(tot['SALDO'].sum())}")

        if len(tot) > 1:
            st.plotly_chart(
                px.bar(tot.sort_values("SALDO"), x="SALDO", y="PERSONA", orientation="h",
                       labels={"SALDO": "$", "PERSONA": ""}),
                width="stretch")
    else:
        st.info("Nada que mostrar con esos filtros.")

    with st.expander("✏️ Editar"):
        edit = st.data_editor(
            enc, num_rows="dynamic", width="stretch", height=400, key="ed_enc",
            column_config={
                "TIPO": st.column_config.SelectboxColumn(options=["ENCARGO", "PRÉSTAMO"]),
                "COSTO": st.column_config.NumberColumn(format="%.2f"),
                "COBRAR": st.column_config.NumberColumn(format="%.2f"),
                "ABONADO": st.column_config.NumberColumn(format="%.2f"),
                "FECHA": st.column_config.DateColumn(format="DD/MM/YYYY"),
            },
        )
        if st.button("💾 Guardar", type="primary"):
            guardar("ENCARGOS", edit)
            st.success("Actualizado.")
            st.rerun()

# --------------------------------------------------------------------------
# CAJA CHICA
# --------------------------------------------------------------------------
elif seccion == "💵 Caja Chica":
    st.header("Caja Chica")
    k = st.columns(2)
    k[0].metric("Total", money(caja["MONTO"].sum()))
    k[1].metric("Ubicaciones", len(caja))

    with st.form("nueva_caja", clear_on_submit=True):
        c = st.columns([2, 1, 1, 1])
        u = c[0].text_input("Ubicación").strip().upper()
        m = c[1].number_input("Monto", step=1.0, value=0.0)
        mo = c[2].selectbox("Moneda", ["$", "Bs"])
        t = c[3].number_input("Tasa", step=0.01, value=0.0)
        if st.form_submit_button("Agregar / actualizar", type="primary"):
            if not u:
                st.error("Falta la ubicación.")
            else:
                df = leer("CAJA")
                if u in set(df["UBICACION"]):
                    df.loc[df["UBICACION"] == u, ["MONTO", "MONEDA", "TASA"]] = [m, mo, t or None]
                else:
                    df = pd.concat([df, pd.DataFrame([{"UBICACION": u, "MONTO": m,
                                                       "MONEDA": mo, "TASA": t or None}])],
                                   ignore_index=True)
                guardar("CAJA", df)
                st.rerun()

    edit = st.data_editor(caja, num_rows="dynamic", width="stretch", key="ed_caja")
    if st.button("💾 Guardar caja", type="primary"):
        guardar("CAJA", edit)
        st.success("Caja actualizada.")
        st.rerun()

    if not caja.empty:
        st.plotly_chart(
            px.pie(caja, values="MONTO", names="UBICACION", hole=0.45),
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
