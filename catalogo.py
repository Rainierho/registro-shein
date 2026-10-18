# -*- coding: utf-8 -*-
"""
Catálogo para clientes: lo disponible, con foto y precio en dólares.
Es una app aparte (Main file: catalogo.py) que lee el mismo Google Sheet que app.py,
pero solo muestra nombre, descripción, foto, precio y disponibilidad. Nunca costos,
deudas ni caja.
"""
from urllib.parse import quote

import pandas as pd
import streamlit as st

SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]


def secreto(clave, default=None):
    try:
        return st.secrets[clave]
    except Exception:
        return default


def _num(v):
    """Número desde valor crudo de Sheets o texto con coma o punto decimal."""
    if isinstance(v, (int, float)):
        return float(v)
    t = "".join(ch for ch in str(v or "") if ch.isdigit() or ch in ".,-")
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".") if t.rfind(",") > t.rfind(".") \
            else t.replace(",", "")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def _matriz(valores):
    if not valores:
        return pd.DataFrame()
    ancho = max(len(f) for f in valores)
    filas = [list(f) + [""] * (ancho - len(f)) for f in valores]
    cab = [str(c).strip().upper() for c in filas[0]]
    return pd.DataFrame(filas[1:], columns=cab)


@st.cache_data(ttl=300, show_spinner="Cargando catálogo...")
def cargar(backend, id_hoja):
    """La caché depende de qué hoja se lee: si cambias los Secrets, no reusa lo viejo."""
    tabs = ("INVENTARIO", "VENTAS", "CONFIG", "LOTES")
    if backend == "gsheets":
        import gspread
        from google.oauth2.service_account import Credentials

        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=SCOPES)
        libro = gspread.authorize(creds).open_by_key(id_hoja)
        datos = {}
        for t in tabs:
            try:
                datos[t] = _matriz(libro.worksheet(t).get_values(
                    value_render_option="UNFORMATTED_VALUE"))
            except gspread.WorksheetNotFound:
                datos[t] = pd.DataFrame()
        return datos
    ruta = secreto("excel_path", "REGISTRO_SHEIN.xlsx")
    datos = {}
    for t in tabs:
        try:
            df = pd.read_excel(ruta, sheet_name=t)
            df.columns = [str(c).strip().upper() for c in df.columns]
            datos[t] = df.fillna("")
        except Exception:
            datos[t] = pd.DataFrame()
    return datos


def wa_numero(tel):
    d = "".join(ch for ch in str(tel or "") if ch.isdigit())
    if d.startswith("58") and len(d) >= 12:
        return d
    if d.startswith("0") and len(d) == 11:
        return "58" + d[1:]
    if len(d) == 10 and d.startswith("4"):
        return "58" + d
    return ""


@st.cache_data(ttl=900, show_spinner=False)
def tasa_usdt():
    """Bs por USDT en Binance P2P: mediana de los 10 primeros anuncios de compra.
    La mediana ignora el anuncio suelto muy por debajo o por encima del mercado."""
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


def ve(x, simbolo):
    """Formato venezolano: $7,00 · Bs 6.160,81"""
    s = f"{x:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"{simbolo}{s}" if simbolo == "$" else f"{simbolo} {s}"


# --- diagnóstico: nunca mostrar "vacío" cuando el problema es de configuración
BACKEND = secreto("backend")
if BACKEND is None:
    st.set_page_config(page_title="Catálogo", page_icon="🛍️")
    st.error("El catálogo no encuentra sus **Secrets**, así que no sabe qué hoja leer.")
    st.markdown("En share.streamlit.io → esta app → **⋮ → Settings → Secrets**, pega los "
                "mismos secrets de tu app principal y guarda. La app se reinicia sola.")
    st.stop()
if BACKEND == "gsheets" and (not secreto("spreadsheet_id")
                             or not secreto("gcp_service_account")):
    st.set_page_config(page_title="Catálogo", page_icon="🛍️")
    faltan = [k for k in ("spreadsheet_id", "gcp_service_account") if not secreto(k)]
    st.error("A los Secrets del catálogo les falta: " + ", ".join(faltan) + ".")
    st.stop()
try:
    datos = cargar(BACKEND, secreto("spreadsheet_id") or secreto("excel_path", ""))
except Exception as e:
    st.set_page_config(page_title="Catálogo", page_icon="🛍️")
    st.error(f"No se pudo leer la hoja ({type(e).__name__}: {e}).")
    st.caption("Revisa que el spreadsheet_id sea el mismo de la app principal y que la "
               "hoja esté compartida con el correo del service account.")
    st.stop()
cfg = datos["CONFIG"]
conf = dict(zip(cfg["CLAVE"].astype(str).str.upper(), cfg["VALOR"])) if len(cfg) else {}
NOMBRE = str(conf.get("CATALOGO_NOMBRE") or "Catálogo")
WHATSAPP = wa_numero(conf.get("CATALOGO_WHATSAPP"))
MOSTRAR_USDT = str(conf.get("CATALOGO_PRECIO_USDT") or "").upper() == "SI"
TASA_USDT = tasa_usdt() if MOSTRAR_USDT else None

st.set_page_config(page_title=NOMBRE, page_icon="🛍️", layout="wide")

inv, ven = datos["INVENTARIO"], datos["VENTAS"]
if inv.empty or "NOMBRE" not in inv.columns:
    st.title(NOMBRE)
    if BACKEND != "gsheets":
        st.error("El catálogo está en modo Excel local y no encontró el archivo. En "
                 "Streamlit Cloud los Secrets deben decir backend = \"gsheets\".")
    else:
        st.info("La pestaña INVENTARIO de tu hoja está vacía o no existe.")
    st.stop()

for c in ("DESCRIPCION", "FOTO", "OCULTAR"):
    if c not in inv.columns:
        inv[c] = ""
inv = inv[inv["NOMBRE"].astype(str).str.strip() != ""].copy()
inv["NOMBRE"] = inv["NOMBRE"].astype(str).str.strip()
inv["CANTIDAD"] = inv["CANTIDAD"].map(_num).fillna(0)
inv["PRECIO"] = inv["PRECIO_VENTA"].map(_num)
vendidas = (ven["NOMBRE"].astype(str).str.strip().value_counts()
            if len(ven) and "NOMBRE" in ven.columns else pd.Series(dtype=int))
inv["QUEDAN"] = (inv["CANTIDAD"] - inv["NOMBRE"].map(vendidas).fillna(0)).astype(int)
lotes = datos["LOTES"]
en_camino = set()
if len(lotes) and {"LOTE", "ESTADO"} <= set(lotes.columns):
    en_camino = {_num(l_) for l_, e in zip(lotes["LOTE"], lotes["ESTADO"])
                 if str(e).strip().upper() == "EN CAMINO" and _num(l_) is not None}
inv["LOTE_N"] = inv["LOTE"].map(_num) if "LOTE" in inv.columns else None
disp = inv[(inv["QUEDAN"] > 0) & (inv["PRECIO"].fillna(0) > 0)
           & (~inv["LOTE_N"].isin(en_camino))
           & (inv["OCULTAR"].astype(str).str.strip().str.upper() != "SI")].copy()
disp["TIPO"] = disp["NOMBRE"].str.split().str[0].str.capitalize()

st.title(f"🛍️ {NOMBRE}")
if WHATSAPP:
    visible = f"0{WHATSAPP[2:5]}-{WHATSAPP[5:]}"          # 584141234567 -> 0414-1234567
    hola = quote("¡Hola! Vi tu catálogo y quiero hacer un pedido.")
    st.markdown(f"📲 **Escríbenos por WhatsApp: "
                f"[{visible}](https://wa.me/{WHATSAPP}?text={hola})**")
st.caption((f"{len(disp)} productos disponibles" if len(disp) != 1
            else "1 producto disponible")
           + (" · bolívares a tasa USDT" if TASA_USDT else ""))

if disp.empty:
    st.info("Por ahora no hay productos disponibles. ¡Vuelve pronto!")
    st.stop()

c = st.columns([3, 2, 2])
busq = c[0].text_input("Buscar", placeholder="anillo, collar, reloj…").strip().lower()
tipos = c[1].multiselect("Tipo", sorted(disp["TIPO"].unique()))
orden = c[2].selectbox("Ordenar", ["Recomendados", "Menor precio", "Mayor precio", "Nombre"])

v = disp
if busq:
    v = v[v["NOMBRE"].str.lower().str.contains(busq, regex=False)
          | v["DESCRIPCION"].astype(str).str.lower().str.contains(busq, regex=False)]
if tipos:
    v = v[v["TIPO"].isin(tipos)]
if orden == "Menor precio":
    v = v.sort_values("PRECIO")
elif orden == "Mayor precio":
    v = v.sort_values("PRECIO", ascending=False)
elif orden == "Nombre":
    v = v.sort_values("NOMBRE")
else:   # primero los que tienen foto, luego los que quedan pocos
    v = v.assign(_f=~v["FOTO"].astype(str).str.startswith("http")).sort_values(["_f", "QUEDAN"])

if v.empty:
    st.info("No hay productos con esa búsqueda.")
    st.stop()

por_fila = 3
filas = [v.iloc[i:i + por_fila] for i in range(0, len(v), por_fila)]
for grupo in filas:
    cols = st.columns(por_fila)
    for col, (_, p) in zip(cols, grupo.iterrows()):
        with col.container(border=True):
            foto = str(p["FOTO"]).strip()
            if foto.startswith("http"):
                st.image(foto, width="stretch")
            st.markdown(f"**{p['NOMBRE'].title()}**")
            desc = str(p["DESCRIPCION"]).strip()
            if desc:
                st.caption(desc[:140] + ("…" if len(desc) > 140 else ""))
            precio = ve(p["PRECIO"], "$")
            if TASA_USDT:
                precio += f"  ·  {ve(p['PRECIO'] * TASA_USDT, 'Bs')}"
            st.markdown(f"### {precio}" if not TASA_USDT else f"#### {precio}")
            st.caption("🔥 ¡Última unidad!" if p["QUEDAN"] == 1
                       else ("Quedan pocas" if p["QUEDAN"] <= 2 else "Disponible"))
            if WHATSAPP:
                ref = ve(p["PRECIO"], "$") + (f" / {ve(p['PRECIO'] * TASA_USDT, 'Bs')}"
                                               if TASA_USDT else "")
                msg = f"¡Hola! Me interesa *{p['NOMBRE'].title()}* ({ref}) que vi en el catálogo."
                st.link_button("📲 Lo quiero", f"https://wa.me/{WHATSAPP}?text={quote(msg)}",
                               width="stretch")

st.divider()
st.caption("Las existencias se actualizan cada pocos minutos. Escríbenos para confirmar.")
