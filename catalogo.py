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


FOTO_PROPIA = "foto_propia"      # formato viejo: una foto subida ligada al nombre


def fotos_de(valor):
    """Celda FOTO -> fotos en orden: enlaces, «propia:<id>» o el formato viejo."""
    salida = []
    for parte in str(valor or "").split("|"):
        q = parte.strip()
        if q.startswith("http") or q.startswith("propia:") or q == FOTO_PROPIA:
            salida.append(q)
    return salida


@st.cache_data(ttl=300, show_spinner=False)
def cargar_fotos(backend, id_hoja):
    """Fotos subidas desde la app (pestaña FOTOS): {id o "viejo:NOMBRE": base64}.
    Solo se lee si algún producto disponible las usa."""
    df = pd.DataFrame()
    try:
        if backend == "gsheets":
            import gspread
            from google.oauth2.service_account import Credentials
            creds = Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]), scopes=SCOPES)
            libro = gspread.authorize(creds).open_by_key(id_hoja)
            df = _matriz(libro.worksheet("FOTOS").get_values(value_render_option="UNFORMATTED_VALUE"))
        else:
            df = pd.read_excel(id_hoja, sheet_name="FOTOS", dtype=str).fillna("")
            df.columns = [str(c).strip().upper() for c in df.columns]
    except Exception:
        return {}
    salida = {}
    ids = df["ID"] if "ID" in df.columns else [""] * len(df)
    for n, img, id_ in zip(df.get("NOMBRE", []), df.get("IMG", []), ids):
        id_ = str(id_ or "").strip()
        salida[id_ if id_ else f"viejo:{str(n).strip()}"] = str(img)
    return salida


def fuente_foto(ref, nombre):
    """Una foto -> dirección que entiende el navegador (enlace o data:image/jpeg;base64)."""
    if ref.startswith("http"):
        return ref
    clave = ref[len("propia:"):] if ref.startswith("propia:") else f"viejo:{nombre}"
    b64 = FOTOS_PROPIAS.get(clave)
    return f"data:image/jpeg;base64,{b64}" if b64 else None


def pasarela(fuentes, alto=320):
    """Varias fotos en una tira deslizable: con el dedo en el teléfono, con flechas y puntos
    en la computadora."""
    import html as h
    slides = "".join(f'<div class="s"><img loading="lazy" alt="" src="{h.escape(f, quote=True)}">'
                     f'</div>' for f in fuentes)
    puntos = "".join(f'<span class="d" data-i="{i}"></span>' for i in range(len(fuentes)))
    pagina = f"""
<style>
*{{box-sizing:border-box;margin:0;padding:0}} html,body{{background:transparent;overflow:hidden}}
.w{{position:relative;height:{alto}px;border-radius:12px;overflow:hidden;background:rgba(161,202,220,.18)}}
.t{{display:flex;height:100%;overflow-x:auto;scroll-snap-type:x mandatory;scrollbar-width:none;
   -webkit-overflow-scrolling:touch;scroll-behavior:smooth}}
.t::-webkit-scrollbar{{display:none}}
.s{{flex:0 0 100%;height:100%;scroll-snap-align:center;display:flex;align-items:center;
   justify-content:center}}
.s img{{max-width:100%;max-height:100%;object-fit:contain;user-select:none;-webkit-user-drag:none}}
.b{{position:absolute;top:50%;transform:translateY(-50%);border:0;width:34px;height:34px;
   border-radius:50%;background:rgba(0,0,0,.45);color:#fff;font-size:20px;line-height:34px;
   cursor:pointer}}
.b:hover{{background:rgba(0,0,0,.65)}} .p{{left:8px}} .n{{right:8px}}
.ds{{position:absolute;bottom:10px;width:100%;text-align:center}}
.d{{display:inline-block;width:8px;height:8px;margin:0 3px;border-radius:50%;cursor:pointer;
   background:rgba(255,255,255,.55);box-shadow:0 0 3px rgba(0,0,0,.6)}}
.d.on{{background:#fff;transform:scale(1.25)}}
</style>
<div class="w"><div class="t" id="t">{slides}</div>
<button class="b p" aria-label="anterior" onclick="ir(-1)">&#8249;</button>
<button class="b n" aria-label="siguiente" onclick="ir(1)">&#8250;</button>
<div class="ds">{puntos}</div></div>
<script>
const t=document.getElementById('t'), ds=[...document.querySelectorAll('.d')], n=ds.length;
let cur=0, quieto;
const actual=()=>Math.round(t.scrollLeft/Math.max(t.clientWidth,1));
function pinta(i){{ds.forEach((d,j)=>d.classList.toggle('on',j===i));}}
function a(i){{cur=i; pinta(i); t.scrollTo({{left:i*t.clientWidth}});}}
function ir(k){{a((cur+k+n)%n);}}            // usa la foto destino, no la del medio del giro
t.addEventListener('scroll',()=>{{clearTimeout(quieto);
  quieto=setTimeout(()=>{{cur=actual(); pinta(cur);}},120);}});   // deslizar con el dedo
ds.forEach(d=>d.addEventListener('click',()=>a(+d.dataset.i)));
pinta(0);
</script>"""
    if hasattr(st, "iframe"):              # Streamlit 1.5x+: components.html va a desaparecer
        st.iframe(pagina, height=alto)
    else:
        import streamlit.components.v1 as components
        components.html(pagina, height=alto)


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

st.set_page_config(page_title=NOMBRE, page_icon="🛍️", layout="wide",
                   initial_sidebar_state="collapsed")

# ---------------------------------------------------------------------------
# Presentación: paleta de Lyho Shop, logo y versión para teléfono
# ---------------------------------------------------------------------------
import base64
import html as _html
import pathlib
import re

PALETA = {"teal": "#5A97AC", "celeste": "#A1CADC", "crema": "#EEEDE9", "rosa": "#F7AEC2",
          "vino": "#AE5166", "texto": "#2E3B42"}
RUTA_LOGO = pathlib.Path(__file__).resolve().parent / "assets" / "logo_lyhoshop.png"


@st.cache_data(show_spinner=False)
def logo_b64(ruta):
    try:
        return base64.b64encode(pathlib.Path(ruta).read_bytes()).decode()
    except OSError:
        return None


def es_movil():
    """Teléfono según el navegador. ?vista=movil o ?vista=pc lo fuerzan (para probar)."""
    forzada = str(st.query_params.get("vista", "")).lower()
    if forzada in ("movil", "móvil", "pc"):
        return forzada != "pc"
    try:
        agente = st.context.headers.get("User-Agent", "") or ""
    except Exception:
        agente = ""
    return bool(re.search(r"Mobi|Android|iPhone|iPod", agente))


MOVIL = es_movil()
ALTO_FOTO = 210 if MOVIL else 300

css = f"""
:root{{--teal:{PALETA['teal']};--celeste:{PALETA['celeste']};--crema:{PALETA['crema']};
      --rosa:{PALETA['rosa']};--vino:{PALETA['vino']};--texto:{PALETA['texto']}}}
.stApp{{background:var(--crema)}}
[data-testid="stAppHeader"],[data-testid="stHeader"]{{display:none!important}}
[data-testid="stToolbar"],[data-testid="stDecoration"],#MainMenu,footer{{display:none!important}}
[data-testid="stMainBlockContainer"]{{padding-top:1rem;max-width:1180px}}
.lyho-cab{{background:linear-gradient(135deg,var(--teal),#467F94);border-radius:20px;
  padding:22px 16px 18px;text-align:center;box-shadow:0 6px 18px rgba(46,59,66,.18)}}
.lyho-cab img{{width:min(430px,90%);height:auto;filter:drop-shadow(0 3px 4px rgba(0,0,0,.25))}}
.lyho-cab h1{{color:#fff;margin:0}}
.lyho-wa{{display:inline-block;margin-top:12px;background:var(--rosa);color:var(--texto)!important;
  text-decoration:none!important;font-weight:700;padding:8px 18px;border-radius:999px}}
.lyho-cuenta{{color:var(--teal);font-weight:700;margin:14px 2px 0}}
[class*="st-key-tarjeta_"]{{background:#fff;border:1px solid rgba(174,81,102,.16);border-radius:16px;
  padding:10px 10px 12px;box-shadow:0 3px 10px rgba(46,59,66,.08)}}
[class*="st-key-tarjeta_"] [data-testid="stImage"] img{{height:{ALTO_FOTO}px;width:100%;
  object-fit:contain;background:rgba(161,202,220,.18);border-radius:12px}}
.lyho-nom{{font-weight:700;color:var(--texto);line-height:1.25;margin:4px 0 2px}}
.lyho-desc{{color:#6B777D;font-size:.82rem;line-height:1.3;margin:0 0 4px}}
.lyho-precio{{color:var(--vino);font-weight:800;font-size:1.4rem;line-height:1.2}}
.lyho-bs{{color:var(--teal);font-weight:700;font-size:.85rem}}
.lyho-chip{{display:inline-block;font-size:.72rem;font-weight:700;padding:2px 9px;border-radius:999px;
  background:var(--rosa);color:var(--vino);margin:4px 0 2px}}
.lyho-chip.ok{{background:rgba(161,202,220,.45);color:#3F7184}}
[data-testid="stBaseLinkButton-primary"]{{background:var(--vino)!important;border-color:var(--vino)!important;
  color:#fff!important;border-radius:999px!important;font-weight:700}}
[data-testid="stBaseLinkButton-primary"]:hover{{filter:brightness(.93)}}
.lyho-flot{{position:fixed;right:18px;bottom:18px;z-index:999;background:#25D366;color:#fff!important;
  text-decoration:none!important;font-weight:700;padding:12px 20px;border-radius:999px;
  box-shadow:0 6px 16px rgba(0,0,0,.25)}}
.lyho-pie{{color:#7C878C;font-size:.8rem;text-align:center;margin:18px 0 70px}}
"""
if MOVIL:
    css += """
[data-testid="stMainBlockContainer"]{padding:.55rem .5rem 5.5rem!important}
.lyho-cab{padding:16px 10px 14px;border-radius:16px}
.st-key-grilla [data-testid="stHorizontalBlock"]{flex-wrap:nowrap!important;gap:.5rem!important}
.st-key-grilla [data-testid="stColumn"]{min-width:0!important;width:calc(50% - .25rem)!important;
  flex:1 1 0!important}
[class*="st-key-tarjeta_"]{padding:6px 6px 9px;border-radius:14px}
.lyho-nom{font-size:.86rem}
.lyho-precio{font-size:1.12rem}
.lyho-bs{font-size:.76rem}
[data-testid="stBaseLinkButton-primary"]{font-size:.84rem!important;padding:.3rem .4rem!important;
  min-height:2.1rem!important}
.lyho-flot{left:12px;right:12px;bottom:12px;text-align:center;padding:13px 16px}
"""
st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


def cabecera(total):
    b64 = logo_b64(str(RUTA_LOGO))
    marca = (f'<img src="data:image/png;base64,{b64}" alt="{_html.escape(NOMBRE)}">' if b64
             else f"<h1>{_html.escape(NOMBRE)}</h1>")
    wa = ""
    if WHATSAPP:
        visible = f"0{WHATSAPP[2:5]}-{WHATSAPP[5:]}"            # 584141234567 -> 0414-1234567
        hola = quote("¡Hola! Vi tu catálogo y quiero hacer un pedido.")
        wa = (f'<a class="lyho-wa" href="https://wa.me/{WHATSAPP}?text={hola}" target="_blank">'
              f"📲 Escríbenos: {visible}</a>")
    st.markdown(f'<div class="lyho-cab">{marca}<br>{wa}</div>', unsafe_allow_html=True)
    cuenta = f"{total} productos disponibles" if total != 1 else "1 producto disponible"
    st.markdown(f'<div class="lyho-cuenta">{cuenta}'
                + (" · bolívares a tasa USDT" if TASA_USDT else "") + "</div>",
                unsafe_allow_html=True)


inv, ven = datos["INVENTARIO"], datos["VENTAS"]
if inv.empty or "NOMBRE" not in inv.columns:
    cabecera(0)
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
disp["FOTOS_L"] = disp["FOTO"].map(fotos_de)
FOTOS_PROPIAS = (cargar_fotos(BACKEND, secreto("spreadsheet_id") or secreto("excel_path", ""))
                 if disp["FOTOS_L"].map(lambda l: any(not r.startswith("http") for r in l)).any()
                 else {})

cabecera(len(disp))

if disp.empty:
    st.info("Por ahora no hay productos disponibles. ¡Vuelve pronto!")
    st.stop()

ORDENES = ["Recomendados", "Menor precio", "Mayor precio", "Nombre"]
if MOVIL:
    busq = st.text_input("Buscar", placeholder="🔎 Buscar: anillo, collar, reloj…",
                         label_visibility="collapsed").strip().lower()
    with st.expander("Filtrar y ordenar"):
        tipos = st.multiselect("Tipo", sorted(disp["TIPO"].unique()), placeholder="Todos")
        orden = st.selectbox("Ordenar", ORDENES)
else:
    c = st.columns([3, 2, 2])
    busq = c[0].text_input("Buscar", placeholder="anillo, collar, reloj…").strip().lower()
    tipos = c[1].multiselect("Tipo", sorted(disp["TIPO"].unique()), placeholder="Todos")
    orden = c[2].selectbox("Ordenar", ORDENES)

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
    v = v.assign(_f=v["FOTOS_L"].map(len) == 0).sort_values(["_f", "QUEDAN"])

if v.empty:
    st.info("No hay productos con esa búsqueda.")
    st.stop()


def tarjeta(p, k):
    with st.container(key=f"tarjeta_{k}"):
        fuentes = [f for f in (fuente_foto(r, p["NOMBRE"]) for r in p["FOTOS_L"]) if f]
        if len(fuentes) > 1:
            pasarela(fuentes, alto=ALTO_FOTO)
        elif fuentes:
            f0 = fuentes[0]
            if f0.startswith("data:"):
                f0 = base64.b64decode(f0.split(",", 1)[1])
            st.image(f0, width="stretch")
        nombre = _html.escape(p["NOMBRE"].title())
        desc = str(p["DESCRIPCION"]).strip()
        bloque = f'<div class="lyho-nom">{nombre}</div>'
        if desc and not MOVIL:
            corto = desc[:110] + ("…" if len(desc) > 110 else "")
            bloque += f'<div class="lyho-desc">{_html.escape(corto)}</div>'
        bloque += f'<div class="lyho-precio">{ve(p["PRECIO"], "$")}</div>'
        if TASA_USDT:
            bloque += f'<div class="lyho-bs">{ve(p["PRECIO"] * TASA_USDT, "Bs")}</div>'
        if p["QUEDAN"] == 1:
            bloque += '<span class="lyho-chip">🔥 ¡Última unidad!</span>'
        elif p["QUEDAN"] <= 2:
            bloque += '<span class="lyho-chip">Quedan pocas</span>'
        else:
            bloque += '<span class="lyho-chip ok">Disponible</span>'
        st.markdown(bloque, unsafe_allow_html=True)
        if WHATSAPP:
            ref = ve(p["PRECIO"], "$") + (f" / {ve(p['PRECIO'] * TASA_USDT, 'Bs')}"
                                           if TASA_USDT else "")
            msg = f"¡Hola! Me interesa *{p['NOMBRE'].title()}* ({ref}) que vi en el catálogo."
            st.link_button("📲 Lo quiero", f"https://wa.me/{WHATSAPP}?text={quote(msg)}",
                           type="primary", width="stretch")


por_fila = 2 if MOVIL else 3
with st.container(key="grilla"):
    k = 0
    for i in range(0, len(v), por_fila):
        cols = st.columns(por_fila)
        for col, (_, p) in zip(cols, v.iloc[i:i + por_fila].iterrows()):
            with col:
                tarjeta(p, k)
            k += 1

st.markdown('<div class="lyho-pie">Las existencias se actualizan cada pocos minutos. '
            "Escríbenos para confirmar tu pedido.</div>", unsafe_allow_html=True)
if WHATSAPP:
    hola = quote("¡Hola! Vi tu catálogo y quiero hacer un pedido.")
    st.markdown(f'<a class="lyho-flot" href="https://wa.me/{WHATSAPP}?text={hola}" '
                'target="_blank">💬 Pedir por WhatsApp</a>', unsafe_allow_html=True)
