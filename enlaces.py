# -*- coding: utf-8 -*-
"""
Vistas previas de cada producto para WhatsApp, Instagram y Facebook.

Esas aplicaciones arman la tarjeta de un enlace leyendo la página SIN ejecutarla, y Streamlit
no permite poner esos datos por producto. Este servicio atiende solo las direcciones /p/...:

    https://lyhoshop.com/p/anillo-de-dragon  ->  página mínima con foto, nombre y precio,
                                                 que lleva a la persona al catálogo
    https://lyhoshop.com/p/img/<id>.jpg      ->  fotos subidas desde la app (para la tarjeta)

Corre en el servidor como servicio (lyho-enlaces, puerto 8503) y el túnel le manda lo que
empieza por /p/. Lee el mismo Google Sheet que el catálogo, con los mismos secrets.
"""
import base64
import html
import json
import os
import re
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import unquote

from comun import FOTO_PROPIA, fotos_de, slug

DIR = os.path.dirname(os.path.abspath(__file__))
PUERTO = int(os.environ.get("LYHO_ENLACES_PUERTO", "8503"))
CACHE_SEG = 300
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

_cache = {"t": 0.0, "datos": None}
_lock = threading.Lock()


def _secrets():
    try:
        import tomllib
    except ModuleNotFoundError:                 # Python < 3.11
        import tomli as tomllib
    with open(os.path.join(DIR, ".streamlit", "secrets.toml"), "rb") as f:
        return tomllib.load(f)


def _tabla(filas):
    """Lista de filas del Sheet -> lista de dicts con encabezados en mayúsculas."""
    if not filas:
        return []
    cab = [str(c).strip().upper() for c in filas[0]]
    return [dict(zip(cab, f + [""] * (len(cab) - len(f)))) for f in filas[1:]]


def _leer_hoja():
    s = _secrets()
    tabs = ("INVENTARIO", "VENTAS", "CONFIG", "LOTES", "FOTOS")
    if s.get("backend") == "gsheets":
        import gspread
        from google.oauth2.service_account import Credentials
        creds = Credentials.from_service_account_info(dict(s["gcp_service_account"]),
                                                      scopes=SCOPES)
        libro = gspread.authorize(creds).open_by_key(s["spreadsheet_id"])
        datos = {}
        for t in tabs:
            try:
                datos[t] = _tabla(libro.worksheet(t).get_values(
                    value_render_option="UNFORMATTED_VALUE"))
            except gspread.WorksheetNotFound:
                datos[t] = []
        return datos
    import pandas as pd                          # modo Excel local (pruebas)
    ruta = s.get("excel_path", os.path.join(DIR, "REGISTRO_SHEIN.xlsx"))
    datos = {}
    for t in tabs:
        try:
            df = pd.read_excel(ruta, sheet_name=t, dtype=str).fillna("")
            datos[t] = _tabla([list(df.columns)] + df.values.tolist())
        except Exception:
            datos[t] = []
    return datos


def _num(v):
    try:
        return float(str(v).replace(",", ".")) if str(v).strip() else None
    except ValueError:
        return None


def _ve(x):
    s = f"{x:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"${s}"


def datos():
    """Productos visibles por slug, fotos subidas por id y configuración (caché 5 min)."""
    with _lock:
        if _cache["datos"] is not None and time.time() - _cache["t"] < CACHE_SEG:
            return _cache["datos"]
        crudo = _leer_hoja()
        conf = {str(r.get("CLAVE", "")).upper(): r.get("VALOR", "") for r in crudo["CONFIG"]}
        vendidas = {}
        for r in crudo["VENTAS"]:
            n = str(r.get("NOMBRE", "")).strip().upper()
            vendidas[n] = vendidas.get(n, 0) + 1
        en_camino = {_num(r.get("LOTE")) for r in crudo["LOTES"]
                     if str(r.get("ESTADO", "")).strip().upper() == "EN CAMINO"}
        productos = {}
        for r in crudo["INVENTARIO"]:
            nombre = str(r.get("NOMBRE", "")).strip().upper()
            precio = _num(r.get("PRECIO_VENTA"))
            if (not nombre or not precio or precio <= 0
                    or str(r.get("OCULTAR", "")).strip().upper() == "SI"
                    or _num(r.get("LOTE")) in en_camino):
                continue
            quedan = int((_num(r.get("CANTIDAD")) or 0) - vendidas.get(nombre, 0))
            productos[slug(nombre)] = {"nombre": nombre, "precio": precio, "quedan": quedan,
                                       "desc": str(r.get("DESCRIPCION", "")).strip(),
                                       "fotos": fotos_de(r.get("FOTO"))}
        fotos = {}
        for r in crudo["FOTOS"]:
            id_ = str(r.get("ID", "") or "").strip()
            fotos[id_ if id_ else f"viejo:{str(r.get('NOMBRE', '')).strip()}"] = r.get("IMG", "")
        base = str(conf.get("CATALOGO_URL") or "https://lyhoshop.com").rstrip("/")
        _cache.update(t=time.time(), datos={"productos": productos, "fotos": fotos,
                                            "nombre": str(conf.get("CATALOGO_NOMBRE")
                                                          or "Lyho Shop"),
                                            "base": base})
        return _cache["datos"]


def url_foto(ref, nombre, base):
    if ref.startswith("http"):
        return "https://" + ref.split("://", 1)[1]          # las tarjetas piden https
    if ref.startswith("propia:"):
        return f"{base}/p/img/{ref[len('propia:'):]}.jpg"
    if ref == FOTO_PROPIA:
        return f"{base}/p/img/viejo-{slug(nombre)}.jpg"
    return None


def pagina(d, s):
    """HTML con los datos de la tarjeta y un salto inmediato al catálogo."""
    p = d["productos"].get(s)
    base, tienda = d["base"], d["nombre"]
    destino = f"{base}/?p={s}" if p else base + "/"
    if p:
        titulo = f"{p['nombre'].title()} · {_ve(p['precio'])}"
        desc = ("¡Última unidad! " if p["quedan"] == 1 else
                "Agotado por ahora. " if p["quedan"] <= 0 else "")
        desc += (p["desc"][:140] + " · " if p["desc"] else "") + f"Pídelo por WhatsApp en {tienda}"
        img = next((u for u in (url_foto(r, p["nombre"], base) for r in p["fotos"]) if u), None)
    else:
        titulo, desc, img = tienda, "Ropa y accesorios · pide por WhatsApp", None
    img = img or f"{base}/p/logo.png"
    e = lambda x: html.escape(str(x), quote=True)
    return f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8">
<title>{e(titulo)}</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta property="og:type" content="product">
<meta property="og:site_name" content="{e(tienda)}">
<meta property="og:title" content="{e(titulo)}">
<meta property="og:description" content="{e(desc)}">
<meta property="og:image" content="{e(img)}">
<meta property="og:url" content="{e(base + '/p/' + s)}">
<meta name="twitter:card" content="summary_large_image">
<meta http-equiv="refresh" content="0;url={e(destino)}">
<style>body{{font-family:sans-serif;background:#EEEDE9;color:#2E3B42;text-align:center;
padding:40px}} a{{color:#AE5166}}</style>
</head><body>
<p>Abriendo <a href="{e(destino)}">{e(titulo)}</a>…</p>
<script>location.replace({json.dumps(destino)});</script>
</body></html>"""


class Manejador(BaseHTTPRequestHandler):
    def _enviar(self, codigo, tipo, cuerpo, cache="public, max-age=300"):
        self.send_response(codigo)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(cuerpo)))
        self.send_header("Cache-Control", cache)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(cuerpo)

    def do_HEAD(self):
        self.do_GET()

    def do_GET(self):
        ruta = unquote(self.path.split("?", 1)[0])
        try:
            if ruta in ("/p/logo.png", "/p/logo"):
                with open(os.path.join(DIR, "assets", "logo_lyhoshop.png"), "rb") as f:
                    return self._enviar(200, "image/png", f.read(), "public, max-age=86400")
            m = re.fullmatch(r"/p/img/([A-Za-z0-9_\-]+)\.jpg", ruta)
            if m:
                d = datos()
                clave = m.group(1)
                if clave.startswith("viejo-"):          # formato viejo: por nombre
                    p = d["productos"].get(clave[len("viejo-"):])
                    clave = f"viejo:{p['nombre']}" if p else ""
                b64 = d["fotos"].get(clave)
                if not b64:
                    return self._enviar(404, "text/plain; charset=utf-8", b"no existe")
                return self._enviar(200, "image/jpeg", base64.b64decode(b64),
                                    "public, max-age=86400")
            m = re.fullmatch(r"/p/([a-z0-9\-]*)/?", ruta)
            if m:
                return self._enviar(200, "text/html; charset=utf-8",
                                    pagina(datos(), m.group(1)).encode("utf-8"))
            return self._enviar(404, "text/plain; charset=utf-8", b"no existe")
        except Exception as ex:                          # nunca tumbar el servicio
            self.log_error("error: %r", ex)
            return self._enviar(503, "text/plain; charset=utf-8", b"intenta de nuevo", "no-store")

    def log_message(self, formato, *args):              # registro corto en journalctl
        print(f"{self.address_string()} {formato % args}", flush=True)


if __name__ == "__main__":
    print(f"Vistas previas de Lyho Shop en 127.0.0.1:{PUERTO}", flush=True)
    ThreadingHTTPServer(("127.0.0.1", PUERTO), Manejador).serve_forever()
