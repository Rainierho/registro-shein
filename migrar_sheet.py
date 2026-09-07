# -*- coding: utf-8 -*-
"""
Convierte el Google Sheet que YA tienes en Drive en la base de datos de la app.

  python migrar_sheet.py <ID_DEL_SHEET>                  # respalda y normaliza el mismo archivo
  python migrar_sheet.py <ID_ORIGEN> --destino <ID_NUEVO>  # deja el original intacto
  python migrar_sheet.py <ID_DEL_SHEET> --solo-respaldo  # solo descarga respaldo, no escribe

Necesita credenciales.json (service account) en la misma carpeta y que el Sheet
este compartido como Editor con el client_email de esa cuenta.

Que hace:
  - Baja todas las pestañas tal cual y las guarda en respaldo_AAAAMMDD.xlsx (local).
  - Renombra 'e' -> DESCRIPCION, 'Caja Chica' -> CAJA.
  - Elimina las columnas con formulas (EXISTENCIA, INVENTARIO, GANANCIA *) y las filas TOTAL.
  - Reescribe INVENTARIO, VENTAS, DEUDAS y CAJA con el esquema que usa app.py.
"""
import datetime as dt
import sys

import pandas as pd

ESQUEMA = {
    "INVENTARIO": ["DESCRIPCION", "NOMBRE", "CANTIDAD", "PRECIO_COMPRA", "PRECIO_VENTA", "LOTE"],
    "VENTAS": ["NOMBRE", "MONTO", "FECHA", "LOTE"],
    "DEUDAS": ["DEUDOR", "MONTO", "BOLOS", "NOTA"],
    "CAJA": ["UBICACION", "MONTO", "MONEDA", "TASA"],
}
# pestaña en tu Sheet -> pestaña de la app
ORIGEN = {"INVENTARIO": "INVENTARIO", "VENTAS": "VENTAS", "DEUDAS": "DEUDAS", "CAJA": "Caja Chica"}
ALIAS = {"E": "DESCRIPCION", "UBICACION": "UBICACION", "UBICACIÓN": "UBICACION",
         "MONEDA": "MONEDA", "TASA": "TASA", "DEUDOR": "DEUDOR"}
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]
EPOCA_SHEETS = dt.datetime(1899, 12, 30)


def a_dataframe(valores):
    """Matriz cruda de Sheets -> DataFrame con encabezados normalizados."""
    if not valores:
        return pd.DataFrame()
    ancho = max(len(f) for f in valores)
    filas = [f + [""] * (ancho - len(f)) for f in valores]
    cabecera = [str(c).strip().upper() for c in filas[0]]
    cabecera = [ALIAS.get(c, c) for c in cabecera]
    cabecera = [c if c else f"COL_{i}" for i, c in enumerate(cabecera)]
    return pd.DataFrame(filas[1:], columns=cabecera)


def _fecha(v):
    if v in ("", None):
        return None
    if isinstance(v, (int, float)):
        return EPOCA_SHEETS + dt.timedelta(days=float(v))
    return pd.to_datetime(v, dayfirst=True, errors="coerce")


def normalizar(df, tab):
    cols = ESQUEMA[tab]
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df.loc[:, ~df.columns.duplicated()][cols].copy()

    clave = {"INVENTARIO": "NOMBRE", "VENTAS": "NOMBRE",
             "DEUDAS": "DEUDOR", "CAJA": "UBICACION"}[tab]
    df[clave] = df[clave].fillna("").astype(str).str.strip()
    df = df[df[clave] != ""]
    df = df[~df[clave].str.upper().isin(["TOTAL", "TOTALES", "NAN", "NONE"])]

    numericas = {"INVENTARIO": ["CANTIDAD", "PRECIO_COMPRA", "PRECIO_VENTA", "LOTE"],
                 "VENTAS": ["MONTO", "LOTE"], "DEUDAS": ["MONTO", "BOLOS"],
                 "CAJA": ["MONTO", "TASA"]}[tab]
    for c in numericas:
        df[c] = pd.to_numeric(
            df[c].astype(str).str.replace(r"[^\d.,\-]", "", regex=True).str.replace(",", "."),
            errors="coerce")
    if tab == "VENTAS":
        df["FECHA"] = df["FECHA"].map(_fecha)
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce").dt.strftime("%Y-%m-%d")
    for c in cols:
        if c not in numericas and c != "FECHA":
            df[c] = df[c].fillna("").astype(str).str.strip().replace({"nan": "", "None": ""})
    return df.reset_index(drop=True)


def celdas(df):
    salida = [list(df.columns)]
    for _, fila in df.iterrows():
        salida.append(["" if pd.isna(v) else (float(v) if isinstance(v, (int, float)) else str(v))
                       for v in fila])
    return salida


def main():
    import gspread
    from google.oauth2.service_account import Credentials

    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    id_origen = sys.argv[1]
    id_destino = sys.argv[sys.argv.index("--destino") + 1] if "--destino" in sys.argv else id_origen
    solo_respaldo = "--solo-respaldo" in sys.argv

    gc = gspread.authorize(
        Credentials.from_service_account_file("credenciales.json", scopes=SCOPES))
    libro = gc.open_by_key(id_origen)
    print(f"Origen: {libro.title}")
    print("Pestañas:", [w.title for w in libro.worksheets()])

    # 1) respaldo local de TODO tal cual
    crudos = {}
    for ws in libro.worksheets():
        crudos[ws.title] = ws.get_values(value_render_option="UNFORMATTED_VALUE")
    respaldo = f"respaldo_{dt.date.today():%Y%m%d}.xlsx"
    with pd.ExcelWriter(respaldo, engine="openpyxl") as w:
        for nombre, valores in crudos.items():
            a_dataframe(valores).to_excel(w, sheet_name=nombre[:31], index=False)
    print(f"Respaldo -> {respaldo}")
    if solo_respaldo:
        return

    # 2) normalizar
    datos = {}
    for destino, origen in ORIGEN.items():
        if origen not in crudos:
            print(f"  ! falta la pestaña '{origen}', se crea vacía")
            datos[destino] = pd.DataFrame(columns=ESQUEMA[destino])
            continue
        datos[destino] = normalizar(a_dataframe(crudos[origen]), destino)
        print(f"  {origen} -> {destino}: {len(datos[destino])} filas")

    if input(f"\nEscribir en {id_destino}? (s/n) ").strip().lower() != "s":
        print("Cancelado.")
        return

    # 3) escribir
    salida = gc.open_by_key(id_destino)
    for tab, df in datos.items():
        try:
            ws = salida.worksheet(tab)
            ws.clear()
        except gspread.WorksheetNotFound:
            ws = salida.add_worksheet(tab, rows=max(1000, len(df) + 300), cols=len(df.columns) + 2)
        ws.update(values=celdas(df), range_name="A1", value_input_option="RAW")
        ws.freeze(rows=1)
        print(f"  escrito {tab}: {len(df)} filas")
    print("\nListo. spreadsheet_id para los secrets:", id_destino)
    print("Si la pestaña vieja 'Caja Chica' sigue ahí, ya no se usa: puedes borrarla a mano.")


if __name__ == "__main__":
    main()
