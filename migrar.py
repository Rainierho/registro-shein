# -*- coding: utf-8 -*-
"""
Migra el REGISTRO_SHEIN.xlsx original al formato que usa la app.

  python migrar.py REGISTRO_SHEIN_ORIGINAL.xlsx                 -> genera REGISTRO_SHEIN.xlsx limpio
  python migrar.py REGISTRO_SHEIN_ORIGINAL.xlsx --sheets ID     -> ademas sube todo a Google Sheets

Para --sheets necesitas credenciales.json (service account) en la misma carpeta.
"""
import sys
import pandas as pd

TABS = {
    "INVENTARIO": ["DESCRIPCION", "NOMBRE", "CANTIDAD", "PRECIO_COMPRA", "PRECIO_VENTA", "LOTE"],
    "VENTAS": ["NOMBRE", "MONTO", "FECHA", "LOTE"],
    "DEUDAS": ["DEUDOR", "MONTO", "BOLOS", "NOTA"],
    "CAJA": ["UBICACION", "MONTO", "MONEDA", "TASA"],
}


def convertir(origen):
    xl = pd.ExcelFile(origen)

    inv = xl.parse("INVENTARIO")
    inv.columns = [str(c).strip().upper() for c in inv.columns]
    inv = inv.rename(columns={"E": "DESCRIPCION"})
    inv = inv[inv["NOMBRE"].notna()]
    inv = inv[TABS["INVENTARIO"]]

    ven = xl.parse("VENTAS")
    ven.columns = [str(c).strip().upper() for c in ven.columns]
    ven = ven[ven["NOMBRE"].notna()][TABS["VENTAS"]]
    ven["FECHA"] = pd.to_datetime(ven["FECHA"], errors="coerce").dt.strftime("%Y-%m-%d")

    deu = xl.parse("DEUDAS")
    deu.columns = [str(c).strip().upper() for c in deu.columns]
    deu = deu.rename(columns={"DEUDOR ": "DEUDOR"})
    deu = deu[deu["DEUDOR"].notna()]
    deu["NOTA"] = ""
    deu = deu[TABS["DEUDAS"]]

    caja = xl.parse("Caja Chica")
    caja.columns = [str(c).strip().upper() for c in caja.columns]
    caja = caja.rename(columns={"UBICACION": "UBICACION"})
    caja = caja[caja["UBICACION"].notna()]
    caja = caja[~caja["UBICACION"].astype(str).str.upper().str.strip().eq("TOTAL")]
    caja = caja[TABS["CAJA"]]

    for df in (inv, ven, deu, caja):
        for c in df.columns:
            if df[c].dtype == object:
                df[c] = df[c].astype(str).str.strip().replace({"nan": ""})

    return {"INVENTARIO": inv, "VENTAS": ven, "DEUDAS": deu, "CAJA": caja}


def a_excel(datos, destino="REGISTRO_SHEIN.xlsx"):
    with pd.ExcelWriter(destino, engine="openpyxl") as w:
        for tab, df in datos.items():
            df.to_excel(w, sheet_name=tab, index=False)
    print(f"OK -> {destino}")
    for t, d in datos.items():
        print(f"   {t}: {len(d)} filas")


def a_sheets(datos, spreadsheet_id, cred="credenciales.json"):
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    gc = gspread.authorize(Credentials.from_service_account_file(cred, scopes=scopes))
    libro = gc.open_by_key(spreadsheet_id)

    for tab, df in datos.items():
        try:
            ws = libro.worksheet(tab)
            ws.clear()
        except gspread.WorksheetNotFound:
            ws = libro.add_worksheet(tab, rows=max(1000, len(df) + 200), cols=len(df.columns) + 2)
        cuerpo = [list(df.columns)] + df.where(pd.notna(df), "").values.tolist()
        cuerpo = [[float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else str(v)
                   for v in fila] for fila in cuerpo]
        ws.update(values=cuerpo, range_name="A1", value_input_option="RAW")
        print(f"   subido {tab}: {len(df)} filas")
    print("OK -> Google Sheets")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    datos = convertir(sys.argv[1])
    a_excel(datos)
    if "--sheets" in sys.argv:
        a_sheets(datos, sys.argv[sys.argv.index("--sheets") + 1])
