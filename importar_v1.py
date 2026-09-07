# -*- coding: utf-8 -*-
"""
Importa el archivo viejo (Negocio_Shein.xlsx, lotes 1-10) como histórico de solo lectura.

  python importar_v1.py Negocio_Shein.xlsx <ID_DEL_SHEET>

Crea dos pestañas nuevas y NO toca las operativas:
  VENTAS_HIST      200 ventas de ene-ago 2025
  INVENTARIO_HIST  308 compras, con propietario, talla, tipo (I=stock, P=pedido de alguien)

No se fusiona con INVENTARIO ni VENTAS porque en el v1 la existencia se llevaba a mano:
solo 21 de 200 ventas cruzan por nombre con su propio inventario, hay 15 nombres repetidos
y 159 filas son pedidos de terceros. Unirlo rompería el conteo de existencias y el margen.
"""
import sys

import pandas as pd

from migrar_sheet import SCOPES, celdas

HIST = {
    "VENTAS_HIST": ["NOMBRE", "MONTO", "FECHA", "LOTE"],
    "INVENTARIO_HIST": ["NOMBRE", "SKU", "PROPIETARIO", "TALLA", "CANTIDAD", "PRECIO_COMPRA",
                        "PRECIO_VENTA", "TIPO", "CATEGORIA", "LOTE", "FECHA_LOTE"],
}


def convertir(ruta):
    ven = pd.read_excel(ruta, sheet_name="Ventas")
    ven.columns = [str(c).strip().upper() for c in ven.columns]
    ven = ven.rename(columns={"PRECIO": "MONTO", "LOTE": "LOTE"})
    ven = ven[ven["NOMBRE"].notna()]
    ven["FECHA"] = pd.to_datetime(ven["FECHA"], errors="coerce").dt.strftime("%Y-%m-%d")
    ven["NOMBRE"] = ven["NOMBRE"].astype(str).str.strip().str.upper()
    ven = ven[HIST["VENTAS_HIST"]]

    inv = pd.read_excel(ruta, sheet_name="INVENTARIO")
    inv.columns = [str(c).strip().upper() for c in inv.columns]
    inv = inv.rename(columns={
        "PRECIO DE COMPRA (CON CUPONES Y PUNTOS)": "PRECIO_COMPRA",
        "TIPO DE PRODUCTO": "CATEGORIA",
        "FECHA DE ENTREGA": "FECHA_LOTE",
    })
    inv = inv[inv["NOMBRE"].notna()]
    inv["NOMBRE"] = inv["NOMBRE"].astype(str).str.strip().str.upper()
    inv["FECHA_LOTE"] = pd.to_datetime(inv["FECHA_LOTE"], errors="coerce").dt.strftime("%Y-%m-%d")
    for c in HIST["INVENTARIO_HIST"]:
        if c not in inv.columns:
            inv[c] = ""
    inv = inv[HIST["INVENTARIO_HIST"]]

    for df in (ven, inv):
        for c in df.columns:
            if df[c].dtype == object:
                df[c] = df[c].fillna("").astype(str).str.strip().replace({"nan": "", "NaT": ""})
    return {"VENTAS_HIST": ven.reset_index(drop=True), "INVENTARIO_HIST": inv.reset_index(drop=True)}


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    ruta, id_sheet = sys.argv[1], sys.argv[2]
    datos = convertir(ruta)

    inv = datos["INVENTARIO_HIST"]
    ven = datos["VENTAS_HIST"]
    stock = pd.to_numeric(inv[inv["TIPO"] == "I"]["CANTIDAD"], errors="coerce") * \
        pd.to_numeric(inv[inv["TIPO"] == "I"]["PRECIO_COMPRA"], errors="coerce")
    print(f"VENTAS_HIST: {len(ven)} filas · ${pd.to_numeric(ven['MONTO']).sum():,.2f}")
    print(f"INVENTARIO_HIST: {len(inv)} filas · compras de stock ${stock.sum():,.2f}")

    import gspread
    from google.oauth2.service_account import Credentials

    gc = gspread.authorize(
        Credentials.from_service_account_file("credenciales.json", scopes=SCOPES))
    libro = gc.open_by_key(id_sheet)

    if input(f"\nCrear las pestañas históricas en {id_sheet}? (s/n) ").strip().lower() != "s":
        print("Cancelado.")
        return

    for tab, df in datos.items():
        try:
            ws = libro.worksheet(tab)
            ws.clear()
        except gspread.WorksheetNotFound:
            ws = libro.add_worksheet(tab, rows=len(df) + 200, cols=len(df.columns) + 2)
        ws.update(values=celdas(df), range_name="A1", value_input_option="RAW")
        ws.freeze(rows=1)
        print(f"  escrito {tab}: {len(df)} filas")
    print("\nListo. Las pestañas operativas no se tocaron.")


if __name__ == "__main__":
    main()
