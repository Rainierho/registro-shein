# -*- coding: utf-8 -*-
"""
Convierte la pestaña DEUDAS en ENCARGOS, que es lo que realmente son:
pedidos que compraste con plata de SHEIN y le cobras a la persona.

  python migrar_encargos.py <ID_DEL_SHEET>

Cada deudor pasa a una fila de ENCARGOS con COBRAR = lo que debe y COSTO en blanco.
La app trata el costo vacío como "desconocido" y no lo mete en la ganancia.
DEUDAS queda intacta como archivo; la app ya no la usa.
"""
import sys

import pandas as pd

from migrar_sheet import SCOPES, celdas

COLS = ["PERSONA", "TIPO", "DESCRIPCION", "COSTO", "COBRAR", "ABONADO", "FECHA",
        "NOTA"]


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    id_sheet = sys.argv[1]

    import gspread
    from google.oauth2.service_account import Credentials

    gc = gspread.authorize(
        Credentials.from_service_account_file("credenciales.json", scopes=SCOPES))
    libro = gc.open_by_key(id_sheet)

    valores = libro.worksheet("DEUDAS").get_values(value_render_option="UNFORMATTED_VALUE")
    cab = [str(c).strip().upper() for c in valores[0]]
    d = pd.DataFrame(valores[1:], columns=cab)
    d = d[d["DEUDOR"].astype(str).str.strip() != ""]

    filas = []
    for _, r in d.iterrows():
        nota = []
        if str(r.get("BOLOS", "")).strip():
            nota.append(f"bolos: {r['BOLOS']}")
        if str(r.get("NOTA", "")).strip():
            nota.append(str(r["NOTA"]))
        filas.append({
            "PERSONA": str(r["DEUDOR"]).strip().upper(),
            "TIPO": "",                        # ENCARGO o PRÉSTAMO, lo clasificas en la app
            "DESCRIPCION": "",
            "COSTO": "",                       # desconocido
            "COBRAR": pd.to_numeric(r["MONTO"], errors="coerce"),
            "ABONADO": 0,
            "FECHA": "",
            "NOTA": " · ".join(nota),
        })
    enc = pd.DataFrame(filas, columns=COLS)

    print(f"ENCARGOS: {len(enc)} filas · por cobrar "
          f"${pd.to_numeric(enc['COBRAR']).sum():,.2f}")
    print(enc[["PERSONA", "COBRAR", "NOTA"]].to_string(index=False))
    print("\nTIPO y COSTO quedan vacíos: clasifica cada fila como ENCARGO o PRÉSTAMO "
          "desde la app.")

    try:
        ws = libro.worksheet("ENCARGOS")
        actuales = len(ws.get_values()) - 1
        if actuales > 0:
            print(f"\nENCARGOS ya tiene {actuales} filas. Se reemplazarían.")
    except gspread.WorksheetNotFound:
        ws = None

    if input("\nCrear/reemplazar la pestaña ENCARGOS? (s/n) ").strip().lower() != "s":
        print("Cancelado.")
        return

    if ws is None:
        ws = libro.add_worksheet("ENCARGOS", rows=500, cols=len(COLS) + 2)
    else:
        ws.clear()
    ws.update(values=celdas(enc), range_name="A1", value_input_option="RAW")
    ws.freeze(rows=1)
    print(f"  escrito ENCARGOS: {len(enc)} filas")
    print("\nDEUDAS quedó intacta como archivo.")


if __name__ == "__main__":
    main()
