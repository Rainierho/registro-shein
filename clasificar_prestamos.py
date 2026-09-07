# -*- coding: utf-8 -*-
"""
Marca como PRÉSTAMO las filas de ENCARGOS que no tienen tipo, poniendo COSTO = COBRAR
(sale plata de SHEIN y vuelve igual, ganancia cero).

  python clasificar_prestamos.py <ID_DEL_SHEET>

Solo toca filas con TIPO vacío. Las que ya clasificaste, y las que tengan COSTO puesto,
se dejan como están.
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
    ws = gc.open_by_key(id_sheet).worksheet("ENCARGOS")

    valores = ws.get_values(value_render_option="UNFORMATTED_VALUE")
    ancho = max(len(f) for f in valores)
    filas = [list(f) + [""] * (ancho - len(f)) for f in valores]
    df = pd.DataFrame(filas[1:], columns=[str(c).strip().upper() for c in filas[0]],
                      dtype=object)
    for c in COLS:
        if c not in df.columns:
            df[c] = pd.Series([""] * len(df), dtype=object)
    df = df[COLS].astype(object)
    df = df[df["PERSONA"].astype(str).str.strip() != ""]

    sin_tipo = df["TIPO"].astype(str).str.strip() == ""
    if not sin_tipo.any():
        print("No hay filas sin clasificar. Nada que hacer.")
        return

    df.loc[sin_tipo, "TIPO"] = "PRÉSTAMO"
    vacio_costo = sin_tipo & (df["COSTO"].astype(str).str.strip() == "")
    df.loc[vacio_costo, "COSTO"] = df.loc[vacio_costo, "COBRAR"]

    print(f"{int(sin_tipo.sum())} filas pasan a PRÉSTAMO con COSTO = COBRAR:\n")
    print(df.loc[sin_tipo, ["PERSONA", "TIPO", "COSTO", "COBRAR"]].to_string(index=False))
    total = pd.to_numeric(df.loc[sin_tipo, "COBRAR"], errors="coerce").sum()
    print(f"\nTotal: ${total:,.2f} · ganancia cero, todo es recuperación.")

    if input("\nAplicar? (s/n) ").strip().lower() != "s":
        print("Cancelado.")
        return

    ws.clear()
    ws.update(values=celdas(df), range_name="A1", value_input_option="RAW")
    ws.freeze(rows=1)
    print(f"  actualizado ENCARGOS: {len(df)} filas")


if __name__ == "__main__":
    main()
