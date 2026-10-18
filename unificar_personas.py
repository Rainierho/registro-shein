# -*- coding: utf-8 -*-
"""
Unifica nombres de personas escritos distinto en ENCARGOS.

  python unificar_personas.py <ID_DEL_SHEET>                       # solo lista los nombres
  python unificar_personas.py <ID_DEL_SHEET> "YANIS EN BOLOS" YANIS  # renombra

Las filas conservan su monto; solo cambia el nombre, así que a partir de ahí se agrupan
en "Quién debe" y aparecen una sola vez en el desplegable.
"""
import sys

import pandas as pd

from migrar_sheet import SCOPES, celdas

COLS = ["PERSONA", "TIPO", "DESCRIPCION", "COSTO", "COBRAR", "ABONADO", "MONEDA",
        "FECHA", "ULT_ABONO", "NOTA"]


def cargar(ws):
    valores = ws.get_values(value_render_option="UNFORMATTED_VALUE")
    ancho = max(len(f) for f in valores)
    filas = [list(f) + [""] * (ancho - len(f)) for f in valores]
    df = pd.DataFrame(filas[1:], columns=[str(c).strip().upper() for c in filas[0]],
                      dtype=object)
    for c in COLS:
        if c not in df.columns:
            df[c] = pd.Series([""] * len(df), dtype=object)
    df = df[COLS + [c for c in df.columns if c not in COLS]].astype(object)
    return df[df["PERSONA"].astype(str).str.strip() != ""]


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
    df = cargar(ws)

    if len(sys.argv) < 4:
        resumen = (df.assign(COBRAR=pd.to_numeric(df["COBRAR"], errors="coerce"),
                             ABONADO=pd.to_numeric(df["ABONADO"], errors="coerce").fillna(0))
                   .groupby("PERSONA")
                   .agg(FILAS=("COBRAR", "size"), TOTAL=("COBRAR", "sum"),
                        ABONADO=("ABONADO", "sum")).reset_index()
                   .sort_values("TOTAL", ascending=False))
        resumen["SALDO"] = resumen["TOTAL"] - resumen["ABONADO"]
        print(resumen.to_string(index=False))
        print("\nPara unificar dos nombres:")
        print(f'  python unificar_personas.py {id_sheet} "NOMBRE VIEJO" "NOMBRE NUEVO"')
        return

    viejo, nuevo = sys.argv[2].strip().upper(), sys.argv[3].strip().upper()
    afectadas = df["PERSONA"].astype(str).str.strip().str.upper() == viejo
    if not afectadas.any():
        print(f"No hay filas con PERSONA = {viejo}")
        return

    print(f"{int(afectadas.sum())} filas pasan de {viejo} a {nuevo}:\n")
    print(df.loc[afectadas, ["PERSONA", "TIPO", "COBRAR", "ABONADO"]].to_string(index=False))
    if input("\nAplicar? (s/n) ").strip().lower() != "s":
        print("Cancelado.")
        return

    df.loc[afectadas, "PERSONA"] = nuevo
    ws.clear()
    ws.update(values=celdas(df), range_name="A1", value_input_option="RAW")
    ws.freeze(rows=1)
    print(f"  actualizado ENCARGOS: {len(df)} filas")


if __name__ == "__main__":
    main()
