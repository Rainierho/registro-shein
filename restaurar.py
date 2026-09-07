# -*- coding: utf-8 -*-
"""
Restaura una pestaña del Sheet desde el respaldo local.

  python restaurar.py respaldo_20260906.xlsx DEUDAS <ID_DEL_SHEET>
  python restaurar.py respaldo_20260906.xlsx TODAS <ID_DEL_SHEET>

Lee el respaldo (que tiene los datos tal como estaban antes de migrar), lo
normaliza igual que migrar_sheet.py y reescribe solo esa pestaña.
"""
import sys

import pandas as pd

from migrar_sheet import ESQUEMA, ORIGEN, SCOPES, celdas, normalizar


def main():
    if len(sys.argv) < 4:
        print(__doc__)
        sys.exit(1)
    ruta, cual, id_sheet = sys.argv[1], sys.argv[2].upper(), sys.argv[3]

    import gspread
    from google.oauth2.service_account import Credentials

    gc = gspread.authorize(
        Credentials.from_service_account_file("credenciales.json", scopes=SCOPES))
    libro = gc.open_by_key(id_sheet)

    objetivos = list(ESQUEMA) if cual == "TODAS" else [cual]
    for tab in objetivos:
        if tab not in ESQUEMA:
            print(f"Pestaña desconocida: {tab}. Opciones: {list(ESQUEMA)} o TODAS")
            sys.exit(1)

        origen = ORIGEN[tab]  # nombre de la pestaña dentro del respaldo
        df = pd.read_excel(ruta, sheet_name=origen)
        df.columns = [str(c).strip().upper() for c in df.columns]
        df = df.rename(columns={"E": "DESCRIPCION"})
        datos = normalizar(df, tab)

        print(f"\n{origen} -> {tab}: {len(datos)} filas")
        print(datos.head(10).to_string(index=False))
        if input(f"Sobrescribir la pestaña {tab}? (s/n) ").strip().lower() != "s":
            print("  saltada")
            continue

        try:
            ws = libro.worksheet(tab)
            ws.clear()
        except gspread.WorksheetNotFound:
            ws = libro.add_worksheet(tab, rows=max(1000, len(datos) + 300),
                                     cols=len(datos.columns) + 2)
        ws.update(values=celdas(datos), range_name="A1", value_input_option="RAW")
        ws.freeze(rows=1)
        print(f"  restaurada {tab}")


if __name__ == "__main__":
    main()
