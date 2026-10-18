# -*- coding: utf-8 -*-
"""
Recupera datos desde la hoja de respaldos diarios.

  python restaurar_respaldo.py <ID_RESPALDOS>                                  # lista fechas
  python restaurar_respaldo.py <ID_RESPALDOS> 2026-10-18                       # muestra su contenido
  python restaurar_respaldo.py <ID_RESPALDOS> 2026-10-18 VENTAS <ID_PRINCIPAL> # restaura una pestaña
  python restaurar_respaldo.py <ID_RESPALDOS> 2026-10-18 TODAS  <ID_PRINCIPAL> # restaura todo

Antes de escribir guarda una copia local de lo que vas a pisar (antes_de_restaurar_*.xlsx)
y pide confirmación.
"""
import datetime as dt
import sys

import pandas as pd

from migrar_sheet import SCOPES


def bloques(valores):
    """Matriz de la pestaña de respaldo -> {TAB: DataFrame}."""
    salida, tab, filas = {}, None, []

    def cerrar():
        if tab and filas:
            cab = [str(c).strip() for c in filas[0] if str(c).strip()]
            datos = [list(f[:len(cab)]) + [""] * (len(cab) - len(f[:len(cab)]))
                     for f in filas[1:]]
            salida[tab] = pd.DataFrame(datos, columns=cab)
        elif tab:
            salida[tab] = pd.DataFrame()

    for fila in valores:
        fila = list(fila)
        primera = str(fila[0]).strip() if fila else ""
        if primera.startswith("### "):
            cerrar()
            tab, filas = primera[4:].strip(), []
        elif tab is not None:
            if not any(str(c).strip() for c in fila):
                cerrar()
                tab, filas = None, []
            else:
                filas.append(fila)       # se conservan números como números
    cerrar()
    return salida


def main():
    import gspread
    from google.oauth2.service_account import Credentials

    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    gc = gspread.authorize(
        Credentials.from_service_account_file("credenciales.json", scopes=SCOPES))
    resp = gc.open_by_key(sys.argv[1])

    if len(sys.argv) == 2:
        fechas = sorted((w.title for w in resp.worksheets() if w.title[:2] == "20"),
                        reverse=True)
        print("Respaldos disponibles:", ", ".join(fechas) or "ninguno")
        return

    fecha = sys.argv[2]
    datos = bloques(resp.worksheet(fecha).get_values(value_render_option="UNFORMATTED_VALUE"))
    print(f"Respaldo {fecha}:")
    for tab, df in datos.items():
        print(f"  {tab}: {len(df)} filas")
    if len(sys.argv) < 5:
        print("\nPara restaurar: agrega la pestaña (o TODAS) y el ID de la hoja principal.")
        return

    cual, id_principal = sys.argv[3].upper(), sys.argv[4]
    objetivos = list(datos) if cual == "TODAS" else [cual]
    faltan = [t for t in objetivos if t not in datos]
    if faltan:
        print("No están en el respaldo:", ", ".join(faltan))
        sys.exit(1)

    principal = gc.open_by_key(id_principal)
    copia = f"antes_de_restaurar_{dt.datetime.now():%Y%m%d_%H%M}.xlsx"
    with pd.ExcelWriter(copia, engine="openpyxl") as w:
        for tab in objetivos:
            try:
                actual = principal.worksheet(tab).get_values()
                pd.DataFrame(actual[1:], columns=actual[0] if actual else None) \
                    .to_excel(w, sheet_name=tab[:31], index=False)
            except gspread.WorksheetNotFound:
                pass
    print(f"\nCopia de lo actual -> {copia}")

    if input(f"Sobrescribir {', '.join(objetivos)} con el respaldo del {fecha}? (s/n) ") \
            .strip().lower() != "s":
        print("Cancelado.")
        return
    for tab in objetivos:
        df = datos[tab]
        try:
            ws = principal.worksheet(tab)
            ws.clear()
        except gspread.WorksheetNotFound:
            ws = principal.add_worksheet(tab, rows=max(1000, len(df) + 200),
                                         cols=len(df.columns) + 2)
        valores = [list(df.columns)] + [["" if v is None else v for v in f]
                                         for f in df.values.tolist()]
        ws.update(values=valores, range_name="A1", value_input_option="RAW")
        print(f"  restaurada {tab}: {len(df)} filas")
    print("\nListo. En la app, usa «🔄 Refrescar datos».")


if __name__ == "__main__":
    main()
