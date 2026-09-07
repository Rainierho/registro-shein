# Registro SHEIN — app Streamlit

App para cargar ventas/inventario, ver estadísticas y mantener el histórico que hoy vive en el Excel.

```
app.py                        la app completa
migrar.py                     pasa el Excel original al formato de la app / lo sube a Sheets
migrar_sheet.py               normaliza el Google Sheet que ya tienes en Drive
requirements.txt
REGISTRO_SHEIN.xlsx           datos ya migrados (119 productos, 305 ventas, 8 deudas, 5 cajas)
.streamlit/secrets.toml.ejemplo
```

## Secciones

- **Dashboard**: ingresos, costo, ganancia neta, margen, ticket promedio, stock, ganancia potencial, inversión. Series por mes, acumulado, top productos, invertido vs recuperado por lote, agotados y stock.
- **Ventas**: alta rápida (elige producto → precio y lote se autocompletan, permite N unidades), filtros, y editor del histórico.
- **Inventario**: alta de producto/lote y editor. `EXISTENCIA` se calcula contando VENTAS por `NOMBRE` (misma lógica del `COUNTIF` del Excel).
- **Deudas** y **Caja Chica**: alta/actualización por clave y editor.

## Cambios respecto al Excel

- La columna `e` pasó a llamarse `DESCRIPCION`; la hoja `Caja Chica` pasó a `CAJA`.
- Se eliminaron las columnas con fórmulas (`EXISTENCIA`, `INVENTARIO`, `GANANCIA POTENCIAL/TOTAL/HASTA AHORA`): la app las calcula.
- Lo que el Excel llamaba "ganancia" era ingreso (no restaba el costo). Aquí se separan **Ingresos** y **Ganancia neta** = ingreso − precio de compra.
- Las 11 ventas cuyo producto no existe en INVENTARIO (CAMISA BEIGE, FAJA, PANTALON, etc.) se cuentan con costo 0 y aparecen advertidas en el Dashboard.
- Se quitaron las filas de TOTAL: los totales los calcula la app.

## Despliegue en Streamlit Community Cloud

El disco de Streamlit Cloud es efímero: escribir el .xlsx ahí **no persiste**. Por eso producción usa Google Sheets.

**1. Google Sheet**
- Crea una hoja nueva en Drive, llámala `REGISTRO_SHEIN_APP`. El ID es lo que va entre `/d/` y `/edit` en la URL.

**2. Service account**
- console.cloud.google.com → proyecto nuevo → APIs habilitadas: *Google Sheets API* y *Google Drive API*.
- IAM → Cuentas de servicio → crear → Claves → Agregar clave → JSON. Descarga el archivo.
- Comparte el Google Sheet (Editor) con el `client_email` del JSON.

**3. Cargar el histórico**

Si tu registro ya vive en un Google Sheet (caso normal), no hace falta pasar por el Excel:

```bash
pip install -r requirements.txt
# renombra el JSON descargado a credenciales.json
python migrar_sheet.py <ID_DE_TU_SHEET>                    # respalda y normaliza el mismo archivo
python migrar_sheet.py <ID_ORIGEN> --destino <ID_NUEVO>    # o deja el original intacto
```

Baja un `respaldo_AAAAMMDD.xlsx` con todo tal cual antes de tocar nada, y pide confirmación antes de escribir.

Partiendo del Excel:

```bash
python migrar.py REGISTRO_SHEIN_ORIGINAL.xlsx --sheets <ID_DEL_SHEET>
```

**4. Repo y deploy**
- Sube la carpeta a GitHub (**sin** `credenciales.json` ni `secrets.toml`).
- share.streamlit.io → New app → tu repo → `app.py`.
- Settings → Secrets: pega el contenido de `.streamlit/secrets.toml.ejemplo` con tus valores. El `private_key` va entre comillas y con los `\n` literales tal cual vienen en el JSON.

## Correr local

```bash
pip install -r requirements.txt
streamlit run app.py
```
Sin `secrets.toml` arranca en modo `excel` y lee/escribe `REGISTRO_SHEIN.xlsx` de la carpeta. Para probar contra Sheets desde local, crea `.streamlit/secrets.toml` con `backend = "gsheets"`.

## Notas

- `NOMBRE` en INVENTARIO es la clave que une con VENTAS: si repites un nombre en dos lotes, el conteo de existencia se mezcla. La app bloquea altas con nombre duplicado.
- El botón **Descargar Excel** de la barra lateral genera un respaldo con todo, incluidas las columnas calculadas.
