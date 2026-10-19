# Lyho Shop en tu propio servidor

Resultado: `lyhoshop.com` → catálogo · `app.lyhoshop.com` → tu registro (solo tú entras).

## 0. Sube esta carpeta a GitHub

El servidor descarga el script desde tu repo, así que primero:

```powershell
git add .
git commit -m "Despliegue en VPS"
git push
```

## 1. Comprar el VPS

- Plan más pequeño que tenga **2 GB de RAM o más**.
- Región: **Estados Unidos, costa este** (Nueva York).
- Sistema: **Ubuntu 24.04**.
- Contraseña de root: una larga, guárdala bien.
- Sin extras.

Al terminar te llega por correo la **IP** del servidor.

## 2. Cloudflare: quitar lo provisional

En tu dominio → **DNS → Records**: borra el registro **A `lyhoshop.com`** (192.0.2.1) y el **CNAME `www`**.
Si creaste la regla de redirección, bórrala en **Rules → Redirect Rules**.

(El túnel crea sus propios registros; si estos siguen ahí, da error.)

## 3. Crear el túnel

**Zero Trust → Networks → Tunnels → Create a tunnel → Cloudflared**, nombre `lyhoshop`.

Te muestra un comando de instalación. Copia **solo el token**: el texto largo después de
`service install`. Lo usas en el paso 5.

En el mismo asistente, **Public hostnames** (agrega tres):

| Subdominio | Dominio | Tipo | URL |
|---|---|---|---|
| *(vacío)* | lyhoshop.com | HTTP | `localhost:8502` |
| www | lyhoshop.com | HTTP | `localhost:8502` |
| app | lyhoshop.com | HTTP | `localhost:8501` |

## 4. Proteger `app.lyhoshop.com`

**Zero Trust → Access → Applications → Add an application → Self-hosted**

- Dominio: `app.lyhoshop.com`
- Política: **Allow** · Include → **Emails** → tu correo
- Método de inicio: **One-time PIN**

Hazlo **antes** del paso 5, para que la app nunca quede abierta.

## 5. Instalar el servidor

Desde PowerShell en tu PC (Windows ya trae `ssh`):

```powershell
ssh root@LA_IP_DEL_VPS
```

Ya dentro del servidor:

```bash
curl -fsSL https://raw.githubusercontent.com/Rainierho/registro-shein/main/deploy/setup_vps.sh -o setup_vps.sh
bash setup_vps.sh EL_TOKEN_DEL_TUNEL
```

Tarda unos minutos. Al final avisa que faltan los secrets; es normal. Escribe `exit`.

## 6. Copiar los secrets

Desde PowerShell, en la carpeta del proyecto:

```powershell
scp .streamlit\secrets.toml root@LA_IP_DEL_VPS:/opt/lyhoshop/.streamlit/secrets.toml
ssh root@LA_IP_DEL_VPS lyho-deploy
```

Debe responder `active` en las dos apps.

## 7. Probar

- `https://lyhoshop.com` → el catálogo.
- `https://app.lyhoshop.com` → te pide tu correo y un código; luego tu registro.

En tu app → **Inventario → Catálogo para clientes**, cambia la dirección a `https://lyhoshop.com`.

## Día a día

Después de cada `git push`:

```powershell
ssh root@LA_IP_DEL_VPS lyho-deploy
```

Ver qué pasa si algo falla:

```bash
journalctl -u lyho-app -n 50      # registro
journalctl -u lyho-catalogo -n 50 # catálogo
systemctl status cloudflared      # túnel
```

## Las apps viejas de Streamlit Cloud

Déjalas una semana mientras confirmas que todo anda, y luego bórralas desde share.streamlit.io.
Leen y escriben el mismo Sheet, así que no se pierde nada; solo evita usar las dos a la vez.

## Vistas previas de productos (lyhoshop.com/p/...)

Para que al compartir el enlace de un producto aparezca su foto y precio, el túnel tiene que
mandar las direcciones `/p/...` al servicio `lyho-enlaces` (puerto 8503).

En Cloudflare → **Networking → Tunnels → lyhoshop → Routes → Add route → Published application**:

| Subdominio | Dominio | Path | Service URL |
|---|---|---|---|
| *(vacío)* | lyhoshop.com | `^/p/` | `http://localhost:8503` |

Luego **arrástrala al primer lugar** de la lista (con los puntitos de la izquierda): las rutas se
revisan de arriba hacia abajo, y si queda debajo de `lyhoshop.com`, esa se la come.

Prueba: `https://lyhoshop.com/p/` debe llevarte al catálogo. Para ver la tarjeta como la ve
WhatsApp, pega el enlace de un producto en https://www.opengraph.xyz
