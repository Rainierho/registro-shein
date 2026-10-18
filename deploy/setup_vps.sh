#!/usr/bin/env bash
# =============================================================================
# Lyho Shop · deja un VPS Ubuntu 24.04 listo con las dos apps
#   lyhoshop.com      -> catálogo  (catalogo.py, puerto 8502)
#   app.lyhoshop.com  -> registro  (app.py,      puerto 8501)
# Uso (como root):  bash setup_vps.sh [TOKEN_DEL_TUNEL]
# Se puede correr varias veces: lo que ya está hecho lo deja igual.
# =============================================================================
set -euo pipefail

REPO="https://github.com/Rainierho/registro-shein.git"
DIR="/opt/lyhoshop"
USUARIO="lyho"
TOKEN="${1:-}"

paso() { printf '\n\033[1;36m==> %s\033[0m\n' "$*"; }
aviso() { printf '\033[1;33m!!  %s\033[0m\n' "$*"; }

[ "$(id -u)" -eq 0 ] || { echo "Ejecútalo como root (o con sudo)."; exit 1; }
export DEBIAN_FRONTEND=noninteractive

paso "Zona horaria de Venezuela"
timedatectl set-timezone America/Caracas || true

paso "Actualizando el sistema (tarda unos minutos)"
apt-get update -q
apt-get -y -q upgrade
apt-get install -y -q git python3-venv python3-pip ufw fail2ban unattended-upgrades curl ca-certificates

paso "Actualizaciones de seguridad automáticas"
dpkg-reconfigure -f noninteractive unattended-upgrades || true

paso "Memoria de intercambio (por si el plan trae poca RAM)"
RAM_MB=$(awk '/MemTotal/ {print int($2/1024)}' /proc/meminfo)
if [ "$RAM_MB" -lt 3000 ] && ! swapon --show | grep -q /swapfile; then
    fallocate -l 2G /swapfile && chmod 600 /swapfile && mkswap /swapfile && swapon /swapfile
    grep -q '^/swapfile' /etc/fstab || echo '/swapfile none swap sw 0 0' >> /etc/fstab
    echo "    swap de 2 GB creada (RAM: ${RAM_MB} MB)"
else
    echo "    no hace falta (RAM: ${RAM_MB} MB)"
fi

paso "Firewall: todo cerrado menos SSH (las apps salen por el túnel, sin abrir puertos)"
ufw default deny incoming >/dev/null
ufw default allow outgoing >/dev/null
ufw allow OpenSSH >/dev/null
ufw --force enable >/dev/null
systemctl enable --now fail2ban >/dev/null 2>&1 || true

paso "Usuario de servicio '${USUARIO}' (las apps no corren como root)"
id "$USUARIO" >/dev/null 2>&1 || useradd --system --create-home --shell /usr/sbin/nologin "$USUARIO"

paso "Código desde GitHub"
if [ -d "$DIR/.git" ]; then
    sudo -u "$USUARIO" git -C "$DIR" pull --ff-only
else
    git clone "$REPO" "$DIR"
    chown -R "$USUARIO:$USUARIO" "$DIR"
fi

paso "Entorno de Python y dependencias"
[ -x "$DIR/.venv/bin/python" ] || sudo -u "$USUARIO" python3 -m venv "$DIR/.venv"
sudo -u "$USUARIO" "$DIR/.venv/bin/pip" install -q --upgrade pip
sudo -u "$USUARIO" "$DIR/.venv/bin/pip" install -q -r "$DIR/requirements.txt"

paso "Configuración de Streamlit (solo escucha en el propio servidor)"
mkdir -p "$DIR/.streamlit"
cat > "$DIR/.streamlit/config.toml" <<'TOML'
[server]
headless = true
address = "127.0.0.1"
maxUploadSize = 50

[browser]
gatherUsageStats = false
TOML
chown -R "$USUARIO:$USUARIO" "$DIR/.streamlit"

paso "Servicios: arrancan solos y se reinician si fallan"
crear_servicio() {   # nombre, archivo, puerto, descripción
    cat > "/etc/systemd/system/$1.service" <<UNIT
[Unit]
Description=Lyho Shop · $4
After=network-online.target
Wants=network-online.target

[Service]
User=$USUARIO
WorkingDirectory=$DIR
ExecStart=$DIR/.venv/bin/streamlit run $2 --server.port $3
Restart=always
RestartSec=5
Environment=PYTHONUNBUFFERED=1

[Install]
WantedBy=multi-user.target
UNIT
}
crear_servicio lyho-app      app.py      8501 "registro (app.lyhoshop.com)"
crear_servicio lyho-catalogo catalogo.py 8502 "catálogo (lyhoshop.com)"
systemctl daemon-reload
systemctl enable lyho-app lyho-catalogo >/dev/null 2>&1

paso "Comando 'lyho-deploy' para actualizar desde GitHub"
cat > /usr/local/bin/lyho-deploy <<DEPLOY
#!/usr/bin/env bash
# Baja lo último de GitHub, instala dependencias nuevas y reinicia las dos apps.
set -euo pipefail
sudo -u $USUARIO git -C $DIR pull --ff-only
sudo -u $USUARIO $DIR/.venv/bin/pip install -q -r $DIR/requirements.txt
if [ -f $DIR/.streamlit/secrets.toml ]; then
    chown $USUARIO:$USUARIO $DIR/.streamlit/secrets.toml
    chmod 600 $DIR/.streamlit/secrets.toml
fi
systemctl restart lyho-app lyho-catalogo
sleep 4
for s in lyho-app lyho-catalogo; do
    printf '%-14s %s\n' "\$s" "\$(systemctl is-active \$s)"
done
DEPLOY
chmod +x /usr/local/bin/lyho-deploy

paso "Cloudflare Tunnel (cloudflared)"
if ! command -v cloudflared >/dev/null 2>&1; then
    install -d -m 0755 /usr/share/keyrings
    curl -fsSL https://pkg.cloudflare.com/cloudflare-main.gpg -o /usr/share/keyrings/cloudflare-main.gpg
    echo 'deb [signed-by=/usr/share/keyrings/cloudflare-main.gpg] https://pkg.cloudflare.com/cloudflared any main' \
        > /etc/apt/sources.list.d/cloudflared.list
    apt-get update -q && apt-get install -y -q cloudflared
fi
if [ -n "$TOKEN" ]; then
    if systemctl is-enabled cloudflared >/dev/null 2>&1; then
        echo "    el túnel ya estaba instalado"
    else
        cloudflared service install "$TOKEN"
    fi
else
    aviso "Sin token del túnel: cuando lo tengas, corre  cloudflared service install <TOKEN>"
fi

paso "Secrets"
SECRETS="$DIR/.streamlit/secrets.toml"
if [ -f "$SECRETS" ]; then
    chown "$USUARIO:$USUARIO" "$SECRETS"; chmod 600 "$SECRETS"
    systemctl restart lyho-app lyho-catalogo
    sleep 5
    echo "    secrets encontrados; apps reiniciadas"
else
    aviso "Falta $SECRETS"
    aviso "Cópialo desde tu PC (ver la guía) y luego corre:  lyho-deploy"
fi

paso "Estado"
for s in lyho-app lyho-catalogo cloudflared; do
    printf '    %-14s %s\n' "$s" "$(systemctl is-active "$s" 2>/dev/null || true)"
done
for p in 8501 8502; do
    if curl -fs -o /dev/null "http://127.0.0.1:$p/_stcore/health"; then
        echo "    puerto $p responde"
    else
        echo "    puerto $p todavía no responde (normal si faltan los secrets)"
    fi
done
printf '\n\033[1;32mListo.\033[0m Para actualizar en el futuro: lyho-deploy\n'
