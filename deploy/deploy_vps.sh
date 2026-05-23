#!/usr/bin/env bash
set -euo pipefail

APP_DIR="${APP_DIR:-/opt/irsimple}"

cd "$APP_DIR"
python3 -m venv .venv
. .venv/bin/activate
pip install --upgrade pip
pip install -r requirements-web.txt

if [ ! -f .env ]; then
  cp .env.example .env
  echo "Edite $APP_DIR/.env e configure IRSIMPLE_SECRET_KEY antes de iniciar o servico."
fi

sudo mkdir -p "$APP_DIR/web_uploads"
sudo touch "$APP_DIR/irsimple.db"
sudo chown -R www-data:www-data "$APP_DIR/irsimple.db" "$APP_DIR/web_uploads"
sudo chmod 664 "$APP_DIR/irsimple.db"
sudo chmod 775 "$APP_DIR" "$APP_DIR/web_uploads"

sudo cp deploy/irsimple.service /etc/systemd/system/irsimple.service
sudo systemctl daemon-reload
sudo systemctl enable irsimple
sudo systemctl restart irsimple
sudo systemctl status irsimple --no-pager
