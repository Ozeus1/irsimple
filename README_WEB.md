# IRSimple Web

Interface HTML para o `irsimple_app.py`, preparada para migracao posterior para VPS.

## Execucao local

```powershell
cd C:\Users\orlei\OneDrive\ProjPython\IRSIMPLE
python -m pip install -r requirements-web.txt
python web_app.py
```

Acesse `http://127.0.0.1:5050`.

## Preparacao para VPS Linux

```bash
sudo apt update
sudo apt install -y python3-venv python3-pip nginx git
sudo mkdir -p /opt/irsimple
sudo chown "$USER":"$USER" /opt/irsimple
git clone https://github.com/Ozeus1/irsimple.git /opt/irsimple
cd /opt/irsimple
cp .env.example .env
nano .env
```

Configure pelo menos `IRSIMPLE_SECRET_KEY` com uma chave longa.
Tambem configure `IRSIMPLE_LOGIN_EMAIL`. Se `IRSIMPLE_PASSWORD_HASH` ficar vazio, o primeiro acesso usa a senha temporaria de `IRSIMPLE_TEMP_PASSWORD` e direciona para troca de senha.

Para gerar o hash da senha:

```bash
cd /opt/irsimple
source .venv/bin/activate
read -s IRSIMPLE_PASSWORD
python -c "from werkzeug.security import generate_password_hash; import os; print(generate_password_hash(os.environ['IRSIMPLE_PASSWORD']))"
unset IRSIMPLE_PASSWORD
```

Copie o resultado para `IRSIMPLE_PASSWORD_HASH` no arquivo `.env`, ou deixe vazio e altere a senha pela tela web no primeiro acesso.

## Execucao manual na VPS

```bash
cd /opt/irsimple
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements-web.txt
gunicorn -c gunicorn.conf.py wsgi:app
```

## Servico systemd

```bash
cd /opt/irsimple
bash deploy/deploy_vps.sh
```

O servico usa `/opt/irsimple/.env` e executa `gunicorn -c gunicorn.conf.py wsgi:app`.

## Nginx e dominio

O arquivo `deploy/nginx-irsimple.conf` ja esta configurado para `ir.casatemporadaceara.cloud`.

```bash
sudo cp deploy/nginx-irsimple.conf /etc/nginx/sites-available/irsimple
sudo ln -s /etc/nginx/sites-available/irsimple /etc/nginx/sites-enabled/irsimple
sudo nginx -t
sudo systemctl reload nginx
```

Para HTTPS:

```bash
sudo apt install -y certbot python3-certbot-nginx
sudo certbot --nginx -d ir.casatemporadaceara.cloud
```

## Dados e usuarios

- O banco padrao continua sendo `irsimple.db`.
- Os registros sao associados ao usuario informado na tela inicial.
- Uploads ficam em `web_uploads/<usuario>/`.
- Para producao, proteja `irsimple.db` e `web_uploads/` com backup e permissoes restritas.
- Recomenda-se backup diario de `irsimple.db` e `web_uploads/`.
- O acesso web exige `IRSIMPLE_LOGIN_EMAIL` e `IRSIMPLE_PASSWORD_HASH` no `.env`.

## Migracao de dados locais para VPS

Os dados financeiros nao devem ser enviados ao GitHub. Exporte localmente e copie por SSH:

```powershell
python scripts/sync_user_data.py export --user local --output private_export/irsimple_user_seed.json
scp private_export/irsimple_user_seed.json root@SEU_IP:/opt/irsimple/private_seed.json
```

Na VPS:

```bash
cd /opt/irsimple
sudo systemctl stop irsimple
sudo -u www-data .venv/bin/python scripts/sync_user_data.py import --user orlei1@yahoo.com --input /opt/irsimple/private_seed.json
sudo chown www-data:www-data irsimple.db
sudo systemctl start irsimple
```

## Publicacao no GitHub

```bash
git remote add origin https://github.com/Ozeus1/irsimple.git
git push -u origin master
```

## Funcoes web incluidas

- Configuracao de anos e prejuizos iniciais.
- Importacao das planilhas B3 de negociacao e movimentacao.
- Importacao de notas de corretagem PDF para IRRF e conferencia.
- Dados manuais de posicao inicial, eventos, migracao, split, bonificacao e opcoes.
- Calculo mensal com IRRF, bases e imposto.
- Declaracao anual com bens, rendimentos, dividas e CNPJ quando disponivel.
- Historico, carteira mensal, opcoes a conferir, relatorios consolidados B3 e conferencia B3.
- Evolucao patrimonial com comparacao Ibovespa, Selic e Poupanca.
- Exportacao Excel e PDF pelo navegador.
