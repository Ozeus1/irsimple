# IRSimple Web

Interface HTML para o `irsimple_app.py`, preparada para migracao posterior para VPS.

## Execucao local

```powershell
cd C:\Users\orlei\OneDrive\ProjPython\IRSIMPLE
python -m pip install -r requirements-web.txt
python web_app.py
```

Acesse `http://127.0.0.1:5050`.

## Execucao em VPS Linux

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements-web.txt
export IRSIMPLE_SECRET_KEY="troque-esta-chave"
gunicorn -w 2 -b 0.0.0.0:5050 web_app:app
```

Use Nginx/Apache como proxy reverso para HTTPS.

## Dados e usuarios

- O banco padrao continua sendo `irsimple.db`.
- Os registros sao associados ao usuario informado na tela inicial.
- Uploads ficam em `web_uploads/<usuario>/`.
- Para producao, proteja `irsimple.db` e `web_uploads/` com backup e permissoes restritas.

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
