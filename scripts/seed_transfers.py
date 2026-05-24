#!/usr/bin/env python3
"""
Popula custody_transfers e corporate_events no banco de producao.
Execute no servidor: python3 scripts/seed_transfers.py

Modo de operacao:
  - Apaga todos os registros existentes deste usuario e reinicia do zero.
  - Rode quantas vezes quiser — e idempotente.
"""
import sqlite3
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DB_FILE = BASE_DIR / "irsimple.db"

conn = sqlite3.connect(str(DB_FILE))
conn.row_factory = sqlite3.Row

# Garante que as tabelas existam
sys.path.insert(0, str(BASE_DIR))
import irsimple_app as core
core.init_db()

# Descobre user_id do admin (orlei)
user = conn.execute(
    "SELECT id FROM users WHERE username LIKE '%orlei%' OR role = 'admin' ORDER BY id LIMIT 1"
).fetchone()
if not user:
    print("ERRO: usuario admin nao encontrado. Verifique o banco.")
    conn.close()
    sys.exit(1)

uid = user["id"]
print(f"Usando user_id = {uid}")

# ── Limpa registros existentes para recriar do zero ──────────────────────────
conn.execute("DELETE FROM custody_transfers WHERE user_id = ?", (uid,))
conn.execute("DELETE FROM corporate_events WHERE user_id = ?", (uid,))
print("Registros anteriores removidos.")

# ── 1. Evento corporativo: Banco Modal → XP em 30/08/2023 ────────────────────
corp_events = [
    {
        "event_date": "2023-08-30",
        "event_type": "aquisicao_corretora",
        "ticker": "",
        "ticker_new": "",
        "factor": "1",
        "bonus_qty": "0",
        "bonus_cost": "0",
        "broker_from": "MODAL DTVM LTDA",
        "broker_to": "XP INVESTIMENTOS CCTVM S/A",
        "obs": "Banco Modal (MODAL DTVM LTDA / BANCO MODAL S.A.) adquirido pela XP em 30/08/2023. CNPJ 30.723.886/0001-62",
    },
]

# ── 2. Transferencias de custodia ─────────────────────────────────────────────
custody_transfers = [
    # 16/04/2025 — XP → TORO  (protocolo #20250416161448607264)
    {
        "transfer_date": "2025-04-16",
        "protocol": "#20250416161448607264",
        "broker_from": "XP INVESTIMENTOS CCTVM S/A",
        "account_from": "12832350",
        "broker_to": "TORO CTVM SA",
        "account_to": "1613418",
        "ticker": "AGRO3",
        "asset_type": "Acoes - ON",
        "quantity": "20",
        "status": "finalizado",
        "obs": "AGRO3 - BRASILAGRO - CIA BRAS DE PROP AGRICOLAS",
    },
    # 11/03/2026 — CM Capital → BTG  (protocolo #20260311193935030315)
    {
        "transfer_date": "2026-03-11",
        "protocol": "#20260311193935030315",
        "broker_from": "CM CAPITAL MARKETS CORR.",
        "account_from": "1103483",
        "broker_to": "BANCO BTG PACTUAL S/A",
        "account_to": "11941614",
        "ticker": "ASAI3",
        "asset_type": "Acoes - ON",
        "quantity": "1400",
        "status": "finalizado",
        "obs": "ASAI3 - SENDAS DISTRIBUIDORA S.A.",
    },
    {
        "transfer_date": "2026-03-11",
        "protocol": "#20260311193935030315",
        "broker_from": "CM CAPITAL MARKETS CORR.",
        "account_from": "1103483",
        "broker_to": "BANCO BTG PACTUAL S/A",
        "account_to": "11941614",
        "ticker": "BBAS3",
        "asset_type": "Acoes - ON",
        "quantity": "400",
        "status": "finalizado",
        "obs": "BBAS3 - BCO BRASIL S.A.",
    },
    {
        "transfer_date": "2026-03-11",
        "protocol": "#20260311193935030315",
        "broker_from": "CM CAPITAL MARKETS CORR.",
        "account_from": "1103483",
        "broker_to": "BANCO BTG PACTUAL S/A",
        "account_to": "11941614",
        "ticker": "BBSE3",
        "asset_type": "Acoes - ON",
        "quantity": "195",
        "status": "finalizado",
        "obs": "BBSE3 - BB SEGURIDADE PARTICIPACOES S.A.",
    },
    {
        "transfer_date": "2026-03-11",
        "protocol": "#20260311193935030315",
        "broker_from": "CM CAPITAL MARKETS CORR.",
        "account_from": "1103483",
        "broker_to": "BANCO BTG PACTUAL S/A",
        "account_to": "11941614",
        "ticker": "LREN3",
        "asset_type": "Acoes - ON",
        "quantity": "500",
        "status": "finalizado",
        "obs": "LREN3 - LOJAS RENNER S.A.",
    },
    {
        "transfer_date": "2026-03-11",
        "protocol": "#20260311193935030315",
        "broker_from": "CM CAPITAL MARKETS CORR.",
        "account_from": "1103483",
        "broker_to": "BANCO BTG PACTUAL S/A",
        "account_to": "11941614",
        "ticker": "BBDC4",
        "asset_type": "Acoes - PN",
        "quantity": "600",
        "status": "finalizado",
        "obs": "BBDC4 - BCO BRADESCO S.A.",
    },
    {
        "transfer_date": "2026-03-11",
        "protocol": "#20260311193935030315",
        "broker_from": "CM CAPITAL MARKETS CORR.",
        "account_from": "1103483",
        "broker_to": "BANCO BTG PACTUAL S/A",
        "account_to": "11941614",
        "ticker": "CMIG4",
        "asset_type": "Acoes - PN",
        "quantity": "600",
        "status": "finalizado",
        "obs": "CMIG4 - CIA ENERGETICA DE MINAS GERAIS",
    },
    {
        "transfer_date": "2026-03-11",
        "protocol": "#20260311193935030315",
        "broker_from": "CM CAPITAL MARKETS CORR.",
        "account_from": "1103483",
        "broker_to": "BANCO BTG PACTUAL S/A",
        "account_to": "11941614",
        "ticker": "WEGE3",
        "asset_type": "Acoes - ON",
        "quantity": "200",
        "status": "finalizado",
        "obs": "WEGE3 - WEG S.A.",
    },
]

# ── Insere eventos corporativos ───────────────────────────────────────────────
for ev in corp_events:
    conn.execute(
        """INSERT INTO corporate_events
           (user_id, event_date, event_type, ticker, ticker_new, factor,
            bonus_qty, bonus_cost, broker_from, broker_to, obs)
           VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
        (uid, ev["event_date"], ev["event_type"], ev["ticker"], ev["ticker_new"],
         ev["factor"], ev["bonus_qty"], ev["bonus_cost"],
         ev["broker_from"], ev["broker_to"], ev["obs"]),
    )
    print(f"  [OK] Evento corporativo: {ev['event_date']} {ev['event_type']} {ev['broker_from']} -> {ev['broker_to']}")

# ── Insere transferencias de custodia ─────────────────────────────────────────
for tr in custody_transfers:
    conn.execute(
        """INSERT INTO custody_transfers
           (user_id, transfer_date, protocol, broker_from, account_from,
            broker_to, account_to, ticker, asset_type, quantity, status, obs)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
        (uid, tr["transfer_date"], tr["protocol"],
         tr["broker_from"], tr["account_from"],
         tr["broker_to"], tr["account_to"],
         tr["ticker"], tr["asset_type"], tr["quantity"],
         tr["status"], tr["obs"]),
    )
    print(f"  [OK] Transferencia: {tr['transfer_date']} {tr['ticker']} {tr['quantity']} {tr['broker_from']} -> {tr['broker_to']}")

conn.commit()
conn.close()
print(f"\nConcluido: {len(corp_events)} eventos corporativos, {len(custody_transfers)} transferencias inseridas.")
