#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
IRSimple Web - interface HTML para o calculo de IRPF Bolsa.

Esta aplicacao reaproveita o motor e os parsers do irsimple_app.py, mas expõe
as funcoes principais em paginas HTML para facilitar migracao futura para VPS.
"""

from __future__ import annotations

import io
import os
import sqlite3
from collections import defaultdict
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Any

from flask import Flask, Response, flash, redirect, render_template, request, send_file, session, url_for
from werkzeug.utils import secure_filename

import irsimple_app as core


APP_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = APP_DIR / "web_uploads"
SECRET_KEY = os.environ.get("IRSIMPLE_SECRET_KEY", "dev-change-me")

app = Flask(__name__)
app.secret_key = SECRET_KEY


def current_username() -> str:
    username = session.get("username") or core.DEFAULT_USER
    return str(username).strip() or core.DEFAULT_USER


def current_user_id() -> int:
    return core.get_user_id(current_username())


def user_upload_dir() -> Path:
    folder = UPLOAD_DIR / secure_filename(current_username())
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def money(value: Any) -> Decimal:
    return core.money(value)


def q2(value: Any) -> Decimal:
    return core.q2(value)


def fmt(value: Any) -> str:
    return core.fmt_money(value)


def fmt_dec(value: Any) -> str:
    return core.fmt_decimal(value)


def db_rows(sql: str, params: tuple[Any, ...] = ()) -> list[sqlite3.Row]:
    with core.db_connect() as conn:
        return conn.execute(sql, params).fetchall()


def save_uploaded_file(field_name: str, subfolder: str = "") -> Path | None:
    file = request.files.get(field_name)
    if not file or not file.filename:
        return None
    folder = user_upload_dir() / subfolder
    folder.mkdir(parents=True, exist_ok=True)
    filename = secure_filename(file.filename)
    path = folder / filename
    file.save(path)
    return path


def active_source(file_type: str) -> str:
    row = db_rows(
        "SELECT path FROM source_files WHERE user_id = ? AND file_type = ? AND active = 1 ORDER BY id DESC LIMIT 1",
        (current_user_id(), file_type),
    )
    return str(row[0]["path"]) if row else ""


def source_files(file_type: str) -> list[str]:
    rows = db_rows(
        "SELECT path FROM source_files WHERE user_id = ? AND file_type = ? AND active = 1 ORDER BY id",
        (current_user_id(), file_type),
    )
    return [str(row["path"]) for row in rows]


def save_source(file_type: str, path: Path, replace: bool = True) -> None:
    uid = current_user_id()
    with core.db_connect() as conn:
        if replace:
            conn.execute("UPDATE source_files SET active = 0 WHERE user_id = ? AND file_type = ?", (uid, file_type))
        conn.execute(
            "INSERT OR IGNORE INTO source_files (user_id, file_type, path, active) VALUES (?, ?, ?, 1)",
            (uid, file_type, str(path)),
        )
        conn.execute(
            "UPDATE source_files SET active = 1 WHERE user_id = ? AND file_type = ? AND path = ?",
            (uid, file_type, str(path)),
        )


def save_app_config(values: dict[str, Any]) -> None:
    uid = current_user_id()
    with core.db_connect() as conn:
        for key, value in values.items():
            conn.execute(
                "INSERT INTO app_config (user_id, key, value) VALUES (?, ?, ?) ON CONFLICT(user_id, key) DO UPDATE SET value = excluded.value",
                (uid, key, str(value)),
            )


def app_config() -> dict[str, str]:
    rows = db_rows("SELECT key, value FROM app_config WHERE user_id = ?", (current_user_id(),))
    return {str(row["key"]): str(row["value"] or "") for row in rows}


def save_trades_movements(trades: list[core.Trade], movements: list[core.Movement], neg_file: Path, mov_file: Path) -> None:
    uid = current_user_id()
    with core.db_connect() as conn:
        conn.execute("DELETE FROM trades WHERE user_id = ?", (uid,))
        conn.execute("DELETE FROM movements WHERE user_id = ?", (uid,))
        conn.executemany(
            """
            INSERT INTO trades (user_id, dt, side, market, broker, code, qty, price, value, category, expiry, source_file)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    uid,
                    trade.dt.isoformat(),
                    trade.side,
                    trade.market,
                    trade.broker,
                    trade.code,
                    str(trade.qty),
                    str(trade.price),
                    str(trade.value),
                    trade.category,
                    trade.expiry.isoformat() if trade.expiry else "",
                    str(neg_file),
                )
                for trade in trades
            ],
        )
        conn.executemany(
            """
            INSERT INTO movements (user_id, dt, direction, kind, product, broker, code, qty, unit_price, value, category, source_file)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    uid,
                    mov.dt.isoformat(),
                    mov.direction,
                    mov.kind,
                    mov.product,
                    mov.broker,
                    mov.code,
                    str(mov.qty),
                    str(mov.unit_price),
                    str(mov.value),
                    mov.category,
                    str(mov_file),
                )
                for mov in movements
            ],
        )


def load_trades() -> list[core.Trade]:
    rows = db_rows(
        """
        SELECT dt, side, market, broker, code, qty, price, value, category, expiry
        FROM trades
        WHERE user_id = ?
        ORDER BY dt, id
        """,
        (current_user_id(),),
    )
    trades: list[core.Trade] = []
    for row in rows:
        code = core.normalize_ticker(row["code"])
        trades.append(
            core.Trade(
                dt=core.parse_date(row["dt"]) or date.today(),
                side=row["side"] or "",
                market=row["market"] or "",
                broker=row["broker"] or "",
                code=code,
                qty=money(row["qty"]),
                price=money(row["price"]),
                value=q2(row["value"]),
                category=core.refine_asset_category(code, row["category"], row["market"] or ""),
                expiry=core.parse_date(row["expiry"]),
            )
        )
    return trades


def load_movements() -> list[core.Movement]:
    rows = db_rows(
        """
        SELECT dt, direction, kind, product, broker, code, qty, unit_price, value, category
        FROM movements
        WHERE user_id = ?
        ORDER BY dt, id
        """,
        (current_user_id(),),
    )
    movements: list[core.Movement] = []
    for row in rows:
        code = core.normalize_ticker(row["code"])
        movements.append(
            core.Movement(
                dt=core.parse_date(row["dt"]) or date.today(),
                direction=row["direction"] or "",
                kind=row["kind"] or "",
                product=row["product"] or "",
                broker=row["broker"] or "",
                code=code,
                qty=money(row["qty"]),
                unit_price=money(row["unit_price"]),
                value=q2(row["value"]),
                category=core.refine_asset_category(code, row["category"], product=row["product"] or ""),
            )
        )
    return movements


def load_positions() -> list[core.Position]:
    rows = db_rows(
        "SELECT ativo, quantidade, custo, categoria, corretora FROM manual_positions WHERE user_id = ? ORDER BY id",
        (current_user_id(),),
    )
    result: list[core.Position] = []
    for row in rows:
        code = core.normalize_ticker(row["ativo"])
        if code:
            result.append(
                core.Position(
                    code=code,
                    qty=money(row["quantidade"]),
                    cost=q2(row["custo"]),
                    category=(row["categoria"] or core.classify_asset(code)),
                    broker=row["corretora"] or "",
                )
            )
    return result


def load_events() -> list[dict[str, Any]]:
    rows = db_rows(
        "SELECT data, tipo, ativo, quantidade, valor, categoria, observacao FROM manual_events WHERE user_id = ? ORDER BY id",
        (current_user_id(),),
    )
    return [dict(row) for row in rows]


def save_manual_rows(table: str, rows: list[dict[str, str]]) -> None:
    uid = current_user_id()
    with core.db_connect() as conn:
        if table == "positions":
            conn.execute("DELETE FROM manual_positions WHERE user_id = ?", (uid,))
            for row in rows:
                if not core.normalize_ticker(row.get("ativo", "")):
                    continue
                conn.execute(
                    "INSERT INTO manual_positions (user_id, ativo, quantidade, custo, categoria, corretora) VALUES (?, ?, ?, ?, ?, ?)",
                    (uid, row.get("ativo", ""), row.get("quantidade", ""), row.get("custo", ""), row.get("categoria", ""), row.get("corretora", "")),
                )
        elif table == "events":
            conn.execute("DELETE FROM manual_events WHERE user_id = ?", (uid,))
            for row in rows:
                if not row.get("tipo") and not row.get("ativo"):
                    continue
                conn.execute(
                    """
                    INSERT INTO manual_events (user_id, data, tipo, ativo, quantidade, valor, categoria, observacao)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        uid,
                        row.get("data", ""),
                        row.get("tipo", ""),
                        row.get("ativo", ""),
                        row.get("quantidade", ""),
                        row.get("valor", ""),
                        row.get("categoria", ""),
                        row.get("observacao", ""),
                    ),
                )


def load_brokerage_taxes() -> list[core.BrokerageNoteTax]:
    rows = db_rows(
        "SELECT trade_date, note_number, source_file, irrf_common, irrf_daytrade, irrf_fii FROM brokerage_note_taxes WHERE user_id = ?",
        (current_user_id(),),
    )
    result: list[core.BrokerageNoteTax] = []
    for row in rows:
        dt = core.parse_date(row["trade_date"])
        if dt:
            result.append(
                core.BrokerageNoteTax(
                    dt=dt,
                    irrf_common=money(row["irrf_common"]),
                    irrf_daytrade=money(row["irrf_daytrade"]),
                    irrf_fii=money(row["irrf_fii"]),
                    note_number=str(row["note_number"] or ""),
                    source_file=str(row["source_file"] or ""),
                )
            )
    return result


def save_brokerage_notes(notes: list[dict[str, Any]]) -> tuple[int, int, int]:
    uid = current_user_id()
    inserted = updated = skipped = 0
    with core.db_connect() as conn:
        existing_rows = conn.execute(
            "SELECT id, note_number, trade_date, source_file, page FROM brokerage_note_taxes WHERE user_id = ?",
            (uid,),
        ).fetchall()
        existing = {
            (str(row["note_number"] or ""), str(row["trade_date"] or ""), str(row["source_file"] or ""), str(row["page"] or "")): row["id"]
            for row in existing_rows
        }
        seen: set[tuple[str, str, str, str]] = set()
        for note in notes:
            key = (str(note["note_number"]), note["trade_date"].isoformat(), str(note["source_file"]), str(note["page"]))
            if key in seen:
                skipped += 1
                continue
            seen.add(key)
            values = (
                note["broker"],
                note["source_file"],
                note["page"],
                str(note["irrf_common"]),
                str(note["irrf_daytrade"]),
                str(note["irrf_fii"]),
                str(note["irrf_base"]),
                str(note["buy_normal"]),
                str(note["sell_normal"]),
                str(note["buy_fii"]),
                str(note["sell_fii"]),
                str(note["buy_options"]),
                str(note["sell_options"]),
            )
            if key in existing:
                conn.execute(
                    """
                    UPDATE brokerage_note_taxes
                    SET broker = ?, source_file = ?, page = ?, irrf_common = ?, irrf_daytrade = ?, irrf_fii = ?, irrf_base = ?,
                        buy_normal = ?, sell_normal = ?, buy_fii = ?, sell_fii = ?, buy_options = ?, sell_options = ?
                    WHERE id = ? AND user_id = ?
                    """,
                    values + (existing[key], uid),
                )
                updated += 1
            else:
                conn.execute(
                    """
                    INSERT INTO brokerage_note_taxes
                    (user_id, note_number, trade_date, broker, source_file, page, irrf_common, irrf_daytrade, irrf_fii, irrf_base,
                     buy_normal, sell_normal, buy_fii, sell_fii, buy_options, sell_options)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (uid, note["note_number"], note["trade_date"].isoformat(), *values),
                )
                inserted += 1
    return inserted, updated, skipped


def asset_cnpj_map() -> dict[str, str]:
    rows = db_rows("SELECT ticker, cnpj FROM asset_cnpjs WHERE user_id = ?", (current_user_id(),))
    return {core.normalize_ticker(row["ticker"]): row["cnpj"] for row in rows}


def events_for_year(events: list[dict[str, Any]], year: int, start_year: int) -> list[dict[str, Any]]:
    selected = []
    for event in events:
        event_date = core.parse_date(event.get("data"))
        if event_date is None and year == start_year:
            selected.append(event)
        elif event_date is not None and event_date.year == year:
            selected.append(event)
    return selected


def carry_positions(result: core.CalculationResult) -> list[core.Position]:
    return [
        core.Position(
            code=pos.code,
            qty=pos.qty,
            cost=q2(pos.cost),
            previous_qty=pos.qty,
            previous_cost=q2(pos.cost),
            category=pos.category,
            broker=pos.broker,
        )
        for pos in result.positions.values()
        if pos.qty > 0
    ]


def carry_losses(result: core.CalculationResult) -> dict[str, Decimal]:
    last = result.monthly[-1]
    return {"normal": last.normal_loss_after, "daytrade": last.daytrade_loss_after, "fii": last.fii_loss_after}


def loss_effect_date(base_date: date | None) -> date | None:
    if base_date is None:
        return None
    if base_date.day == 1:
        return base_date
    year = base_date.year + 1 if base_date.month == 12 else base_date.year
    month = 1 if base_date.month == 12 else base_date.month + 1
    return date(year, month, 1)


def calculate_all() -> dict[int, core.CalculationResult]:
    cfg = app_config()
    start_year = int(cfg.get("start_year") or date.today().year - 1)
    end_year = int(cfg.get("end_year") or start_year)
    trades = load_trades()
    movements = load_movements()
    category_map = core.infer_categories_from_movements(movements)
    for trade in trades:
        if trade.code in category_map and trade.category == "normal":
            trade.category = category_map[trade.code]
    positions = load_positions()
    events = load_events()
    losses_input = {
        "normal": abs(money(cfg.get("loss_normal"))),
        "daytrade": abs(money(cfg.get("loss_daytrade"))),
        "fii": abs(money(cfg.get("loss_fii"))),
    }
    base_date = core.parse_date(cfg.get("loss_start_date"))
    loss_start = loss_effect_date(base_date)
    manual_losses_applied = False
    current_positions = positions
    current_losses = {"normal": Decimal("0"), "daytrade": Decimal("0"), "fii": Decimal("0")}
    results: dict[int, core.CalculationResult] = {}
    taxes = load_brokerage_taxes()
    cnpjs = asset_cnpj_map()
    for year in range(start_year, end_year + 1):
        initial_losses = dict(current_losses)
        dated_losses: dict[str, Decimal] = {}
        dated_loss_start: date | None = None
        if not manual_losses_applied:
            if loss_start is None or loss_start < date(year, 1, 1):
                initial_losses = dict(losses_input)
                manual_losses_applied = True
            elif loss_start.year == year:
                dated_losses = dict(losses_input)
                dated_loss_start = loss_start
                manual_losses_applied = True
        engine = core.IRSimpleEngine(year, cnpjs)
        result = engine.calculate(
            trades,
            movements,
            current_positions,
            initial_losses,
            events_for_year(events, year, start_year),
            dated_losses,
            dated_loss_start,
            taxes,
        )
        results[year] = result
        current_positions = carry_positions(result)
        current_losses = carry_losses(result)
    return results


def selected_year(results: dict[int, core.CalculationResult]) -> int:
    cfg = app_config()
    year = int(cfg.get("selected_year") or max(results or {date.today().year - 1: None}))
    return year if year in results else max(results)


def result_summary(result: core.CalculationResult) -> dict[str, Decimal]:
    return {
        "normal": q2(sum((m.normal_result for m in result.monthly), Decimal("0"))),
        "daytrade": q2(sum((m.daytrade_result for m in result.monthly), Decimal("0"))),
        "fii": q2(sum((m.fii_result for m in result.monthly), Decimal("0"))),
        "opcoes": q2(sum((m.options_result for m in result.monthly), Decimal("0"))),
        "futuro": q2(sum((m.future_result for m in result.monthly), Decimal("0"))),
        "imposto": q2(sum((m.tax_payable for m in result.monthly), Decimal("0"))),
    }


def annual_rows(result: core.CalculationResult) -> dict[str, list[list[Any]]]:
    cnpjs = asset_cnpj_map()
    bens = []
    for code, pos in sorted(result.positions.items()):
        if pos.qty <= 0 or pos.category in {"opcoes", "futuro"}:
            continue
        discr = f"{code} - Quantidade: {fmt_dec(pos.qty)} - Preco medio: R$ {fmt(pos.avg_price)}"
        bens.append([core.annual_asset_group(pos), cnpjs.get(core.normalize_ticker(code), ""), discr, fmt_dec(pos.previous_qty), fmt(pos.previous_cost), fmt_dec(pos.qty), fmt(pos.cost)])
    return {
        "bens": bens,
        "isentos": [[row["codigo"], split_cnpj(row["descricao"])[0], split_cnpj(row["descricao"])[1], fmt(row["valor"])] for row in result.exempt_income],
        "sujeitos": [[row["codigo"], split_cnpj(row["descricao"])[0], split_cnpj(row["descricao"])[1], fmt(row["valor"])] for row in result.taxable_income],
        "dividas": [[row["codigo"], "", row["descricao"], fmt(row["situacao_anterior"]), fmt(row["situacao_atual"]), fmt(row["valor_pago"])] for row in result.debts],
    }


def split_cnpj(description: str) -> tuple[str, str]:
    match = core.CNPJ_RE.search(str(description or ""))
    if not match:
        return "", str(description or "")
    cnpj = match.group(0)
    clean = str(description).replace(f"CNPJ {cnpj}", "").strip(" -")
    return cnpj, clean


def monthly_rows(result: core.CalculationResult) -> list[dict[str, Any]]:
    rows = []
    for m in result.monthly:
        rows.append(
            {
                "month": m.month,
                "mes": core.MONTHS[m.month - 1],
                "normal": m.normal_result,
                "daytrade": m.daytrade_result,
                "fii": m.fii_result,
                "opcoes": m.options_result,
                "futuro": m.future_result,
                "base_normal": m.normal_base,
                "base_dt": m.daytrade_base,
                "base_fii": m.fii_base,
                "imposto": m.tax_payable,
                "irrf_comum": m.irrf_common_month,
                "irrf_dt": m.irrf_daytrade_month,
                "irrf_fii": m.irrf_fii_month,
            }
        )
    return rows


def brokerage_summary(year: int) -> dict[int, defaultdict[str, Decimal]]:
    rows = db_rows(
        """
        SELECT trade_date, irrf_common, irrf_daytrade, irrf_fii, buy_normal, sell_normal, buy_fii, sell_fii, buy_options, sell_options
        FROM brokerage_note_taxes
        WHERE user_id = ?
        """,
        (current_user_id(),),
    )
    summary: dict[int, defaultdict[str, Decimal]] = defaultdict(lambda: defaultdict(Decimal))
    for row in rows:
        dt = core.parse_date(row["trade_date"])
        if not dt or dt.year != year:
            continue
        bucket = summary[dt.month]
        for key in ["irrf_common", "irrf_daytrade", "irrf_fii", "buy_normal", "sell_normal", "buy_fii", "sell_fii", "buy_options", "sell_options"]:
            bucket[key] += money(row[key])
    return summary


def portfolio_value(positions: dict[str, core.Position]) -> Decimal:
    return q2(sum((pos.cost for pos in positions.values() if pos.qty > 0 and pos.category not in {"opcoes", "futuro"}), Decimal("0")))


def wealth_rows(results: dict[int, core.CalculationResult], mode: str = "anual") -> list[dict[str, Any]]:
    rows = []
    if mode == "mensal":
        for year, result in sorted(results.items()):
            last_value = Decimal("0")
            for m in result.monthly:
                positions = result.monthly_positions.get(m.month, {})
                if positions:
                    last_value = portfolio_value(positions)
                rows.append(
                    {
                        "periodo": f"{year}-{m.month:02d}",
                        "patrimonio": last_value,
                        "acoes": m.normal_result,
                        "fundos": m.fii_result,
                        "outros": q2(m.daytrade_result + m.options_result + m.future_result),
                        "total": q2(m.normal_result + m.daytrade_result + m.fii_result + m.options_result + m.future_result),
                        "imposto": m.tax_payable,
                    }
                )
    else:
        for year, result in sorted(results.items()):
            summary = result_summary(result)
            rows.append(
                {
                    "periodo": str(year),
                    "patrimonio": portfolio_value(result.positions),
                    "acoes": summary["normal"],
                    "fundos": summary["fii"],
                    "outros": q2(summary["daytrade"] + summary["opcoes"] + summary["futuro"]),
                    "total": q2(summary["normal"] + summary["daytrade"] + summary["fii"] + summary["opcoes"] + summary["futuro"]),
                    "imposto": summary["imposto"],
                }
            )
    return rows


def add_benchmarks(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if not rows:
        return rows
    periods = [str(row["periodo"]) for row in rows]
    parsed = [core.parse_month_key(period) for period in periods]
    parsed = [item for item in parsed if item is not None]
    if not parsed:
        return rows
    start_year, start_month = min(parsed)
    end_year, end_month = max(parsed)
    start = date(start_year, start_month, 1)
    end = date(end_year, end_month, 28)
    try:
        selic = core.fetch_bcb_monthly_percent(4390, start, end)
        poupanca = core.fetch_bcb_monthly_percent(195, start, end)
        ibov = core.fetch_ibovespa_monthly_close(start, end)
    except Exception as exc:
        flash(f"Nao foi possivel atualizar indices: {exc}", "warning")
        return rows
    base = next((money(row["patrimonio"]) for row in rows if money(row["patrimonio"]) > 0), Decimal("100"))
    keys = core.iter_month_keys(start_year, start_month, end_year, end_month)
    first_ibov = next((ibov[k] for k in keys if ibov.get(k, 0) > 0), Decimal("0"))
    selic_value = base
    poupanca_value = base
    series = {"selic": {}, "poupanca": {}, "ibovespa": {}}
    for key in keys:
        if key in selic:
            selic_value = q2(selic_value * (Decimal("1") + selic[key] / Decimal("100")))
        if key in poupanca:
            poupanca_value = q2(poupanca_value * (Decimal("1") + poupanca[key] / Decimal("100")))
        series["selic"][key] = selic_value
        series["poupanca"][key] = poupanca_value
        if first_ibov and ibov.get(key, 0) > 0:
            series["ibovespa"][key] = q2(base * ibov[key] / first_ibov)
    for row in rows:
        parsed_key = core.parse_month_key(str(row["periodo"]))
        if not parsed_key:
            continue
        key = f"{parsed_key[0]}-{parsed_key[1]:02d}"
        row["selic"] = series["selic"].get(key, Decimal("0"))
        row["poupanca"] = series["poupanca"].get(key, Decimal("0"))
        row["ibovespa"] = series["ibovespa"].get(key, Decimal("0"))
    return rows


def require_results() -> dict[int, core.CalculationResult]:
    results = calculate_all()
    if not results:
        flash("Importe as planilhas e configure o periodo antes de calcular.", "warning")
    return results


@app.context_processor
def inject_helpers() -> dict[str, Any]:
    return {
        "fmt": fmt,
        "fmt_dec": fmt_dec,
        "months": core.MONTHS,
        "username": current_username(),
    }


@app.route("/", methods=["GET", "POST"])
def index() -> str | Response:
    if request.method == "POST":
        username = request.form.get("username", "").strip() or core.DEFAULT_USER
        session["username"] = username
        core.get_user_id(username)
        flash(f"Usuario ativo: {username}", "success")
        return redirect(url_for("dashboard"))
    return render_template("login.html")


@app.route("/dashboard")
def dashboard() -> str:
    cfg = app_config()
    trade_count = db_rows("SELECT count(*) AS n FROM trades WHERE user_id = ?", (current_user_id(),))[0]["n"]
    movement_count = db_rows("SELECT count(*) AS n FROM movements WHERE user_id = ?", (current_user_id(),))[0]["n"]
    note_count = db_rows("SELECT count(*) AS n FROM brokerage_note_taxes WHERE user_id = ?", (current_user_id(),))[0]["n"]
    return render_template(
        "dashboard.html",
        cfg=cfg,
        trade_count=trade_count,
        movement_count=movement_count,
        note_count=note_count,
        neg_file=active_source("negociacao"),
        mov_file=active_source("movimentacao"),
        consolidated=source_files("consolidado"),
    )


@app.route("/config", methods=["POST"])
def update_config() -> Response:
    save_app_config(
        {
            "name": request.form.get("name", ""),
            "start_year": request.form.get("start_year", ""),
            "end_year": request.form.get("end_year", ""),
            "selected_year": request.form.get("selected_year", ""),
            "loss_start_date": request.form.get("loss_start_date", ""),
            "loss_normal": request.form.get("loss_normal", "0,00"),
            "loss_daytrade": request.form.get("loss_daytrade", "0,00"),
            "loss_fii": request.form.get("loss_fii", "0,00"),
        }
    )
    flash("Configuracao salva.", "success")
    return redirect(request.referrer or url_for("dashboard"))


@app.route("/importar", methods=["POST"])
def importar() -> Response:
    neg = save_uploaded_file("negociacao", "b3")
    mov = save_uploaded_file("movimentacao", "b3")
    neg_path = neg or Path(active_source("negociacao"))
    mov_path = mov or Path(active_source("movimentacao"))
    if not neg_path.exists() or not mov_path.exists():
        flash("Envie as planilhas de negociacao e movimentacao ou mantenha arquivos ja cadastrados.", "danger")
        return redirect(url_for("dashboard"))
    trades = core.read_b3_negotiation(neg_path)
    movements = core.read_b3_movements(mov_path)
    category_map = core.infer_categories_from_movements(movements)
    for trade in trades:
        if trade.code in category_map and trade.category == "normal":
            trade.category = category_map[trade.code]
    save_trades_movements(trades, movements, neg_path, mov_path)
    save_source("negociacao", neg_path)
    save_source("movimentacao", mov_path)
    flash(f"Importados {len(trades)} negocios e {len(movements)} movimentacoes.", "success")
    return redirect(url_for("dashboard"))


@app.route("/notas", methods=["GET", "POST"])
def notas() -> str | Response:
    if request.method == "POST":
        files = request.files.getlist("notas")
        paths = []
        folder = user_upload_dir() / "notas"
        folder.mkdir(parents=True, exist_ok=True)
        for file in files:
            if not file.filename:
                continue
            path = folder / secure_filename(file.filename)
            file.save(path)
            paths.append(path)
        if not paths:
            flash("Selecione um ou mais PDFs de nota de corretagem.", "warning")
        else:
            notes = core.parse_brokerage_note_pdfs(paths)
            inserted, updated, skipped = save_brokerage_notes(notes)
            flash(f"Notas processadas: {len(notes)}. Novas: {inserted}. Atualizadas: {updated}. Duplicadas no lote: {skipped}.", "success")
        return redirect(url_for("notas"))
    rows = db_rows(
        "SELECT trade_date, note_number, irrf_common, irrf_daytrade, irrf_fii, irrf_base, source_file FROM brokerage_note_taxes WHERE user_id = ? ORDER BY trade_date DESC, id DESC LIMIT 300",
        (current_user_id(),),
    )
    return render_template("notas.html", rows=rows)


@app.route("/manual", methods=["GET", "POST"])
def manual() -> str | Response:
    if request.method == "POST":
        kind = request.form.get("kind")
        if kind == "position":
            row = {
                "ativo": request.form.get("ativo", ""),
                "quantidade": request.form.get("quantidade", ""),
                "custo": request.form.get("custo", ""),
                "categoria": request.form.get("categoria", ""),
                "corretora": request.form.get("corretora", ""),
            }
            rows = [dict(r) for r in db_rows("SELECT ativo, quantidade, custo, categoria, corretora FROM manual_positions WHERE user_id = ? ORDER BY id", (current_user_id(),))]
            rows.append(row)
            save_manual_rows("positions", rows)
        elif kind == "event":
            row = {
                "data": request.form.get("data", ""),
                "tipo": request.form.get("tipo", ""),
                "ativo": request.form.get("ativo", ""),
                "quantidade": request.form.get("quantidade", ""),
                "valor": request.form.get("valor", ""),
                "categoria": request.form.get("categoria", ""),
                "observacao": request.form.get("observacao", ""),
            }
            rows = [dict(r) for r in db_rows("SELECT data, tipo, ativo, quantidade, valor, categoria, observacao FROM manual_events WHERE user_id = ? ORDER BY id", (current_user_id(),))]
            rows.append(row)
            save_manual_rows("events", rows)
        flash("Registro manual salvo.", "success")
        return redirect(url_for("manual"))
    positions = db_rows("SELECT id, ativo, quantidade, custo, categoria, corretora FROM manual_positions WHERE user_id = ? ORDER BY id", (current_user_id(),))
    events = db_rows("SELECT id, data, tipo, ativo, quantidade, valor, categoria, observacao FROM manual_events WHERE user_id = ? ORDER BY id", (current_user_id(),))
    return render_template("manual.html", positions=positions, events=events)


@app.route("/manual/delete/<table>/<int:item_id>", methods=["POST"])
def delete_manual(table: str, item_id: int) -> Response:
    sql_table = "manual_positions" if table == "positions" else "manual_events"
    with core.db_connect() as conn:
        conn.execute(f"DELETE FROM {sql_table} WHERE user_id = ? AND id = ?", (current_user_id(), item_id))
    flash("Registro excluido.", "success")
    return redirect(url_for("manual"))


@app.route("/calculo")
def calculo() -> str:
    results = require_results()
    year = int(request.args.get("year") or selected_year(results))
    save_app_config({"selected_year": year})
    result = results[year]
    return render_template(
        "calculo.html",
        years=sorted(results),
        year=year,
        result=result,
        rows=monthly_rows(result),
        summary=result_summary(result),
        warnings=result.warnings,
    )


@app.route("/anual")
def anual() -> str:
    results = require_results()
    year = int(request.args.get("year") or selected_year(results))
    result = results[year]
    return render_template("anual.html", years=sorted(results), year=year, result=result, rows=annual_rows(result))


@app.route("/historico")
def historico() -> str:
    results = require_results()
    rows = []
    for year, result in sorted(results.items()):
        summary = result_summary(result)
        active_positions = [pos for pos in result.positions.values() if pos.qty > 0]
        rows.append({"year": year, "exercise": year + 1, "assets": len(active_positions), "portfolio": portfolio_value(result.positions), **summary})
    return render_template("historico.html", rows=rows)


@app.route("/carteira")
def carteira() -> str:
    results = require_results()
    year = int(request.args.get("year") or selected_year(results))
    result = results[year]
    rows = []
    for month in range(1, 13):
        loans = result.monthly_loans.get(month, {})
        for code, pos in sorted(result.monthly_positions.get(month, {}).items()):
            if pos.qty <= 0:
                continue
            rows.append(
                {
                    "year": year,
                    "month": core.MONTHS[month - 1],
                    "code": code,
                    "category": pos.category,
                    "qty": pos.qty,
                    "avg": pos.avg_price,
                    "cost": pos.cost,
                    "loan": loans.get(code, Decimal("0")),
                }
            )
    return render_template("carteira.html", years=sorted(results), year=year, rows=rows)


@app.route("/opcoes", methods=["GET", "POST"])
def opcoes() -> str | Response:
    if request.method == "POST":
        row = {
            "data": request.form.get("data", ""),
            "tipo": request.form.get("tipo", ""),
            "ativo": request.form.get("ativo", ""),
            "quantidade": request.form.get("quantidade", ""),
            "valor": request.form.get("valor", ""),
            "categoria": "opcoes",
            "observacao": request.form.get("observacao", "Lancado pela interface web."),
        }
        rows = [dict(r) for r in db_rows("SELECT data, tipo, ativo, quantidade, valor, categoria, observacao FROM manual_events WHERE user_id = ? ORDER BY id", (current_user_id(),))]
        rows.append(row)
        save_manual_rows("events", rows)
        flash("Conferencia de opcao registrada. Recalcule para refletir.", "success")
        return redirect(url_for("opcoes"))
    results = require_results()
    pending = []
    for year, result in sorted(results.items()):
        for row in result.pending_options:
            pending.append({"year": year, **row})
    events = [row for row in load_events() if "opcao" in core.normalize_header(row.get("tipo")) or "exercicio" in core.normalize_header(row.get("tipo"))]
    return render_template("opcoes.html", pending=pending, events=events)


@app.route("/patrimonio")
def patrimonio() -> str:
    results = require_results()
    mode = request.args.get("mode", "anual")
    rows = wealth_rows(results, mode)
    if request.args.get("benchmarks") == "1":
        rows = add_benchmarks(rows)
    return render_template("patrimonio.html", rows=rows, mode=mode)


@app.route("/consolidado", methods=["GET", "POST"])
def consolidado() -> str | Response:
    if request.method == "POST":
        files = request.files.getlist("relatorios")
        for file in files:
            if not file.filename:
                continue
            folder = user_upload_dir() / "consolidado"
            folder.mkdir(parents=True, exist_ok=True)
            path = folder / secure_filename(file.filename)
            file.save(path)
            save_source("consolidado", path, replace=False)
        flash("Relatorios consolidados salvos.", "success")
        return redirect(url_for("consolidado"))
    selected = request.args.get("file")
    files = source_files("consolidado")
    sheets = []
    if selected and Path(selected).exists():
        sheets = core.read_consolidated_sheets(Path(selected))
    return render_template("consolidado.html", files=files, selected=selected, sheets=sheets)


@app.route("/conferencia-b3")
def conferencia_b3() -> str:
    results = require_results()
    rows = []
    for path_text in source_files("consolidado"):
        path = Path(path_text)
        period = core.parse_consolidated_period(path)
        if period is None:
            continue
        year, month = period
        result = results.get(year)
        if not result:
            continue
        extracted = core.read_consolidated_positions(path).get("carteira", {})
        calc_positions = result.monthly_positions.get(month, {})
        assets = set(extracted) | set(calc_positions)
        for code in sorted(assets):
            b3_qty = extracted.get(code, Decimal("0"))
            calc_qty = calc_positions.get(code, core.Position(code=code)).qty
            diff = q2(calc_qty - b3_qty)
            rows.append({"year": year, "month": core.MONTHS[month - 1], "code": code, "b3": b3_qty, "calc": calc_qty, "diff": diff, "file": path.name})
    return render_template("conferencia_b3.html", rows=rows)


@app.route("/export/excel")
def export_excel() -> Response:
    if core.Workbook is None:
        flash("openpyxl nao esta instalado.", "danger")
        return redirect(url_for("calculo"))
    results = require_results()
    year = int(request.args.get("year") or selected_year(results))
    result = results[year]
    wb = core.Workbook()
    ws = wb.active
    ws.title = "Calculo mensal"
    ws.append(["Mes", "Resultado normal", "Day trade", "FII", "Opcoes", "Futuro", "Base normal", "Base DT", "Base FII", "Imposto"])
    for row in monthly_rows(result):
        ws.append([row["mes"], row["normal"], row["daytrade"], row["fii"], row["opcoes"], row["futuro"], row["base_normal"], row["base_dt"], row["base_fii"], row["imposto"]])
    for name, rows in annual_rows(result).items():
        sheet = wb.create_sheet(name[:31])
        for row in rows:
            sheet.append(row)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"IRSimple_web_{year}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/export/pdf")
def export_pdf() -> Response:
    if core.SimpleDocTemplate is None:
        flash("reportlab nao esta instalado.", "danger")
        return redirect(url_for("calculo"))
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    results = require_results()
    year = int(request.args.get("year") or selected_year(results))
    result = results[year]
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=A4, rightMargin=24, leftMargin=24, topMargin=24, bottomMargin=24)
    styles = getSampleStyleSheet()
    story = [Paragraph(f"IRSimple Web - Ano base {year} - Exercicio {year + 1}", styles["Title"]), Spacer(1, 10)]
    monthly = [["Mes", "Normal", "DT", "FII", "Opcoes", "Futuro", "Imposto"]]
    for row in monthly_rows(result):
        monthly.append([row["mes"], fmt(row["normal"]), fmt(row["daytrade"]), fmt(row["fii"]), fmt(row["opcoes"]), fmt(row["futuro"]), fmt(row["imposto"])])
    add_pdf_table(story, "Calculo mensal", monthly)
    rows = annual_rows(result)
    for title, table_rows in [("Bens e Direitos", rows["bens"]), ("Rendimentos Isentos", rows["isentos"]), ("Rendimentos Sujeitos Exclusiva", rows["sujeitos"]), ("Dividas e Onus", rows["dividas"])]:
        story.append(PageBreak())
        add_pdf_table(story, title, table_rows[:200] or [["Sem registros"]])
    doc.build(story)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"IRSimple_web_{year}.pdf", mimetype="application/pdf")


def add_pdf_table(story: list[Any], title: str, rows: list[list[Any]]) -> None:
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph, Spacer, Table, TableStyle

    styles = getSampleStyleSheet()
    story.append(Paragraph(title, styles["Heading2"]))
    story.append(Spacer(1, 6))
    table = Table([[str(cell) for cell in row] for row in rows], repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e5e7eb")),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cbd5e1")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 7),
            ]
        )
    )
    story.append(table)


if __name__ == "__main__":
    core.init_db()
    app.run(host=os.environ.get("IRSIMPLE_HOST", "127.0.0.1"), port=int(os.environ.get("IRSIMPLE_PORT", "5050")), debug=True)
