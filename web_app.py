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
from functools import wraps
from pathlib import Path
from typing import Any, Callable

from flask import Flask, Response, flash, redirect, render_template, request, send_file, session, url_for
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

import irsimple_app as core


APP_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = APP_DIR / "web_uploads"
SECRET_KEY = os.environ.get("IRSIMPLE_SECRET_KEY", "dev-change-me")

app = Flask(__name__)
app.secret_key = SECRET_KEY


LOGIN_EMAIL = os.environ.get("IRSIMPLE_LOGIN_EMAIL", "orlei1@yahoo.com").strip().lower()
PASSWORD_HASH = os.environ.get("IRSIMPLE_PASSWORD_HASH", "").strip()
TEMP_PASSWORD = os.environ.get("IRSIMPLE_TEMP_PASSWORD", "irsimple@2026")


@app.template_filter("basename")
def basename_filter(value: Any) -> str:
    return Path(str(value or "")).name


@app.template_filter("date_input")
def date_input_filter(value: Any) -> str:
    parsed = core.parse_date(value)
    return parsed.isoformat() if parsed else str(value or "")


def current_username() -> str:
    username = session.get("username") or LOGIN_EMAIL or core.DEFAULT_USER
    return str(username).strip().lower() or core.DEFAULT_USER


def is_authenticated() -> bool:
    return bool(session.get("authenticated") and str(session.get("username", "")).lower() == LOGIN_EMAIL)


def login_required(view: Callable[..., Any]) -> Callable[..., Any]:
    @wraps(view)
    def wrapped(*args: Any, **kwargs: Any) -> Any:
        if not is_authenticated():
            return redirect(url_for("index", next=request.full_path if request.query_string else request.path))
        return view(*args, **kwargs)

    return wrapped


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


def password_hash_from_db() -> str:
    try:
        user_id = core.get_user_id(LOGIN_EMAIL)
        rows = db_rows("SELECT value FROM app_config WHERE user_id = ? AND key = 'login_password_hash'", (user_id,))
        return str(rows[0]["value"] or "") if rows else ""
    except Exception:
        return ""


def configured_password_hash() -> str:
    return password_hash_from_db() or PASSWORD_HASH


def verify_login_password(password: str) -> bool:
    stored_hash = configured_password_hash()
    if stored_hash:
        return check_password_hash(stored_hash, password)
    return password == TEMP_PASSWORD


def using_temporary_password() -> bool:
    return not bool(configured_password_hash())


def save_login_password(password: str) -> None:
    user_id = core.get_user_id(LOGIN_EMAIL)
    password_hash = generate_password_hash(password)
    with core.db_connect() as conn:
        conn.execute(
            "INSERT INTO app_config (user_id, key, value) VALUES (?, 'login_password_hash', ?) ON CONFLICT(user_id, key) DO UPDATE SET value = excluded.value",
            (user_id, password_hash),
        )


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


def month_daytrade_keys(trades: list[core.Trade], year: int, month: int) -> set[tuple[date, str, str, str]]:
    grouped: dict[tuple[date, str, str, str], set[str]] = defaultdict(set)
    for trade in trades:
        if trade.dt.year == year and trade.dt.month == month:
            grouped[(trade.dt, trade.code, trade.category, trade.broker)].add(trade.side)
    return {key for key, sides in grouped.items() if {"compra", "venda"}.issubset(sides)}


def daytrade_detail_rows(trades: list[core.Trade], year: int, month: int, include_future: bool | None = None) -> list[dict[str, Any]]:
    rows = []
    grouped: dict[tuple[date, str, str, str], list[core.Trade]] = defaultdict(list)
    for trade in trades:
        if trade.dt.year == year and trade.dt.month == month:
            grouped[(trade.dt, trade.code, trade.category, trade.broker)].append(trade)
    for (_dt, _code, _category, _broker), group in sorted(grouped.items()):
        buys = [item for item in group if item.side == "compra"]
        sells = [item for item in group if item.side == "venda"]
        buy_qty = sum((item.qty for item in buys), Decimal("0"))
        sell_qty = sum((item.qty for item in sells), Decimal("0"))
        qty = min(buy_qty, sell_qty)
        if qty <= 0:
            continue
        sample = group[0]
        is_future = sample.category == "futuro"
        if include_future is True and not is_future:
            continue
        if include_future is False and is_future:
            continue
        buy_avg = sum((item.value for item in buys), Decimal("0")) / buy_qty if buy_qty else Decimal("0")
        sell_avg = sum((item.value for item in sells), Decimal("0")) / sell_qty if sell_qty else Decimal("0")
        sale = q2(sell_avg * qty)
        cost = q2(buy_avg * qty)
        rows.append(
            {
                "origem": "day trade",
                "data": sample.dt.strftime("%d/%m/%Y"),
                "tipo": "compra/venda",
                "mercado": sample.market,
                "ativo": sample.code,
                "quantidade": qty,
                "preco": sell_avg,
                "preco_medio": buy_avg,
                "custo": cost,
                "venda": sale,
                "lucro": q2(sale - cost),
                "observacao": sample.broker,
            }
        )
    return rows


def include_detail_category(column: str, category: str) -> bool:
    if column in {"normal", "base_normal"}:
        return category not in {"fii", "opcoes", "futuro"}
    if column in {"fii", "base_fii"}:
        return category == "fii"
    if column == "futuro":
        return category == "futuro"
    return False


def sale_detail_rows(results: dict[int, core.CalculationResult], trades: list[core.Trade], year: int, month: int, column: str) -> list[dict[str, Any]]:
    if column in {"daytrade", "base_dt"}:
        return daytrade_detail_rows(trades, year, month, include_future=False)
    if column == "futuro":
        return daytrade_detail_rows(trades, year, month, include_future=True)
    if column not in {"normal", "base_normal", "fii", "base_fii"}:
        return []

    if (year - 1) in results:
        initial_positions = carry_positions(results[year - 1])
    else:
        initial_positions = load_positions()
    positions = {
        pos.code: core.Position(code=pos.code, qty=money(pos.qty), cost=q2(pos.cost), category=pos.category, broker=pos.broker)
        for pos in initial_positions
    }
    start_year = min(results) if results else year
    core.IRSimpleEngine(year)._apply_events_before_trades(positions, events_for_year(load_events(), year, start_year), [])
    daytrade_keys = month_daytrade_keys(trades, year, month)
    short_positions: dict[str, dict[str, Decimal]] = {}
    rows: list[dict[str, Any]] = []
    for trade in sorted([t for t in trades if t.dt.year == year and t.category != "opcoes"], key=lambda item: (item.dt, item.code, item.side)):
        if (trade.dt, trade.code, trade.category, trade.broker) in daytrade_keys:
            continue
        pos = positions.setdefault(trade.code, core.Position(code=trade.code, category=trade.category, broker=trade.broker))
        if trade.side == "compra":
            remaining_qty = trade.qty
            remaining_value = trade.value
            short = short_positions.get(trade.code)
            if short and short["qty"] > 0:
                cover_qty = min(remaining_qty, short["qty"])
                cover_cost = q2(trade.value * cover_qty / trade.qty) if trade.qty else Decimal("0")
                cover_credit = q2(short["credit"] * cover_qty / short["qty"]) if short["qty"] else Decimal("0")
                if trade.dt.month == month and include_detail_category(column, trade.category):
                    rows.append(
                        {
                            "origem": "negociacao",
                            "data": trade.dt.strftime("%d/%m/%Y"),
                            "tipo": "recompra venda descoberta",
                            "mercado": trade.market,
                            "ativo": trade.code,
                            "quantidade": cover_qty,
                            "preco": cover_credit / cover_qty if cover_qty else Decimal("0"),
                            "preco_medio": cover_cost / cover_qty if cover_qty else Decimal("0"),
                            "custo": cover_cost,
                            "venda": cover_credit,
                            "lucro": q2(cover_credit - cover_cost),
                            "observacao": trade.broker,
                        }
                    )
                short["qty"] = q2(short["qty"] - cover_qty)
                short["credit"] = q2(short["credit"] - cover_credit)
                remaining_qty = q2(remaining_qty - cover_qty)
                remaining_value = q2(remaining_value - cover_cost)
                if short["qty"] <= 0:
                    del short_positions[trade.code]
            if remaining_qty > 0:
                pos.qty += remaining_qty
                pos.cost = q2(pos.cost + remaining_value)
            continue

        sell_remaining = trade.qty
        sell_remaining_value = trade.value
        if pos.qty > 0:
            covered_qty = min(pos.qty, sell_remaining)
            sale_value = q2(trade.value * covered_qty / trade.qty) if trade.qty else Decimal("0")
            avg = pos.avg_price
            cost = q2(avg * covered_qty)
            if trade.dt.month == month and include_detail_category(column, trade.category):
                rows.append(
                    {
                        "origem": "negociacao",
                        "data": trade.dt.strftime("%d/%m/%Y"),
                        "tipo": trade.side,
                        "mercado": trade.market,
                        "ativo": trade.code,
                        "quantidade": covered_qty,
                        "preco": trade.price,
                        "preco_medio": avg,
                        "custo": cost,
                        "venda": sale_value,
                        "lucro": q2(sale_value - cost),
                        "observacao": trade.broker,
                    }
                )
            pos.qty = q2(pos.qty - covered_qty)
            pos.cost = q2(max(Decimal("0"), pos.cost - cost))
            sell_remaining = q2(sell_remaining - covered_qty)
            sell_remaining_value = q2(sell_remaining_value - sale_value)
            if pos.qty <= 0:
                pos.qty = Decimal("0")
                pos.cost = Decimal("0")
        if sell_remaining > 0:
            short = short_positions.setdefault(trade.code, {"qty": Decimal("0"), "credit": Decimal("0")})
            short["qty"] = q2(short["qty"] + sell_remaining)
            short["credit"] = q2(short["credit"] + sell_remaining_value)
            if trade.dt.month == month and include_detail_category(column, trade.category):
                rows.append(
                    {
                        "origem": "negociacao",
                        "data": trade.dt.strftime("%d/%m/%Y"),
                        "tipo": "venda descoberta aberta",
                        "mercado": trade.market,
                        "ativo": trade.code,
                        "quantidade": sell_remaining,
                        "preco": trade.price,
                        "preco_medio": "",
                        "custo": "",
                        "venda": sell_remaining_value,
                        "lucro": "",
                        "observacao": "Resultado apurado quando houver recompra ou ajuste manual.",
                    }
                )
    return rows


def option_detail_rows(trades: list[core.Trade], events: list[dict[str, Any]], year: int, month: int) -> list[dict[str, Any]]:
    rows = []
    for trade in sorted(trades, key=lambda item: (item.dt, item.code, item.side)):
        if trade.dt.year == year and trade.dt.month == month and trade.category == "opcoes":
            rows.append(
                {
                    "origem": "negociacao",
                    "data": trade.dt.strftime("%d/%m/%Y"),
                    "tipo": trade.side,
                    "mercado": trade.market,
                    "ativo": trade.code,
                    "quantidade": trade.qty,
                    "preco": trade.price,
                    "preco_medio": "",
                    "custo": trade.value if trade.side == "compra" else "",
                    "venda": trade.value if trade.side == "venda" else "",
                    "lucro": "",
                    "observacao": "Linha operacional; o lucro de opcoes usa pareamento por premio medio no motor.",
                }
            )
    for event in events:
        dt = core.parse_date(event.get("data"))
        kind = core.normalize_header(event.get("tipo"))
        if dt and dt.year == year and dt.month == month and ("opcao" in kind or "exercicio" in kind):
            rows.append(
                {
                    "origem": "evento manual",
                    "data": dt.strftime("%d/%m/%Y"),
                    "tipo": event.get("tipo", ""),
                    "mercado": "",
                    "ativo": event.get("ativo", ""),
                    "quantidade": money(event.get("quantidade")),
                    "preco": "",
                    "preco_medio": "",
                    "custo": "",
                    "venda": money(event.get("valor")),
                    "lucro": "",
                    "observacao": event.get("observacao", ""),
                }
            )
    return rows


DETAIL_LABELS = {
    "normal": "Operacoes normais",
    "daytrade": "Day trade",
    "fii": "FII/FIAGRO",
    "opcoes": "Opcoes",
    "futuro": "Mercado futuro",
    "base_normal": "Base normal",
    "base_dt": "Base day trade",
    "base_fii": "Base FII/FIAGRO",
    "irrf_comum": "IRRF comum",
    "irrf_dt": "IRRF day trade",
    "irrf_fii": "IRRF FII",
    "imposto": "Imposto",
}


def detail_summary_row(result: core.CalculationResult, month: int, column: str) -> list[dict[str, Any]]:
    tax = result.monthly[month - 1]
    if column == "base_normal":
        return [
            {
                "origem": "calculo",
                "data": "",
                "tipo": "Base normal",
                "mercado": "",
                "ativo": "",
                "quantidade": "",
                "preco": "",
                "preco_medio": "",
                "custo": "",
                "venda": tax.normal_base,
                "lucro": "",
                "observacao": (
                    f"Resultado normal R$ {fmt(tax.normal_result)}; opcoes R$ {fmt(tax.options_result)}; "
                    f"futuro comum R$ {fmt(tax.future_result)}; prejuizo anterior R$ {fmt(tax.normal_loss_before)}."
                ),
            }
        ]
    if column == "base_dt":
        future_dt = q2(tax.future_dollar_daytrade + tax.future_index_daytrade)
        return [
            {
                "origem": "calculo",
                "data": "",
                "tipo": "Base day trade",
                "mercado": "",
                "ativo": "",
                "quantidade": "",
                "preco": "",
                "preco_medio": "",
                "custo": "",
                "venda": tax.daytrade_base,
                "lucro": "",
                "observacao": (
                    f"Resultado day trade R$ {fmt(tax.daytrade_result)}; futuro day trade R$ {fmt(future_dt)}; "
                    f"prejuizo anterior R$ {fmt(tax.daytrade_loss_before)}."
                ),
            }
        ]
    if column == "base_fii":
        return [
            {
                "origem": "calculo",
                "data": "",
                "tipo": "Base FII/FIAGRO",
                "mercado": "",
                "ativo": "",
                "quantidade": "",
                "preco": "",
                "preco_medio": "",
                "custo": "",
                "venda": tax.fii_base,
                "lucro": "",
                "observacao": f"Resultado FII R$ {fmt(tax.fii_result)}; prejuizo anterior R$ {fmt(tax.fii_loss_before)}.",
            }
        ]
    if column == "imposto":
        return [
            {
                "origem": "calculo",
                "data": "",
                "tipo": "Imposto a pagar",
                "mercado": "",
                "ativo": "",
                "quantidade": "",
                "preco": "",
                "preco_medio": "",
                "custo": "",
                "venda": tax.tax_payable,
                "lucro": "",
                "observacao": (
                    f"Imposto devido R$ {fmt(tax.tax_due)}; IRRF comum no mes R$ {fmt(tax.irrf_common_month)}; "
                    f"IRRF DT no mes R$ {fmt(tax.irrf_daytrade_month)}; IRRF FII no mes R$ {fmt(tax.irrf_fii_month)}."
                ),
            }
        ]
    return []


def irrf_detail_rows(year: int, month: int, column: str) -> list[dict[str, Any]]:
    field_by_column = {
        "irrf_comum": "irrf_common",
        "irrf_dt": "irrf_daytrade",
        "irrf_fii": "irrf_fii",
    }
    field = field_by_column.get(column)
    if not field:
        return []
    rows = db_rows(
        f"""
        SELECT trade_date, note_number, source_file, {field} AS value
        FROM brokerage_note_taxes
        WHERE user_id = ? AND substr(trade_date, 1, 7) = ?
        ORDER BY trade_date, note_number, id
        """,
        (current_user_id(), f"{year}-{month:02d}"),
    )
    result = []
    for row in rows:
        value = money(row["value"])
        if value == 0:
            continue
        dt = core.parse_date(row["trade_date"])
        result.append(
            {
                "origem": "nota",
                "data": dt.strftime("%d/%m/%Y") if dt else row["trade_date"],
                "tipo": DETAIL_LABELS.get(column, column),
                "mercado": "",
                "ativo": "",
                "quantidade": "",
                "preco": "",
                "preco_medio": "",
                "custo": "",
                "venda": value,
                "lucro": "",
                "observacao": f"Nota {row['note_number'] or ''} - {Path(str(row['source_file'] or '')).name}",
            }
        )
    return result


def detail_rows(results: dict[int, core.CalculationResult], year: int, month: int, column: str) -> list[dict[str, Any]]:
    result = results[year]
    trades = load_trades()
    movements = load_movements()
    category_map = core.infer_categories_from_movements(movements)
    for trade in trades:
        if trade.code in category_map and trade.category == "normal":
            trade.category = category_map[trade.code]
    events = load_events()

    if column in {"irrf_comum", "irrf_dt", "irrf_fii"}:
        return irrf_detail_rows(year, month, column)

    rows = detail_summary_row(result, month, column)
    if column == "opcoes":
        rows.extend(option_detail_rows(trades, events_for_year(events, year, min(results) if results else year), year, month))
    elif column == "imposto":
        rows.extend(irrf_detail_rows(year, month, "irrf_comum"))
        rows.extend(irrf_detail_rows(year, month, "irrf_dt"))
        rows.extend(irrf_detail_rows(year, month, "irrf_fii"))
    else:
        rows.extend(sale_detail_rows(results, trades, year, month, column))
    return rows


def detail_totals(rows: list[dict[str, Any]]) -> dict[str, Decimal]:
    return {
        "custo": q2(sum((money(row.get("custo")) for row in rows), Decimal("0"))),
        "venda": q2(sum((money(row.get("venda")) for row in rows), Decimal("0"))),
        "lucro": q2(sum((money(row.get("lucro")) for row in rows), Decimal("0"))),
    }


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
        username = request.form.get("username", "").strip().lower()
        password = request.form.get("password", "")
        if username != LOGIN_EMAIL or not verify_login_password(password):
            flash("Usuario ou senha invalidos.", "danger")
            return render_template("login.html", login_email=username or LOGIN_EMAIL)
        session.clear()
        session["authenticated"] = True
        session["username"] = LOGIN_EMAIL
        core.get_user_id(LOGIN_EMAIL)
        if using_temporary_password():
            flash("Acesso feito com senha temporaria. Cadastre uma nova senha.", "warning")
            return redirect(url_for("change_password"))
        flash(f"Usuario autenticado: {LOGIN_EMAIL}", "success")
        next_url = request.args.get("next") or url_for("dashboard")
        return redirect(next_url if next_url.startswith("/") and not next_url.startswith("//") else url_for("dashboard"))
    return render_template("login.html", login_email=LOGIN_EMAIL)


@app.route("/logout", methods=["POST"])
def logout() -> Response:
    session.clear()
    flash("Sessao encerrada.", "success")
    return redirect(url_for("index"))


@app.route("/senha", methods=["GET", "POST"])
@login_required
def change_password() -> str | Response:
    if request.method == "POST":
        current = request.form.get("current_password", "")
        new = request.form.get("new_password", "")
        confirm = request.form.get("confirm_password", "")
        if not verify_login_password(current):
            flash("Senha atual invalida.", "danger")
        elif len(new) < 8:
            flash("A nova senha deve ter pelo menos 8 caracteres.", "warning")
        elif new != confirm:
            flash("A confirmacao da senha nao confere.", "warning")
        else:
            save_login_password(new)
            flash("Senha alterada com sucesso.", "success")
            return redirect(url_for("dashboard"))
    return render_template("change_password.html", temporary=using_temporary_password())


@app.route("/dashboard")
@login_required
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
@login_required
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
@login_required
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
@login_required
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
        "SELECT id, trade_date, note_number, irrf_common, irrf_daytrade, irrf_fii, irrf_base, source_file FROM brokerage_note_taxes WHERE user_id = ? ORDER BY trade_date DESC, id DESC LIMIT 300",
        (current_user_id(),),
    )
    return render_template("notas.html", rows=rows)


@app.route("/notas/<int:note_id>/download")
@login_required
def download_nota(note_id: int) -> Response:
    rows = db_rows(
        "SELECT source_file FROM brokerage_note_taxes WHERE user_id = ? AND id = ?",
        (current_user_id(), note_id),
    )
    if not rows:
        flash("Nota nao encontrada para o usuario atual.", "warning")
        return redirect(url_for("notas"))
    path = Path(str(rows[0]["source_file"] or ""))
    if not path.exists() or path.suffix.lower() != ".pdf":
        flash("Arquivo PDF da nota nao foi encontrado.", "warning")
        return redirect(url_for("notas"))
    return send_file(path, as_attachment=True, download_name=path.name, mimetype="application/pdf")


@app.route("/manual", methods=["GET", "POST"])
@login_required
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
@login_required
def delete_manual(table: str, item_id: int) -> Response:
    sql_table = "manual_positions" if table == "positions" else "manual_events"
    with core.db_connect() as conn:
        conn.execute(f"DELETE FROM {sql_table} WHERE user_id = ? AND id = ?", (current_user_id(), item_id))
    flash("Registro excluido.", "success")
    return redirect(url_for("manual"))


@app.route("/calculo")
@login_required
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


@app.route("/detalhe")
@login_required
def detalhe() -> str | Response:
    results = require_results()
    if not results:
        return redirect(url_for("dashboard"))
    year = int(request.args.get("year") or selected_year(results))
    month = int(request.args.get("month") or 1)
    column = request.args.get("column", "normal")
    if year not in results:
        flash("Ano sem calculo disponivel.", "warning")
        return redirect(url_for("calculo"))
    if month < 1 or month > 12:
        flash("Mes invalido para detalhamento.", "warning")
        return redirect(url_for("calculo", year=year))
    if column not in DETAIL_LABELS:
        flash("Coluna invalida para detalhamento.", "warning")
        return redirect(url_for("calculo", year=year))
    rows = detail_rows(results, year, month, column)
    return render_template(
        "detalhe.html",
        year=year,
        month=month,
        month_name=core.MONTHS[month - 1],
        column=column,
        label=DETAIL_LABELS[column],
        rows=rows,
        totals=detail_totals(rows),
    )


@app.route("/anual")
@login_required
def anual() -> str:
    results = require_results()
    year = int(request.args.get("year") or selected_year(results))
    result = results[year]
    return render_template("anual.html", years=sorted(results), year=year, result=result, rows=annual_rows(result))


@app.route("/historico")
@login_required
def historico() -> str:
    results = require_results()
    rows = []
    for year, result in sorted(results.items()):
        summary = result_summary(result)
        active_positions = [pos for pos in result.positions.values() if pos.qty > 0]
        rows.append({"year": year, "exercise": year + 1, "assets": len(active_positions), "portfolio": portfolio_value(result.positions), **summary})
    return render_template("historico.html", rows=rows)


@app.route("/carteira")
@login_required
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
@login_required
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
@login_required
def patrimonio() -> str:
    results = require_results()
    mode = request.args.get("mode", "anual")
    rows = wealth_rows(results, mode)
    if request.args.get("benchmarks") == "1":
        rows = add_benchmarks(rows)
    return render_template("patrimonio.html", rows=rows, mode=mode)


@app.route("/consolidado", methods=["GET", "POST"])
@login_required
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
@login_required
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
@login_required
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
@login_required
def export_pdf() -> Response:
    if core.SimpleDocTemplate is None:
        flash("reportlab nao esta instalado.", "danger")
        return redirect(url_for("calculo"))
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import cm
    from reportlab.platypus import PageBreak, SimpleDocTemplate

    results = require_results()
    year = int(request.args.get("year") or selected_year(results))
    result = results[year]
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(A4), leftMargin=1 * cm, rightMargin=1 * cm, topMargin=1 * cm, bottomMargin=1 * cm)
    styles = getSampleStyleSheet()
    story: list[Any] = []

    pdf_header(story, styles, "Impostos em Renda Variavel", year)
    for idx, month in enumerate(result.monthly):
        if idx:
            story.append(PageBreak())
            pdf_header(story, styles, "Impostos em Renda Variavel", year)
        pdf_variable_month(story, month)

    story.append(PageBreak())
    pdf_header(story, styles, "Fundos Imobiliarios", year)
    pdf_fii_monthly(story, result)

    annual = annual_rows(result)
    previous_year = year - 1
    for key, title, headers in [
        ("bens", "Bens e Direitos", ["Codigo", "CNPJ", "Discriminacao", f"Qtd em 31/12/{previous_year}", f"Situacao em 31/12/{previous_year}", f"Qtd em 31/12/{year}", f"Situacao em 31/12/{year}"]),
        ("isentos", "Rendimentos Isentos e Nao Tributaveis", ["Codigo", "CNPJ", "Descricao", "Valor"]),
        ("sujeitos", "Rendimentos Sujeitos a Tributacao Exclusiva", ["Codigo", "CNPJ", "Descricao", "Valor"]),
        ("dividas", "Onus e Dividas", ["Codigo", "CNPJ", "Discriminacao", f"Situacao em 31/12/{previous_year}", f"Situacao em 31/12/{year}", f"Valor pago em {year}"]),
    ]:
        story.append(PageBreak())
        pdf_header(story, styles, title, year)
        pdf_table(story, title, headers, annual[key])

    pdf_consolidated_reports(story, styles, year)
    doc.build(story)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"IRSimple_exercicio_{year + 1}_ano_base_{year}.pdf", mimetype="application/pdf")


def pdf_header(story: list[Any], styles: Any, subtitle: str, year: int) -> None:
    from reportlab.lib.units import cm
    from reportlab.platypus import Paragraph, Spacer

    cfg = app_config()
    story.append(Paragraph("DECLARACAO ANUAL DE IMPOSTO DE RENDA", styles["Title"]))
    story.append(Paragraph(f"Nome: {cfg.get('name') or '-'}", styles["Normal"]))
    story.append(Paragraph(f"{subtitle} - Ano base: {year} - Exercicio: {year + 1}", styles["Heading2"]))
    story.append(Spacer(1, 0.25 * cm))


def pdf_variable_month(story: list[Any], m: core.MonthlyTax) -> None:
    month = core.MONTHS[m.month - 1]
    future_dt = q2(m.future_dollar_daytrade + m.future_index_daytrade)
    future_common = q2(m.future_result - future_dt)
    common_result = q2(m.normal_result + m.options_result + future_common)
    daytrade_result = q2(m.daytrade_result + future_dt)
    common_tax = q2(m.normal_base * Decimal("0.15"))
    daytrade_tax = q2(m.daytrade_base * Decimal("0.20"))
    rows = [
        ["Mercado a Vista", "Operacoes Comuns", "Day Trade"],
        ["Mercado a vista - acoes", fmt(m.normal_result), fmt(m.daytrade_result)],
        ["Mercado a vista - ouro", "-", "-"],
        ["Mercado a vista - ouro at. fin. fora bolsa", "-", "-"],
        ["Mercado de Opcoes", "Operacoes Comuns", "Day Trade"],
        ["Mercado opcoes - acoes", fmt(m.options_result), "-"],
        ["Mercado opcoes - ouro", "-", "-"],
        ["Mercado opcoes - fora de bolsa", "-", "-"],
        ["Mercado opcoes - outros", "-", "-"],
        ["Mercado Futuro", "Operacoes Comuns", "Day Trade"],
        ["Mercado futuro - dolar dos EUA", fmt(m.future_dollar_common), fmt(m.future_dollar_daytrade)],
        ["Mercado futuro - indices", fmt(m.future_index_common), fmt(m.future_index_daytrade)],
        ["Mercado futuro - juros", "-", "-"],
        ["Mercado futuro - outros", fmt(q2(future_common - m.future_dollar_common - m.future_index_common)), fmt(q2(future_dt - m.future_dollar_daytrade - m.future_index_daytrade))],
        ["Mercado a Termo", "Operacoes Comuns", "Day Trade"],
        ["Mercado a termo - acoes/ouro", "-", "-"],
        ["Mercado a termo - outros", "-", "-"],
        ["Resultados", "Operacoes Comuns", "Day Trade"],
        ["RESULTADO LIQUIDO DO MES", fmt(common_result), fmt(daytrade_result)],
        ["Resultado negativo ate o mes anterior", fmt(m.normal_loss_before), fmt(m.daytrade_loss_before)],
        ["BASE DE CALCULO DO IMPOSTO", fmt(m.normal_base), fmt(m.daytrade_base)],
        ["Prejuizo a compensar", fmt(m.normal_loss_after), fmt(m.daytrade_loss_after)],
        ["Aliquota do imposto", "15%", "20%"],
        ["IMPOSTO DEVIDO", fmt(common_tax), fmt(daytrade_tax)],
        ["Consolidacao do Mes", "", ""],
        ["Total do imposto devido", fmt(common_tax), fmt(daytrade_tax)],
        ["IR fonte de Day Trade no mes", "", fmt(m.irrf_daytrade_month)],
        ["IR fonte de Day Trade nos meses anteriores", "", fmt(m.irrf_daytrade_before)],
        ["IR fonte de Day Trade a compensar", "", fmt(m.irrf_daytrade_after)],
        ["IR fonte(Lei no 11.033/2004) no mes", fmt(m.irrf_common_month), ""],
        ["IR fonte(Lei no 11.033/2004) nos meses anteriores", fmt(m.irrf_common_before), ""],
        ["IR fonte(Lei no 11.033/2004) a compensar", fmt(m.irrf_common_after), ""],
        ["Imposto a pagar", fmt(max(Decimal("0"), common_tax - m.irrf_common_before - m.irrf_common_month)), fmt(max(Decimal("0"), daytrade_tax - m.irrf_daytrade_before - m.irrf_daytrade_month))],
        ["Imposto pago (valor + imposto acumulado + multa + juros)", "0,00", "0,00"],
    ]
    pdf_table(story, f"Ganhos Liquidos ou Perdas em {month}", ["", "", ""], rows, section_rows={0, 4, 9, 14, 17, 24})


def pdf_fii_monthly(story: list[Any], result: core.CalculationResult) -> None:
    rows = [
        [
            core.MONTHS[m.month - 1],
            fmt(m.fii_result),
            fmt(m.fii_loss_before),
            fmt(m.fii_base),
            fmt(m.fii_loss_after),
            "20,00 %",
            fmt(m.fii_base * Decimal("0.20")),
            fmt(m.irrf_fii_before),
            fmt(m.irrf_fii_month),
            fmt(m.irrf_fii_after),
            fmt(max(Decimal("0"), m.fii_base * Decimal("0.20") - m.irrf_fii_before - m.irrf_fii_month)),
            "-",
        ]
        for m in result.monthly
    ]
    pdf_table(
        story,
        "Ganhos Liquidos ou Perdas",
        [
            "Mes",
            "Resultado liquido no mes",
            "Resultado negativo ate o mes anterior",
            "Base de calculo do imposto",
            "Prejuizo a compensar",
            "Aliquota do imposto",
            "Imposto devido",
            "Saldo do imposto retido nos meses anteriores",
            "Imposto retido no mes",
            "Imposto a compensar",
            "Imposto a pagar",
            "Imposto pago",
        ],
        rows,
    )


def consolidated_files_for_year(year: int) -> list[Path]:
    paths: list[Path] = []
    for file in source_files("consolidado"):
        path = Path(file)
        if not path.exists():
            continue
        period = core.parse_consolidated_period(path)
        if period and period[0] == year:
            paths.append(path)
    return sorted(paths, key=lambda path: (core.parse_consolidated_period(path) or (year, 99))[1])


def pdf_consolidated_reports(story: list[Any], styles: Any, year: int) -> None:
    from reportlab.platypus import PageBreak

    for path in consolidated_files_for_year(year):
        try:
            sheets = core.read_consolidated_sheets(path)
        except Exception as exc:
            story.append(PageBreak())
            pdf_header(story, styles, "Relatorio Consolidado B3", year)
            pdf_table(story, path.name, ["Erro"], [[f"Falha ao ler arquivo: {exc}"]])
            continue
        for sheet in sheets:
            story.append(PageBreak())
            pdf_header(story, styles, "Relatorio Consolidado B3", year)
            columns = [str(col) for col in (sheet.get("columns") or ["Sem dados"])]
            rows = sheet.get("rows") or [[""]]
            normalized_rows = [list(row[: len(columns)]) + [""] * max(0, len(columns) - len(row)) for row in rows]
            pdf_table(story, f"{path.name} - {sheet.get('name') or 'Sheet'}", columns, normalized_rows)


def pdf_table(story: list[Any], title: str, headers: list[str], rows: list[list[Any]], section_rows: set[int] | None = None) -> None:
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import Paragraph, Spacer, Table, TableStyle

    styles = getSampleStyleSheet()
    story.append(Paragraph(title, styles["Heading2"]))
    normal = styles["Normal"]
    normal.fontSize = 7
    normal.leading = 8
    source_rows = rows if rows else [["-" for _ in headers]]
    normalized_rows = [list(row[: len(headers)]) + [""] * max(0, len(headers) - len(row)) for row in source_rows]
    data = [[Paragraph(str(cell), normal) for cell in headers]] + [[Paragraph(str(cell), normal) for cell in row] for row in normalized_rows]
    table = Table(data, repeatRows=1)
    style_commands = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eaeaea")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#c9c9c9")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]
    for idx in section_rows or set():
        row = idx + 1
        style_commands.extend(
            [
                ("BACKGROUND", (0, row), (-1, row), colors.HexColor("#f1f1f1")),
                ("FONTNAME", (0, row), (-1, row), "Helvetica-Bold"),
            ]
        )
    table.setStyle(TableStyle(style_commands))
    story.append(table)
    story.append(Spacer(1, 0.2 * cm))


if __name__ == "__main__":
    core.init_db()
    app.run(host=os.environ.get("IRSIMPLE_HOST", "127.0.0.1"), port=int(os.environ.get("IRSIMPLE_PORT", "5050")), debug=True)
