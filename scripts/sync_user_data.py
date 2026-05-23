#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Exporta/importa dados de um usuario do IRSimple sem versionar o banco SQLite."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import irsimple_app as core


USER_TABLES = [
    "app_config",
    "manual_positions",
    "manual_events",
    "asset_cnpjs",
    "trades",
    "movements",
    "brokerage_note_taxes",
]


def rows_for_user(conn: Any, table: str, user_id: int) -> list[dict[str, Any]]:
    rows = conn.execute(f"SELECT * FROM {table} WHERE user_id = ? ORDER BY id" if table != "app_config" else f"SELECT * FROM {table} WHERE user_id = ? ORDER BY key", (user_id,)).fetchall()
    result = []
    for row in rows:
        item = dict(row)
        item.pop("id", None)
        item.pop("user_id", None)
        result.append(item)
    return result


def export_user(username: str, output: Path) -> None:
    core.init_db()
    user_id = core.get_user_id(username)
    with core.db_connect() as conn:
        payload = {
            "version": 1,
            "username": username,
            "tables": {table: rows_for_user(conn, table, user_id) for table in USER_TABLES},
        }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Exportado para {output}")


def placeholders(count: int) -> str:
    return ", ".join(["?"] * count)


def import_user(username: str, source: Path) -> None:
    core.init_db()
    payload = json.loads(source.read_text(encoding="utf-8"))
    tables = payload.get("tables") or {}
    user_id = core.get_user_id(username)
    with core.db_connect() as conn:
        for table in USER_TABLES:
            conn.execute(f"DELETE FROM {table} WHERE user_id = ?", (user_id,))
        for table in USER_TABLES:
            for row in tables.get(table, []):
                item = dict(row)
                item["user_id"] = user_id
                columns = list(item)
                sql = f"INSERT INTO {table} ({', '.join(columns)}) VALUES ({placeholders(len(columns))})"
                conn.execute(sql, [item[column] for column in columns])
    print(f"Importado {source} para o usuario {username}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Sincroniza dados de usuario do IRSimple.")
    sub = parser.add_subparsers(dest="command", required=True)
    exp = sub.add_parser("export")
    exp.add_argument("--user", default=core.DEFAULT_USER)
    exp.add_argument("--output", required=True)
    imp = sub.add_parser("import")
    imp.add_argument("--user", required=True)
    imp.add_argument("--input", required=True)
    args = parser.parse_args()
    if args.command == "export":
        export_user(args.user, Path(args.output))
    else:
        import_user(args.user, Path(args.input))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
