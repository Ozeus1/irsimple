#!/usr/bin/env python3
"""Remove o hash de senha do admin, reativando a senha temporaria do .env ou do codigo."""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import irsimple_app as core

EMAIL = os.environ.get("IRSIMPLE_LOGIN_EMAIL", "orlei1@yahoo.com").strip().lower()

core.init_db()
uid = core.get_user_id(EMAIL)
with core.db_connect() as conn:
    conn.execute("DELETE FROM app_config WHERE user_id = ? AND key = 'login_password_hash'", (uid,))

print(f"Hash de senha removido para '{EMAIL}'.")
print("Use a senha temporaria definida em IRSIMPLE_TEMP_PASSWORD (padrao: IRSimple@Reset2026).")
