#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""インポートテスト"""
import sys
print(f"Python: {sys.version}")
print(f"Path: {sys.executable}")

try:
    from ryuushutsu_tool.config import DB_PATH, get_hassei, get_hakken2
    print("config OK")
    print(f"  DB_PATH: {DB_PATH}")
    print(f"  get_hassei('1'): {get_hassei('1')}")
    print(f"  get_hassei('25'): {get_hassei('25')}")
    print(f"  get_hassei('45'): {get_hassei('45')}")
    print(f"  get_hassei('B'): {get_hassei('B')}")
except Exception as e:
    print(f"config FAIL: {e}")

try:
    from ryuushutsu_tool.database import get_connection_string
    print("database OK")
    print(f"  connection_string: {get_connection_string()[:50]}...")
except Exception as e:
    print(f"database FAIL: {e}")

try:
    from ryuushutsu_tool.aggregator import add_derived_columns
    print("aggregator OK")
except Exception as e:
    print(f"aggregator FAIL: {e}")

try:
    from ryuushutsu_tool.excel_writer import generate_excel
    print("excel_writer OK")
except Exception as e:
    print(f"excel_writer FAIL: {e}")

try:
    from ryuushutsu_tool.gui import RyuushutsuApp
    print("gui OK")
except Exception as e:
    print(f"gui FAIL: {e}")

print("\n=== All imports done ===")
