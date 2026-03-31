"""
テスト共通ユーティリティ。

テストコードからのみ参照する定数・ヘルパー関数を定義する。
conftest.py はフィクスチャ専用とし、ここには含めない。
"""

from __future__ import annotations

from auto_office_editor_mcp.helpers import rgb

# ------------------------------------------------------------------
# OLE カラー定数
# ------------------------------------------------------------------
# Office の COM API は色を BGR 順の整数（OLE カラー）で扱う。
# helpers.rgb() で変換した値をここに定数として集約し、
# テストコード内での手書き計算式の重複を排除する。

OLE_YELLOW: int = rgb(255, 255, 0)
OLE_RED: int = rgb(255, 0, 0)
OLE_BLUE: int = rgb(0, 0, 255)
OLE_BLACK: int = rgb(0, 0, 0)
OLE_WHITE: int = rgb(255, 255, 255)
