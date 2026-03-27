"""
LLM が生成した Python コードを安全に exec() で実行するモジュール。

exec_globals に以下を注入した状態でコードを実行する:
    - 標準モジュール（pd, np, os, re, json, datetime, Path, win32com, pythoncom）
    - Office アプリオブジェクト（xl_app, wd_app, ppt_app）
    - win32com 定数（const）
    - ヘルパー関数（rgb, cm_to_pt, inch_to_pt, pt_to_cm, hex_to_rgb_ole,
                     df_to_xl, xl_to_df,
                     df_to_ppt_table, ppt_table_to_df,
                     df_to_wd_table, wd_table_to_df）

コード実行成功時は stdout キャプチャ結果を返す。
例外発生時は Traceback 全文を文字列で返す（LLM のセルフヒーリング用）。
"""

from __future__ import annotations

import datetime
import io
import json
import os
import re
import sys
import traceback
from contextlib import redirect_stdout
from pathlib import Path

import numpy as np
import pandas as pd
import pythoncom
import win32com.client

from .app_manager import manager
from .helpers import (
    cm_to_pt,
    df_to_ppt_table,
    df_to_wd_table,
    df_to_xl,
    hex_to_rgb_ole,
    inch_to_pt,
    ppt_table_to_df,
    pt_to_cm,
    rgb,
    wd_table_to_df,
    xl_to_df,
)


def _build_exec_globals() -> dict:
    """
    exec() に渡すグローバル名前空間を構築して返す。

    アプリオブジェクトは AppManager から取得する。
    AppManager が未起動の場合は None になる（LLMのコードが None チェックできる）。
    """
    # win32com.client.constants は gencache が有効な場合のみ利用可能
    # gencache なし環境でも動くよう安全に取得する
    try:
        const = win32com.client.constants  # type: ignore[attr-defined]
    except AttributeError:
        const = None

    return {
        # --- 標準モジュール ---
        "pd": pd,
        "np": np,
        "os": os,
        "re": re,
        "sys": sys,
        "json": json,
        "datetime": datetime,
        "Path": Path,
        "win32com": win32com,
        "pythoncom": pythoncom,
        # --- Office アプリオブジェクト ---
        "xl_app": manager.xl_app,
        "wd_app": manager.wd_app,
        "ppt_app": manager.ppt_app,
        # --- win32com 定数 ---
        "const": const,
        # --- ヘルパー関数 ---
        "rgb": rgb,
        "hex_to_rgb_ole": hex_to_rgb_ole,
        "cm_to_pt": cm_to_pt,
        "inch_to_pt": inch_to_pt,
        "pt_to_cm": pt_to_cm,
        "df_to_xl": df_to_xl,
        "xl_to_df": xl_to_df,
        "df_to_ppt_table": df_to_ppt_table,
        "ppt_table_to_df": ppt_table_to_df,
        "df_to_wd_table": df_to_wd_table,
        "wd_table_to_df": wd_table_to_df,
    }


def _execute_script_in_sta(code: str) -> str:
    """
    STAワーカースレッド内で実際に exec() を行う内部関数。
    AppManager._worker.run() から呼ばれる。
    """
    exec_globals = _build_exec_globals()
    stdout_buf = io.StringIO()

    try:
        with redirect_stdout(stdout_buf):
            exec(code, exec_globals)  # noqa: S102
    except Exception:
        tb = traceback.format_exc()
        return f"スクリプト実行中にエラーが発生しました。コードを修正して再試行してください:\n\n{tb}"

    output = stdout_buf.getvalue()
    if output:
        return f"実行成功\n\n--- 出力 ---\n{output}"
    return "実行成功"


def execute_script(code: str) -> str:
    """
    LLM が生成した win32com 操作コードを実行する。

    全ての COM 操作を STA ワーカースレッドに委譲することで、
    RPC_E_WRONG_THREAD エラーを回避する。

    - exec_globals に Office オブジェクトとヘルパー関数を注入済みの状態で実行する。
    - コード内の print() 出力は stdout としてキャプチャし、戻り値に含める。
    - 例外が発生した場合は Traceback 全文を文字列で返す（LLM のセルフヒーリング用）。

    Args:
        code: 実行する Python コード文字列

    Returns:
        成功時: "実行成功\n" + キャプチャした stdout（print があれば）
        失敗時: Traceback 全文の文字列
    """
    return manager.run_in_sta(lambda: _execute_script_in_sta(code))
