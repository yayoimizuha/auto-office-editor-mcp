"""
ファイル操作ツール。

MCP ツールとして公開する関数を定義する。
実装は AppManager に委譲する。
"""

from __future__ import annotations

import json

from ..app_manager import manager


def open_file(filepath: str) -> str:
    """
    Office ファイルを開く。

    拡張子から自動的にアプリ（Excel / Word / PowerPoint）を判定してファイルを開く。
    既に開いている場合は何もせずにファイル情報を返す。

    Args:
        filepath: 開くファイルの絶対パス（例: "C:/Users/me/report.xlsx"）

    Returns:
        JSON 文字列: {"app": "excel", "name": "report.xlsx", "path": "C:/..."}
    """
    result = manager.open_file(filepath)
    return json.dumps(result, ensure_ascii=False)


def close_file(filepath: str, save: bool = True) -> str:
    """
    開いている Office ファイルを閉じる。

    Args:
        filepath: 閉じるファイルのパスまたはファイル名（例: "report.xlsx"）
        save: True の場合は上書き保存してから閉じる（デフォルト: True）

    Returns:
        JSON 文字列: {"closed": "report.xlsx", "saved": true}
    """
    result = manager.close_file(filepath, save=save)
    return json.dumps(result, ensure_ascii=False)


def get_open_files() -> str:
    """
    現在開いている Office ファイルの一覧を返す。

    Returns:
        JSON 文字列:
        {
            "excel": ["hoge.xlsx", "fuga.xlsx"],
            "word": ["report.docx"],
            "powerpoint": ["slide.pptx"]
        }
    """
    result = manager.get_open_files()
    return json.dumps(result, ensure_ascii=False)
