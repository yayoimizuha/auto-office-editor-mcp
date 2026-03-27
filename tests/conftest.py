"""
pytest 共通フィクスチャ。

integration マーク付きテストは Office アプリを起動・終了する
session スコープの app_manager フィクスチャを使う。
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

# fixtures ディレクトリの絶対パス
FIXTURES_DIR = Path(__file__).parent / "fixtures"
EXCEL_PATH = FIXTURES_DIR / "test.xlsx"
WORD_PATH = FIXTURES_DIR / "test.docx"
PPT_PATH = FIXTURES_DIR / "test.pptx"

# 編集テスト用コピーファイルの保存先ディレクトリ
EDITABLE_DIR = FIXTURES_DIR / "editable"


@pytest.fixture(scope="session")
def app_manager():
    """
    セッションスコープで AppManager を起動・終了する。

    integration テスト全体で1回だけ Office を起動し、
    全テスト終了後に shutdown する。
    また tools/info.py などが参照するモジュールレベルの manager シングルトンも
    同じインスタンスで差し替えて一致させる。
    """
    from auto_office_editor_mcp.app_manager import AppManager
    import auto_office_editor_mcp.app_manager as _am_module
    import auto_office_editor_mcp.tools.info as _info_module
    import auto_office_editor_mcp.tools.preview as _preview_module
    import auto_office_editor_mcp.executor as _executor_module

    mgr = AppManager()
    mgr.start()

    # モジュールレベルのシングルトンをこのインスタンスで差し替える
    _am_module.manager = mgr
    _info_module.manager = mgr
    _preview_module.manager = mgr
    _executor_module.manager = mgr

    yield mgr

    # テスト環境では COM スレッドと pytest timeout スレッドの競合を避けるため
    # CoUninitialize をスキップする
    mgr.shutdown(uninitialize_com=False)


@pytest.fixture(scope="function")
def xl_open(app_manager):
    """
    test.xlsx を開き、テスト終了後に保存せず閉じる。
    各テスト関数で独立した状態を保つため function スコープ。
    """
    app_manager.open_file(str(EXCEL_PATH))
    yield app_manager
    try:
        app_manager.close_file(str(EXCEL_PATH), save=False)
    except Exception:
        pass


@pytest.fixture(scope="function")
def wd_open(app_manager):
    """
    test.docx を開き、テスト終了後に保存せず閉じる。
    各テスト関数で独立した状態を保つため function スコープ。
    """
    app_manager.open_file(str(WORD_PATH))
    yield app_manager
    try:
        app_manager.close_file(str(WORD_PATH), save=False)
    except Exception:
        pass


@pytest.fixture(scope="function")
def ppt_open(app_manager):
    """
    test.pptx を開き、テスト終了後に保存せず閉じる。
    各テスト関数で独立した状態を保つため function スコープ。
    """
    app_manager.open_file(str(PPT_PATH))
    yield app_manager
    try:
        app_manager.close_file(str(PPT_PATH), save=False)
    except Exception:
        pass


def _make_editable_fixture(src_path: Path, ext: str):
    """
    編集テスト用フィクスチャを生成するファクトリ関数。

    src_path のファイルを EDITABLE_DIR/<クラス名>_<テスト名><ext> にコピーして開く。

    - function スコープ: 各テストが独立したファイルを持つ
    - テスト後に保存して閉じる → 各テストの編集結果を個別に目視確認できる
    - 次回実行時に同名ファイルがあれば上書きコピーしてクリーンな状態から開始

    ファイル名規則:
        fixtures/editable/<クラス名>_<テスト名><ext>
        例: fixtures/editable/TestAddSmileyFace_test_smiley_name.xlsx

    Yields:
        (AppManager, Path): app_manager と編集ファイルのパス
    """

    @pytest.fixture(scope="function")
    def _fixture(request, app_manager):
        EDITABLE_DIR.mkdir(exist_ok=True)

        node = request.node
        class_name = node.cls.__name__ if node.cls else "NoClass"
        test_name = node.originalname  # パラメータなしのテスト関数名
        filename = f"{class_name}_{test_name}{ext}"
        dst = EDITABLE_DIR / filename

        # 前回残ったファイルが開いていれば保存せずに閉じる
        try:
            app_manager.close_file(str(dst), save=False)
        except Exception:
            pass

        # src_path を上書きコピー（クリーンな状態から開始）
        shutil.copy(src_path, dst)
        app_manager.open_file(str(dst))
        yield app_manager, dst

        # テスト後は保存して閉じる（ファイルは fixtures/editable/ に残す → 目視確認用）
        try:
            app_manager.close_file(str(dst), save=True)
        except Exception:
            pass

    return _fixture


xl_editable = _make_editable_fixture(EXCEL_PATH, ".xlsx")
wd_editable = _make_editable_fixture(WORD_PATH, ".docx")
ppt_editable = _make_editable_fixture(PPT_PATH, ".pptx")
