from __future__ import annotations

import pytest

from auto_office_editor_mcp.app_manager import AppManager


class _BrokenApp:
    @property
    def Workbooks(self):
        raise RuntimeError("workbooks unavailable")

    @property
    def Documents(self):
        raise RuntimeError("documents unavailable")

    @property
    def Presentations(self):
        raise RuntimeError("presentations unavailable")


@pytest.mark.parametrize("label", ["Excel", "Word", "PowerPoint"])
def test_save_and_close_all_handles_collection_access_errors(label):
    mgr = AppManager()

    # COM オブジェクトが終了済み/壊れていても shutdown が落ちないことを確認する。
    mgr._save_and_close_all(_BrokenApp(), label)
