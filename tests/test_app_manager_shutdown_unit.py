from __future__ import annotations

from pathlib import Path

import pytest

import auto_office_editor_mcp.app_manager as app_manager_module
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


class _FakeExcelWorkbooks:
    def __init__(self) -> None:
        self.opened: list[str] = []

    def Open(self, path: str) -> None:
        self.opened.append(path)

    def __iter__(self):
        return iter(())


class _FakeWordDocuments:
    def __init__(self) -> None:
        self.opened: list[str] = []

    def Open(self, path: str, **kwargs) -> None:
        self.opened.append(path)

    def __iter__(self):
        return iter(())


class _FakePptPresentations:
    def __init__(self) -> None:
        self.opened: list[tuple[str, bool]] = []

    def Open(self, path: str, WithWindow: bool) -> None:
        self.opened.append((path, WithWindow))

    def __iter__(self):
        return iter(())


class _FakeProtectedViewWindows:
    Count = 0

    def __call__(self, index: int):
        raise IndexError(index)


class _FakeExcelApp:
    def __init__(self) -> None:
        self.Workbooks = _FakeExcelWorkbooks()


class _FakeWordApp:
    def __init__(self) -> None:
        self.Documents = _FakeWordDocuments()
        self.ProtectedViewWindows = _FakeProtectedViewWindows()


class _FakePptApp:
    def __init__(self) -> None:
        self.Presentations = _FakePptPresentations()


@pytest.mark.parametrize("label", ["Excel", "Word", "PowerPoint"])
def test_save_and_close_all_handles_collection_access_errors(label):
    mgr = AppManager()

    # COM オブジェクトが終了済み/壊れていても shutdown が落ちないことを確認する。
    mgr._save_and_close_all(_BrokenApp(), label)


def test_start_only_starts_sta_worker(monkeypatch):
    mgr = AppManager()
    called: list[str] = []

    monkeypatch.setattr(app_manager_module.atexit, "register", lambda func: func)
    monkeypatch.setattr(mgr._worker, "start", lambda: called.append("worker"))
    monkeypatch.setattr(app_manager_module, "_rebuild_win32com_gen_py_cache", lambda: called.append("cache"))

    mgr.start()

    assert called == ["worker", "cache"]
    assert mgr.xl_app is None
    assert mgr.wd_app is None
    assert mgr.ppt_app is None


def test_open_file_launches_only_target_app(monkeypatch, tmp_path):
    mgr = AppManager()
    mgr._started = True
    monkeypatch.setattr(mgr._worker, "run", lambda func: func())

    launched: list[str] = []
    fake_excel = _FakeExcelApp()
    fake_word = _FakeWordApp()
    fake_ppt = _FakePptApp()

    def fake_launch(app_name: str):
        launched.append(app_name)
        mapping = {
            "excel": fake_excel,
            "word": fake_word,
            "powerpoint": fake_ppt,
        }
        return mapping[app_name]

    monkeypatch.setattr(mgr, "_launch", fake_launch)

    workbook_path = tmp_path / "sample.xlsx"
    workbook_path.write_text("stub")

    result = mgr.open_file(str(workbook_path))

    assert result["app"] == "excel"
    assert launched == ["excel"]
    assert fake_excel.Workbooks.opened == [str(Path(workbook_path).resolve())]
    assert mgr.wd_app is None
    assert mgr.ppt_app is None
