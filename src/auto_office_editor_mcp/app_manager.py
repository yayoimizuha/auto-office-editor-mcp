"""
Office アプリケーション（Excel / Word / PowerPoint）のライフサイクル管理。

MCPサーバー起動時は COM 用の STA ワーカースレッドだけを用意し、
Office アプリ本体は必要になった時点で個別に起動する。
サーバー停止時は起動済みアプリだけを安全に終了する。
ファイルの open / close もここで担う。

## STAワーカースレッドについて

COM の STA (Single-Threaded Apartment) モデルでは、COM オブジェクトは
作成したスレッドからしかアクセスできない。FastMCP は asyncio ベースで
動作しており、ツール関数は asyncio イベントループとは別のスレッドから
呼ばれる場合がある。そのため、全ての COM 操作を専用の STA ワーカースレッドに
委譲する設計にしている。

呼び出し側は run_in_sta(callable) を使って COM 操作を STA スレッドに投げる。
"""

from __future__ import annotations

import atexit
import logging
import queue
import shutil
import sys
import threading
import tempfile
from pathlib import Path
from typing import Any, Callable, Literal

import pythoncom
import win32com.client

logger = logging.getLogger(__name__)

AppName = Literal["excel", "word", "powerpoint"]

# COM ProgID
_PROG_ID: dict[AppName, str] = {
    "excel": "Excel.Application",
    "word": "Word.Application",
    "powerpoint": "PowerPoint.Application",
}

# 拡張子 → アプリ名
_EXT_TO_APP: dict[str, AppName] = {
    ".xlsx": "excel",
    ".xlsm": "excel",
    ".xlsb": "excel",
    ".xls": "excel",
    ".docx": "word",
    ".doc": "word",
    ".dotx": "word",
    ".pptx": "powerpoint",
    ".ppt": "powerpoint",
    ".potx": "powerpoint",
}

# STAワーカースレッドへの停止シグナル
_STOP_SENTINEL = object()


def _rebuild_win32com_gen_py_cache() -> None:
    """
    win32com の生成キャッシュ(gen_py)を毎回作り直せるよう削除する。

    起動時に毎回クリーンな状態から型ライブラリを再生成させる。
    """
    candidate_dirs = {
        Path(tempfile.gettempdir()) / "gen_py",
        Path.home() / "AppData" / "Local" / "Temp" / "gen_py",
    }

    try:
        gen_path = getattr(win32com, "__gen_path__", None)
        if gen_path:
            candidate_dirs.add(Path(str(gen_path)))
    except Exception:
        pass

    for path in candidate_dirs:
        try:
            if path.exists():
                shutil.rmtree(path, ignore_errors=True)
                logger.warning("AppManager: removed broken win32com cache: %s", path)
        except Exception as e:
            logger.warning(
                "AppManager: failed to remove win32com cache %s: %s", path, e
            )

    version_dir = (
        Path(tempfile.gettempdir())
        / "gen_py"
        / f"{sys.version_info.major}.{sys.version_info.minor}"
    )
    try:
        version_dir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        logger.warning(
            "AppManager: failed to recreate win32com cache directory %s: %s",
            version_dir,
            e,
        )


class _STAWorker:
    """
    COM 操作専用の STA ワーカースレッド。

    全ての COM 操作をこのスレッドに委譲することで、
    「別スレッドからの COM アクセス」による RPC_E_WRONG_THREAD を防ぐ。

    使い方:
        worker = _STAWorker()
        worker.start()
        result = worker.run(lambda: some_com_call())
        worker.stop()
    """

    def __init__(self) -> None:
        self._request_queue: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None

    def start(self) -> None:
        """STAワーカースレッドを開始する。"""
        self._thread = threading.Thread(
            target=self._loop, name="STA-COM-Worker", daemon=True
        )
        self._thread.start()

    def stop(self) -> None:
        """STAワーカースレッドを停止する。"""
        if self._thread and self._thread.is_alive():
            self._request_queue.put(_STOP_SENTINEL)
            self._thread.join(timeout=30)

    def is_running(self) -> bool:
        """STA ワーカースレッドが実行中なら True を返す。"""
        return self._thread is not None and self._thread.is_alive()

    def run(self, func: Callable[[], Any]) -> Any:
        """
        func を STAワーカースレッドで実行し、結果を返す。

        func が例外を raise した場合は、そのまま呼び出し元に再 raise する。
        メインスレッドをブロックする（同期呼び出し）。
        """
        if not self.is_running():
            raise RuntimeError("STA ワーカースレッドが起動していません。")

        result_event = threading.Event()
        result_holder: list = [None, None]  # [result, exception]

        def task():
            try:
                result_holder[0] = func()
            except Exception as e:
                result_holder[1] = e
            finally:
                result_event.set()

        self._request_queue.put(task)
        result_event.wait()

        if result_holder[1] is not None:
            raise result_holder[1]
        return result_holder[0]

    def _loop(self) -> None:
        """STAワーカースレッドのメインループ。"""
        pythoncom.CoInitialize()
        try:
            while True:
                item = self._request_queue.get()
                if item is _STOP_SENTINEL:
                    break
                try:
                    item()
                except Exception as e:
                    logger.error("STAWorker: unhandled exception in task: %s", e)
        finally:
            pythoncom.CoUninitialize()


class AppManager:
    """
    Excel / Word / PowerPoint の Application オブジェクトを保持するシングルトン。

    使い方:
        manager = AppManager()
        manager.start()           # MCPサーバー起動時（STA ワーカーのみ開始）
        manager.xl_app            # Excel.Application（STAスレッド内でのみ直接アクセス可）
        manager.open_file(path)   # ファイルを開く
        manager.close_file(path)  # ファイルを閉じる
        manager.shutdown()        # MCPサーバー停止時（atexit でも自動登録）
    """

    def __init__(self, visible: bool = True) -> None:
        self.xl_app: Any = None
        self.wd_app: Any = None
        self.ppt_app: Any = None
        self._started = False
        self._visible = visible
        # MCP 経由で開いたファイルのフルパス（小文字正規化）を追跡する
        self._mcp_opened_files: set[str] = set()
        # 専用 STA ワーカースレッド
        self._worker = _STAWorker()

    # ------------------------------------------------------------------
    # 起動 / 停止
    # ------------------------------------------------------------------

    def start(self, visible: bool | None = None) -> None:
        """
        STAワーカースレッドを起動する。

        Args:
            visible: True の場合は Office アプリを表示、False なら非表示。
                     None の場合は既存設定（コンストラクタや set_visibility）を使用。
        """
        if self._started:
            return

        if visible is not None:
            self._visible = visible

        self._worker.start()
        _rebuild_win32com_gen_py_cache()

        self._started = True
        atexit.register(self.shutdown)
        logger.info("AppManager: STA worker started.")

    def shutdown(self, uninitialize_com: bool = True) -> None:
        """MCP 経由で開いたファイルを保存して閉じ、アプリを終了する。

        MCP 管理外のファイル（ユーザーが手動で開いたもの）が残っている場合は、
        そのアプリは Quit() せず開いたままにする。

        Args:
            uninitialize_com: 後方互換のために残しているが、現在は _STAWorker が
                CoUninitialize() を担うため、このフラグは無視される。
        """
        if not self._started:
            return

        def _do_shutdown():
            for attr, label in [
                ("xl_app", "Excel"),
                ("wd_app", "Word"),
                ("ppt_app", "PowerPoint"),
            ]:
                app = getattr(self, attr)
                self._save_and_close_all(app, label)

            # MCP 管理外のファイルが残っているアプリは Quit() しない
            for attr, label in [
                ("xl_app", "Excel"),
                ("wd_app", "Word"),
                ("ppt_app", "PowerPoint"),
            ]:
                app = getattr(self, attr)
                if not self._has_non_mcp_files(app):
                    self._quit_app(app, label)
                    setattr(self, attr, None)
                else:
                    logger.info(
                        "AppManager: %s has user-opened files; skipping Quit().", label
                    )

        try:
            self._worker.run(_do_shutdown)
        except Exception as e:
            logger.warning("AppManager: shutdown encountered an error: %s", e)
        finally:
            self._worker.stop()
        self._started = False
        logger.info("AppManager: shutdown complete.")

    # ------------------------------------------------------------------
    # STA ワーカー経由の公開インターフェース
    # ------------------------------------------------------------------

    def set_visibility(self, visible: bool) -> None:
        """
        次回 start 時に設定する Office アプリの表示状態を更新する。

        start() 済みのアプリの表示状態は変わらないため、起動前に呼び出すこと。
        """
        self._visible = visible

    def run_in_sta(self, func: Callable[[], Any]) -> Any:
        """
        任意の callable を STAワーカースレッドで実行して結果を返す。

        executor.py や tools/*.py から COM 操作を行う際はこのメソッドを使う。
        func が例外を raise した場合は呼び出し元に再 raise する。
        """
        self._require_started()
        return self._worker.run(func)

    # ------------------------------------------------------------------
    # ファイル操作（内部実装は全て STAスレッドで実行されることを前提）
    # ------------------------------------------------------------------

    def open_file(self, filepath: str) -> dict:
        """
        拡張子からアプリを自動判定してファイルを開く。

        同名ファイルの扱い：
        - 同一パスのファイルが既に開いている → 再オープンせず already_open=True で返す
        - 同名だが別パスのファイルが既に開いている → ValueError を raise
          （Office の仕様上、同名ファイルを別パスから同時に開くことは不可能なため）

        Returns:
            {
                "app": "excel",
                "name": "hoge.xlsx",
                "path": "C:/actual/path/hoge.xlsx",
                "already_open": false
            }
        """
        self._require_started()
        return self._worker.run(lambda: self._open_file_impl(filepath))

    def close_file(self, filepath: str, save: bool = True) -> dict:
        """
        指定ファイルを閉じる。

        MCP 経由で開いたファイルのみ閉じることができる。
        それ以外のファイル（ユーザーが手動で開いたものなど）は PermissionError を raise する。

        Args:
            filepath: ファイルパス（ファイル名だけでも可）
            save: True の場合は上書き保存してから閉じる
        Returns:
            {"closed": "hoge.xlsx", "saved": True}
        """
        self._require_started()
        return self._worker.run(lambda: self._close_file_impl(filepath, save))

    def get_open_files(self) -> dict:
        """
        現在開いているファイルの一覧を返す。

        Returns:
            {
                "excel": ["hoge.xlsx", ...],
                "word": ["report.docx", ...],
                "powerpoint": ["slide.pptx", ...]
            }
        """
        self._require_started()
        return self._worker.run(self._get_open_files_impl)

    # ------------------------------------------------------------------
    # 内部実装（STAスレッドで呼ばれることを前提）
    # ------------------------------------------------------------------

    def _open_file_impl(self, filepath: str) -> dict:
        path = Path(filepath).resolve()

        app_name = _EXT_TO_APP.get(path.suffix.lower())
        if app_name is None:
            raise ValueError(f"未対応の拡張子です: {path.suffix}")

        self._ensure_app(app_name)

        name = path.name

        # 同名ファイルが開いているか確認し、同一パスか別パスかを区別する
        actual_path = self._get_full_name(app_name, name)
        if actual_path is not None:
            if Path(actual_path).resolve() == path:
                # 同一ファイルが既に開いている → MCP 管理下に記録してそのまま返す
                self._mcp_opened_files.add(str(path).lower())
                logger.info("AppManager: '%s' is already open.", name)
                return {
                    "app": app_name,
                    "name": name,
                    "path": actual_path,
                    "already_open": True,
                }
            else:
                # 同名だが別パスのファイルが開いている → Office の制約でオープン不可
                raise ValueError(
                    f"'{name}' と同名の別ファイルが既に開いています。\n"
                    f"  開いているファイル: {actual_path}\n"
                    f"  開こうとしたファイル: {path}\n"
                    f"Office の仕様上、同名ファイルを同時に開くことはできません。"
                )

        # 新規オープン
        if not path.exists():
            raise FileNotFoundError(f"ファイルが見つかりません: {path}")

        str_path = str(path)
        app = self._ensure_app(app_name)
        if app_name == "excel":
            app.Workbooks.Open(str_path)
        elif app_name == "word":
            # オープン前に保護されたビューで開かれているファイルを把握しておく。
            # インターネットからダウンロードしたファイルは保護ビューで開かれることがあり、
            # その場合 Documents コレクションではなく ProtectedViewWindows に入る。
            pv_before: set[str] = set()
            try:
                for i in range(1, app.ProtectedViewWindows.Count + 1):
                    try:
                        pv_before.add(
                            str(
                                Path(app.ProtectedViewWindows(i).SourcePath).resolve()
                            ).lower()
                        )
                    except Exception:
                        pass
            except Exception:
                pass

            # ConfirmConversions=False: 変換ダイアログを抑制
            # AddToRecentFiles=False:   最近使ったファイルへの追加ダイアログを抑制
            # ReadOnly=False:           読み取り専用ダイアログを抑制
            app.Documents.Open(
                str_path,
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
            )

            # 保護されたビューで開かれた場合は編集モードに切り替える
            # （Open() 後に ProtectedViewWindows に追加されていれば保護ビューで開かれた）
            try:
                for i in range(1, app.ProtectedViewWindows.Count + 1):
                    try:
                        pv_win = app.ProtectedViewWindows(i)
                        pv_path = str(Path(pv_win.SourcePath).resolve()).lower()
                        if pv_path == str(path).lower() and pv_path not in pv_before:
                            logger.info(
                                "AppManager: '%s' opened in Protected View; "
                                "switching to edit mode.",
                                name,
                            )
                            pv_win.Edit()
                            break
                    except Exception:
                        pass
            except Exception:
                pass
        elif app_name == "powerpoint":
            app.Presentations.Open(str_path, WithWindow=self._visible)

        # MCP 経由で開いたファイルとして記録する
        self._mcp_opened_files.add(str_path.lower())
        logger.info("AppManager: opened '%s' in %s.", name, app_name)
        return {"app": app_name, "name": name, "path": str_path, "already_open": False}

    def _close_file_impl(self, filepath: str, save: bool) -> dict:
        name = Path(filepath).name
        app_name = _EXT_TO_APP.get(Path(filepath).suffix.lower())
        if app_name is None:
            raise ValueError(f"未対応の拡張子です: {Path(filepath).suffix}")

        self._get_app(app_name)  # 起動確認

        obj = self._find_file_object(app_name, name)
        if obj is None:
            raise FileNotFoundError(f"開いていないファイルです: {name}")
        full_name = obj.FullName.lower()
        if full_name not in self._mcp_opened_files:
            raise PermissionError(
                f"'{name}' は MCP 経由で開いたファイルではないため閉じられません。"
            )

        if app_name == "excel":
            obj.Close(SaveChanges=save)
        else:
            # Word / PowerPoint: Save() → Close()（引数なし）
            if save:
                obj.Save()
            if app_name == "word":
                obj.Close(SaveChanges=False)
            else:
                obj.Close()

        self._mcp_opened_files.discard(full_name)
        logger.info("AppManager: closed '%s' (saved=%s).", name, save)
        return {"closed": name, "saved": save}

    def _get_open_files_impl(self) -> dict:
        result: dict[str, list[str]] = {"excel": [], "word": [], "powerpoint": []}

        if self.xl_app:
            try:
                for wb in self.xl_app.Workbooks:
                    result["excel"].append(wb.Name)
            except Exception:
                pass

        if self.wd_app:
            try:
                for doc in self.wd_app.Documents:
                    result["word"].append(doc.Name)
            except Exception:
                pass

        if self.ppt_app:
            try:
                for pres in self.ppt_app.Presentations:
                    result["powerpoint"].append(pres.Name)
            except Exception:
                pass

        return result

    # ------------------------------------------------------------------
    # 内部ヘルパー（STAスレッドで呼ばれることを前提）
    # ------------------------------------------------------------------

    def _launch(self, app_name: AppName) -> Any:
        """アプリを起動し、表示状態にしてダイアログを無効化して返す。"""
        app = win32com.client.gencache.EnsureDispatch(_PROG_ID[app_name])  # type: ignore[attr-defined]

        # 操作の様子が見えるように表示するかどうかを切り替え
        app.Visible = self._visible

        # ダイアログは全て握りつぶす（処理がハングするのを防ぐ）
        if app_name == "excel":
            app.DisplayAlerts = False
            app.ScreenUpdating = True  # 画面更新は有効にして操作を見えるようにする
        elif app_name == "word":
            app.DisplayAlerts = 0  # wdAlertsNone
            # msoAutomationSecurityForceDisable(3): マクロを無効化してダイアログを抑制
            try:
                app.AutomationSecurity = 3
            except Exception:
                pass
        elif app_name == "powerpoint":
            app.DisplayAlerts = False  # ppAlertsNone (0)

        logger.info("AppManager: launched %s.", _PROG_ID[app_name])
        return app

    def _require_started(self) -> None:
        if not self._started:
            raise RuntimeError(
                "AppManager が起動していません。start() を呼び出してください。"
            )

    def _ensure_app(self, app_name: AppName) -> Any:
        self._require_started()

        attr_name = {
            "excel": "xl_app",
            "word": "wd_app",
            "powerpoint": "ppt_app",
        }[app_name]
        app = getattr(self, attr_name)
        if app is None:
            app = self._launch(app_name)
            setattr(self, attr_name, app)
        return app

    def _get_app(self, app_name: AppName) -> Any:
        self._require_started()
        mapping = {
            "excel": self.xl_app,
            "word": self.wd_app,
            "powerpoint": self.ppt_app,
        }
        return mapping[app_name]

    def _find_file_object(self, app_name: AppName, name: str):
        """
        app_name のアプリで開いているファイルオブジェクトを名前で検索して返す。
        見つからなければ None を返す。
        """
        _collections: dict[AppName, Any] = {
            "excel": self.xl_app.Workbooks if self.xl_app else None,
            "word": self.wd_app.Documents if self.wd_app else None,
            "powerpoint": self.ppt_app.Presentations if self.ppt_app else None,
        }
        collection = _collections.get(app_name)
        if collection is None:
            return None
        try:
            for obj in collection:
                if obj.Name.lower() == name.lower():
                    return obj
        except Exception:
            pass
        return None

    def _get_full_name(self, app_name: AppName, name: str) -> str | None:
        """
        開いているファイルの FullName（絶対パス）を返す。
        同名のファイルが開いていなければ None を返す。
        """
        try:
            obj = self._find_file_object(app_name, name)
            if obj is not None:
                return obj.FullName
        except Exception:
            pass
        return None

    def _has_non_mcp_files(self, app) -> bool:
        """MCP 管理外のファイルがアプリに開かれていれば True を返す。"""
        if app is None:
            return False
        try:
            _collection_attrs: dict[int, Any] = {
                id(self.xl_app): app.Workbooks,
                id(self.wd_app): app.Documents,
                id(self.ppt_app): app.Presentations,
            }
            collection = _collection_attrs.get(id(app))
            if collection is None:
                return False
            return any(
                obj.FullName.lower() not in self._mcp_opened_files
                for obj in list(collection)
            )
        except Exception:
            return False

    def _save_and_close_all(self, app, label: str) -> None:
        """MCP 経由で開いたファイルのみ上書き保存してから閉じる。"""
        if app is None:
            return

        try:
            if label == "Excel":
                collection_obj, use_save_changes_arg = app.Workbooks, True
            elif label == "Word":
                collection_obj, use_save_changes_arg = app.Documents, True
            elif label == "PowerPoint":
                collection_obj, use_save_changes_arg = app.Presentations, False
            else:
                return
        except Exception as e:
            logger.warning(
                "AppManager: failed to access %s collection during shutdown: %s",
                label,
                e,
            )
            return

        try:
            for obj in list(collection_obj):
                try:
                    if obj.FullName.lower() not in self._mcp_opened_files:
                        logger.info(
                            "AppManager: skipping '%s' (not opened via MCP).", obj.Name
                        )
                        continue
                    obj.Save()
                    if use_save_changes_arg:
                        obj.Close(SaveChanges=False)
                    else:
                        obj.Close()
                    logger.info("AppManager: saved and closed '%s'.", obj.Name)
                except Exception as e:
                    logger.warning(
                        "AppManager: failed to close %s file '%s': %s",
                        label,
                        getattr(obj, "Name", "?"),
                        e,
                    )
        except Exception as e:
            logger.warning("AppManager: failed to enumerate %s files: %s", label, e)

    @staticmethod
    def _quit_app(app, label: str) -> None:
        if app is None:
            return
        try:
            app.Quit()
            logger.info("AppManager: %s quit.", label)
        except Exception as e:
            logger.warning("AppManager: failed to quit %s: %s", label, e)


# モジュールレベルのシングルトン
manager = AppManager()
