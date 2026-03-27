"""
Office アプリケーション（Excel / Word / PowerPoint）のライフサイクル管理。

MCPサーバー起動時に3アプリをバックグラウンドで起動・保持し、
サーバー停止時に全アプリを安全に終了する。
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
import threading
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

    def run(self, func: Callable[[], Any]) -> Any:
        """
        func を STAワーカースレッドで実行し、結果を返す。

        func が例外を raise した場合は、そのまま呼び出し元に再 raise する。
        メインスレッドをブロックする（同期呼び出し）。
        """
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
        manager.start()           # MCPサーバー起動時
        manager.xl_app            # Excel.Application（STAスレッド内でのみ直接アクセス可）
        manager.open_file(path)   # ファイルを開く
        manager.close_file(path)  # ファイルを閉じる
        manager.shutdown()        # MCPサーバー停止時（atexit でも自動登録）
    """

    def __init__(self) -> None:
        self.xl_app: Any = None
        self.wd_app: Any = None
        self.ppt_app: Any = None
        self._started = False
        # MCP 経由で開いたファイルのフルパス（小文字正規化）を追跡する
        self._mcp_opened_files: set[str] = set()
        # 専用 STA ワーカースレッド
        self._worker = _STAWorker()

    # ------------------------------------------------------------------
    # 起動 / 停止
    # ------------------------------------------------------------------

    def start(self) -> None:
        """STAワーカースレッドを起動し、その中で3アプリを起動する。"""
        if self._started:
            return

        self._worker.start()

        def _init():
            self.xl_app = self._launch("excel")
            self.wd_app = self._launch("word")
            self.ppt_app = self._launch("powerpoint")

        self._worker.run(_init)

        self._started = True
        atexit.register(self.shutdown)
        logger.info("AppManager: all Office apps started.")

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
            self._save_and_close_all(self.xl_app, "Excel")
            self._save_and_close_all(self.wd_app, "Word")
            self._save_and_close_all(self.ppt_app, "PowerPoint")

            # MCP 管理外のファイルが残っているアプリは Quit() しない
            if not self._has_non_mcp_files(self.xl_app):
                self._quit_app(self.xl_app, "Excel")
                self.xl_app = None
            else:
                logger.info("AppManager: Excel has user-opened files; skipping Quit().")

            if not self._has_non_mcp_files(self.wd_app):
                self._quit_app(self.wd_app, "Word")
                self.wd_app = None
            else:
                logger.info("AppManager: Word has user-opened files; skipping Quit().")

            if not self._has_non_mcp_files(self.ppt_app):
                self._quit_app(self.ppt_app, "PowerPoint")
                self.ppt_app = None
            else:
                logger.info(
                    "AppManager: PowerPoint has user-opened files; skipping Quit()."
                )

        self._worker.run(_do_shutdown)
        self._worker.stop()
        self._started = False
        logger.info("AppManager: shutdown complete.")

    # ------------------------------------------------------------------
    # STA ワーカー経由の公開インターフェース
    # ------------------------------------------------------------------

    def run_in_sta(self, func: Callable[[], Any]) -> Any:
        """
        任意の callable を STAワーカースレッドで実行して結果を返す。

        executor.py や tools/*.py から COM 操作を行う際はこのメソッドを使う。
        func が例外を raise した場合は呼び出し元に再 raise する。
        """
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
        return self._worker.run(self._get_open_files_impl)

    # ------------------------------------------------------------------
    # 内部実装（STAスレッドで呼ばれることを前提）
    # ------------------------------------------------------------------

    def _open_file_impl(self, filepath: str) -> dict:
        path = Path(filepath).resolve()

        app_name = _EXT_TO_APP.get(path.suffix.lower())
        if app_name is None:
            raise ValueError(f"未対応の拡張子です: {path.suffix}")

        self._get_app(app_name)  # 起動確認

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
        app = self._get_app(app_name)
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
            app.Presentations.Open(str_path, WithWindow=True)

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

        if app_name == "excel":
            wb = self._find_workbook(name)
            if wb is None:
                raise FileNotFoundError(f"開いていないファイルです: {name}")
            full_name = wb.FullName.lower()
            if full_name not in self._mcp_opened_files:
                raise PermissionError(
                    f"'{name}' は MCP 経由で開いたファイルではないため閉じられません。"
                )
            wb.Close(SaveChanges=save)
            self._mcp_opened_files.discard(full_name)

        elif app_name == "word":
            doc = self._find_document(name)
            if doc is None:
                raise FileNotFoundError(f"開いていないファイルです: {name}")
            full_name = doc.FullName.lower()
            if full_name not in self._mcp_opened_files:
                raise PermissionError(
                    f"'{name}' は MCP 経由で開いたファイルではないため閉じられません。"
                )
            if save:
                doc.Save()
            doc.Close(SaveChanges=False)
            self._mcp_opened_files.discard(full_name)

        elif app_name == "powerpoint":
            pres = self._find_presentation(name)
            if pres is None:
                raise FileNotFoundError(f"開いていないファイルです: {name}")
            full_name = pres.FullName.lower()
            if full_name not in self._mcp_opened_files:
                raise PermissionError(
                    f"'{name}' は MCP 経由で開いたファイルではないため閉じられません。"
                )
            if save:
                pres.Save()
            pres.Close()
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
        try:
            # gencache は win32com の動的属性なので type: ignore で抑制
            app = win32com.client.gencache.EnsureDispatch(_PROG_ID[app_name])  # type: ignore[attr-defined]
        except Exception:
            # gencache が失敗した場合は通常の Dispatch にフォールバック
            app = win32com.client.Dispatch(_PROG_ID[app_name])

        # 操作の様子が見えるように表示する
        app.Visible = True

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

    def _get_app(self, app_name: AppName) -> Any:
        mapping = {
            "excel": self.xl_app,
            "word": self.wd_app,
            "powerpoint": self.ppt_app,
        }
        app = mapping[app_name]
        if app is None:
            raise RuntimeError(
                f"AppManager が起動していません。start() を呼び出してください。"
            )
        return app

    def _find_workbook(self, name: str):
        try:
            for wb in self.xl_app.Workbooks:
                if wb.Name.lower() == name.lower():
                    return wb
        except Exception:
            pass
        return None

    def _find_document(self, name: str):
        try:
            for doc in self.wd_app.Documents:
                if doc.Name.lower() == name.lower():
                    return doc
        except Exception:
            pass
        return None

    def _find_presentation(self, name: str):
        try:
            for pres in self.ppt_app.Presentations:
                if pres.Name.lower() == name.lower():
                    return pres
        except Exception:
            pass
        return None

    def _get_full_name(self, app_name: AppName, name: str) -> str | None:
        """
        開いているファイルの FullName（絶対パス）を返す。
        同名のファイルが開いていなければ None を返す。
        """
        try:
            if app_name == "excel":
                obj = self._find_workbook(name)
            elif app_name == "word":
                obj = self._find_document(name)
            elif app_name == "powerpoint":
                obj = self._find_presentation(name)
            else:
                return None
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
            collections = []
            if app is self.xl_app:
                collections = list(app.Workbooks)
            elif app is self.wd_app:
                collections = list(app.Documents)
            elif app is self.ppt_app:
                collections = list(app.Presentations)
            return any(
                obj.FullName.lower() not in self._mcp_opened_files
                for obj in collections
            )
        except Exception:
            return False

    def _save_and_close_all(self, app, label: str) -> None:
        """MCP 経由で開いたファイルのみ上書き保存してから閉じる。"""
        if app is None:
            return
        try:
            if label == "Excel":
                for wb in list(app.Workbooks):
                    try:
                        if wb.FullName.lower() not in self._mcp_opened_files:
                            logger.info(
                                "AppManager: skipping '%s' (not opened via MCP).",
                                wb.Name,
                            )
                            continue
                        wb.Save()
                        wb.Close(SaveChanges=False)
                        logger.info("AppManager: saved and closed '%s'.", wb.Name)
                    except Exception as e:
                        logger.warning("AppManager: failed to close workbook: %s", e)
            elif label == "Word":
                for doc in list(app.Documents):
                    try:
                        if doc.FullName.lower() not in self._mcp_opened_files:
                            logger.info(
                                "AppManager: skipping '%s' (not opened via MCP).",
                                doc.Name,
                            )
                            continue
                        doc.Save()
                        doc.Close(SaveChanges=False)
                        logger.info("AppManager: saved and closed '%s'.", doc.Name)
                    except Exception as e:
                        logger.warning("AppManager: failed to close document: %s", e)
            elif label == "PowerPoint":
                for pres in list(app.Presentations):
                    try:
                        if pres.FullName.lower() not in self._mcp_opened_files:
                            logger.info(
                                "AppManager: skipping '%s' (not opened via MCP).",
                                pres.Name,
                            )
                            continue
                        pres.Save()
                        pres.Close()
                        logger.info("AppManager: saved and closed '%s'.", pres.Name)
                    except Exception as e:
                        logger.warning(
                            "AppManager: failed to close presentation: %s", e
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
