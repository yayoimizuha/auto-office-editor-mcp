"""
Excel 統合テスト。

前提: tests/fixtures/test.xlsx が存在すること。

test.xlsx の構成:
  Sheet1: ヘッダー行 + データ3行（氏名/年齢/売上/入社日）
  Sheet2: 空シート
  Sheet3: x/sin(x)/cos(x) の数値データ29行 + グラフ1つ + 星型図形1つ

各テストは xl_open フィクスチャ経由でファイルを開き、
テスト終了後に保存せずに閉じる（ファイルを汚さない）。
"""

from __future__ import annotations

import json
import math

import pandas as pd
import pytest

pytestmark = pytest.mark.integration

WORKBOOK_NAME = "test.xlsx"

# Sheet1 の期待値
SHEET1_HEADERS = ["氏名", "年齢", "売上", "入社日"]
SHEET1_ROW_COUNT = 3  # ヘッダー除くデータ行数

# Sheet3 の期待値
SHEET3_HEADERS = ["x", "sin(x)", "cos(x)"]
SHEET3_DATA_ROWS = 29  # x=0〜1.4, 0.05刻み


# ===========================================================================
# ファイルオープン / クローズ
# ===========================================================================


class TestOpenClose:
    def test_open_returns_correct_app(self, xl_open):
        """open_file が excel アプリを返す"""
        from tests.conftest import EXCEL_PATH

        result = xl_open.open_file(str(EXCEL_PATH))
        assert result["app"] == "excel"

    def test_open_returns_correct_name(self, xl_open):
        """open_file がファイル名を返す"""
        from tests.conftest import EXCEL_PATH

        result = xl_open.open_file(str(EXCEL_PATH))
        assert result["name"] == WORKBOOK_NAME

    def test_open_already_open_returns_flag(self, xl_open):
        """既に開いているファイルを再度開くと already_open=True"""
        from tests.conftest import EXCEL_PATH

        result = xl_open.open_file(str(EXCEL_PATH))
        assert result["already_open"] is True

    def test_open_already_open_returns_full_path(self, xl_open):
        """already_open 時に返す path が実際のフルパス（引数パスと一致）"""
        from tests.conftest import EXCEL_PATH

        result = xl_open.open_file(str(EXCEL_PATH))
        import pathlib

        assert pathlib.Path(result["path"]).resolve() == EXCEL_PATH.resolve()

    def test_open_nonexistent_raises(self, app_manager):
        """存在しないファイルを開こうとすると FileNotFoundError"""
        with pytest.raises(FileNotFoundError):
            app_manager.open_file("C:/nonexistent/hoge.xlsx")

    def test_open_unsupported_ext_raises(self, app_manager):
        """未対応拡張子は ValueError"""
        with pytest.raises(ValueError, match="未対応"):
            app_manager.open_file("C:/hoge.txt")


# ===========================================================================
# get_open_files / シート一覧
# ===========================================================================


class TestGetOpenFiles:
    def test_excel_file_in_list(self, xl_open):
        """開いた後は get_open_files の excel リストに含まれる"""
        files = xl_open.get_open_files()
        assert WORKBOOK_NAME in files["excel"]

    def test_word_list_empty(self, xl_open):
        """Word ファイルは開いていない"""
        files = xl_open.get_open_files()
        assert files["word"] == []

    def test_powerpoint_list_empty(self, xl_open):
        """PowerPoint ファイルは開いていない"""
        files = xl_open.get_open_files()
        assert files["powerpoint"] == []


class TestGetExcelSheets:
    def test_sheet_count(self, xl_open):
        """シートが3枚あること"""
        from auto_office_editor_mcp.tools.info import get_xl_sheets_info

        result = json.loads(get_xl_sheets_info(WORKBOOK_NAME))
        assert len(result) == 3

    def test_sheet_names(self, xl_open):
        """シート名が Sheet1/Sheet2/Sheet3 であること"""
        from auto_office_editor_mcp.tools.info import get_xl_sheets_info

        result = json.loads(get_xl_sheets_info(WORKBOOK_NAME))
        names = [s["name"] for s in result]
        assert names == ["Sheet1", "Sheet2", "Sheet3"]

    def test_sheet_indices(self, xl_open):
        """インデックスが 1始まりで連番であること"""
        from auto_office_editor_mcp.tools.info import get_xl_sheets_info

        result = json.loads(get_xl_sheets_info(WORKBOOK_NAME))
        indices = [s["index"] for s in result]
        assert indices == [1, 2, 3]

    def test_all_sheets_visible(self, xl_open):
        """全シートが表示状態であること"""
        from auto_office_editor_mcp.tools.info import get_xl_sheets_info

        result = json.loads(get_xl_sheets_info(WORKBOOK_NAME))
        assert all(s["visible"] for s in result)

    def test_not_opened_workbook_returns_error(self, app_manager):
        """開いていないワークブック名はエラーを返す"""
        from auto_office_editor_mcp.tools.info import get_xl_sheets_info

        result = json.loads(get_xl_sheets_info("nonexistent.xlsx"))
        assert "error" in result


# ===========================================================================
# get_excel_range / セル範囲読み取り
# ===========================================================================


class TestGetExcelRange:
    def test_sheet1_shape(self, xl_open):
        """Sheet1 A1:D4 は 4行4列"""
        from auto_office_editor_mcp.tools.info import get_xl_range_info

        result = json.loads(get_xl_range_info(WORKBOOK_NAME, "Sheet1", "A1:D4"))
        assert result["rows"] == 4
        assert result["columns"] == 4

    def test_sheet1_headers(self, xl_open):
        """Sheet1 の1行目がヘッダー行"""
        from auto_office_editor_mcp.tools.info import get_xl_range_info

        result = json.loads(get_xl_range_info(WORKBOOK_NAME, "Sheet1", "A1:D1"))
        assert result["values"][0] == SHEET1_HEADERS

    def test_sheet1_first_name(self, xl_open):
        """Sheet1 A2 が '田中太郎'"""
        from auto_office_editor_mcp.tools.info import get_xl_range_info

        result = json.loads(get_xl_range_info(WORKBOOK_NAME, "Sheet1", "A2"))
        assert result["values"][0][0] == "田中太郎"

    def test_sheet1_age_is_numeric(self, xl_open):
        """Sheet1 B2（年齢）が数値"""
        from auto_office_editor_mcp.tools.info import get_xl_range_info

        result = json.loads(get_xl_range_info(WORKBOOK_NAME, "Sheet1", "B2"))
        assert result["values"][0][0] == 30.0

    def test_sheet1_date_is_returned(self, xl_open):
        """Sheet1 D列（入社日）が何らかの値として返る（None でない）"""
        from auto_office_editor_mcp.tools.info import get_xl_range_info

        result = json.loads(get_xl_range_info(WORKBOOK_NAME, "Sheet1", "D2:D4"))
        for row in result["values"]:
            assert row[0] is not None

    def test_sheet2_empty(self, xl_open):
        """Sheet2 A1 は空セルなので values が空リスト"""
        from auto_office_editor_mcp.tools.info import get_xl_range_info

        result = json.loads(get_xl_range_info(WORKBOOK_NAME, "Sheet2", "A1"))
        assert result["values"] == []

    def test_sheet3_headers(self, xl_open):
        """Sheet3 の1行目が x/sin(x)/cos(x)"""
        from auto_office_editor_mcp.tools.info import get_xl_range_info

        result = json.loads(get_xl_range_info(WORKBOOK_NAME, "Sheet3", "A1:C1"))
        assert result["values"][0] == SHEET3_HEADERS

    def test_sheet3_row_count(self, xl_open):
        """Sheet3 A1:C30 は 30行3列（ヘッダー1行 + データ29行）"""
        from auto_office_editor_mcp.tools.info import get_xl_range_info

        result = json.loads(get_xl_range_info(WORKBOOK_NAME, "Sheet3", "A1:C30"))
        assert result["rows"] == 30
        assert result["columns"] == 3

    def test_over_20_rows_note(self, xl_open):
        """21行以上の範囲は note フィールドに警告が入る"""
        from auto_office_editor_mcp.tools.info import get_xl_range_info

        result = json.loads(get_xl_range_info(WORKBOOK_NAME, "Sheet3", "A1:C30"))
        assert result["note"] != ""


# ===========================================================================
# xl_to_df : DataFrame 変換
# ===========================================================================


class TestXlToDf:
    def test_sheet1_shape(self, xl_open):
        """Sheet1 の DataFrame は 3行4列（ヘッダー除く）"""
        from auto_office_editor_mcp.helpers import xl_to_df

        df = xl_open.run_in_sta(
            lambda: xl_to_df(
                xl_open.xl_app.Workbooks(WORKBOOK_NAME)
                .Worksheets("Sheet1")
                .Range("A1:D4")
            )
        )
        assert df.shape == (3, 4)

    def test_sheet1_columns(self, xl_open):
        """Sheet1 の DataFrame の列名がヘッダー行と一致"""
        from auto_office_editor_mcp.helpers import xl_to_df

        df = xl_open.run_in_sta(
            lambda: xl_to_df(
                xl_open.xl_app.Workbooks(WORKBOOK_NAME)
                .Worksheets("Sheet1")
                .Range("A1:D4")
            )
        )
        assert list(df.columns) == SHEET1_HEADERS

    def test_sheet1_first_row_name(self, xl_open):
        """Sheet1 の1行目の氏名が '田中太郎'"""
        from auto_office_editor_mcp.helpers import xl_to_df

        df = xl_open.run_in_sta(
            lambda: xl_to_df(
                xl_open.xl_app.Workbooks(WORKBOOK_NAME)
                .Worksheets("Sheet1")
                .Range("A1:D4")
            )
        )
        assert df.iloc[0]["氏名"] == "田中太郎"

    def test_sheet1_age_values(self, xl_open):
        """Sheet1 の年齢列が正しい値"""
        from auto_office_editor_mcp.helpers import xl_to_df

        df = xl_open.run_in_sta(
            lambda: xl_to_df(
                xl_open.xl_app.Workbooks(WORKBOOK_NAME)
                .Worksheets("Sheet1")
                .Range("A1:D4")
            )
        )
        assert list(df["年齢"]) == [30.0, 25.0, 42.0]

    def test_sheet1_sales_values(self, xl_open):
        """Sheet1 の売上列が正しい値"""
        from auto_office_editor_mcp.helpers import xl_to_df

        df = xl_open.run_in_sta(
            lambda: xl_to_df(
                xl_open.xl_app.Workbooks(WORKBOOK_NAME)
                .Worksheets("Sheet1")
                .Range("A1:D4")
            )
        )
        assert list(df["売上"]) == [1500000.0, 980000.0, 2300000.0]

    def test_sheet2_empty_returns_empty_df(self, xl_open):
        """Sheet2 の空セルは空の DataFrame"""
        from auto_office_editor_mcp.helpers import xl_to_df

        df = xl_open.run_in_sta(
            lambda: xl_to_df(
                xl_open.xl_app.Workbooks(WORKBOOK_NAME).Worksheets("Sheet2").Range("A1")
            )
        )
        assert df.empty

    def test_sheet3_numeric_data(self, xl_open):
        """Sheet3 の数値データが正しく読める"""
        from auto_office_editor_mcp.helpers import xl_to_df

        df = xl_open.run_in_sta(
            lambda: xl_to_df(
                xl_open.xl_app.Workbooks(WORKBOOK_NAME)
                .Worksheets("Sheet3")
                .Range("A1:C30")
            )
        )
        assert df.shape == (SHEET3_DATA_ROWS, 3)
        assert list(df.columns) == SHEET3_HEADERS

    def test_sheet3_x_first_value(self, xl_open):
        """Sheet3 の x 列の最初の値が 0"""
        from auto_office_editor_mcp.helpers import xl_to_df

        df = xl_open.run_in_sta(
            lambda: xl_to_df(
                xl_open.xl_app.Workbooks(WORKBOOK_NAME)
                .Worksheets("Sheet3")
                .Range("A1:C30")
            )
        )
        assert df.iloc[0]["x"] == 0.0

    def test_sheet3_x_last_value(self, xl_open):
        """Sheet3 の x 列の最後の値が 1.4"""
        from auto_office_editor_mcp.helpers import xl_to_df

        df = xl_open.run_in_sta(
            lambda: xl_to_df(
                xl_open.xl_app.Workbooks(WORKBOOK_NAME)
                .Worksheets("Sheet3")
                .Range("A1:C30")
            )
        )
        assert abs(df.iloc[-1]["x"] - 1.4) < 0.001

    def test_sheet3_sin_values(self, xl_open):
        """Sheet3 の sin(x) 列が math.sin と近似一致"""
        from auto_office_editor_mcp.helpers import xl_to_df

        df = xl_open.run_in_sta(
            lambda: xl_to_df(
                xl_open.xl_app.Workbooks(WORKBOOK_NAME)
                .Worksheets("Sheet3")
                .Range("A1:C30")
            )
        )
        for _, row in df.iterrows():
            assert abs(row["sin(x)"] - math.sin(row["x"])) < 0.001

    def test_sheet3_cos_values(self, xl_open):
        """Sheet3 の cos(x) 列が math.cos と近似一致"""
        from auto_office_editor_mcp.helpers import xl_to_df

        df = xl_open.run_in_sta(
            lambda: xl_to_df(
                xl_open.xl_app.Workbooks(WORKBOOK_NAME)
                .Worksheets("Sheet3")
                .Range("A1:C30")
            )
        )
        for _, row in df.iterrows():
            assert abs(row["cos(x)"] - math.cos(row["x"])) < 0.001


# ===========================================================================
# df_to_xl : DataFrame 書き込み往復テスト
# ===========================================================================


class TestDfToXl:
    def test_roundtrip_shape(self, xl_open):
        """Sheet1 を df_to_xl で書き直して xl_to_df で再読み込みすると形状が一致"""
        from auto_office_editor_mcp.helpers import df_to_xl, xl_to_df

        def _do():
            ws = xl_open.xl_app.Workbooks(WORKBOOK_NAME).Worksheets("Sheet1")
            original_df = xl_to_df(ws.Range("A1:D4"))
            df_numbers = original_df[["氏名", "年齢", "売上"]].copy()
            df_to_xl(df_numbers, ws.Range("F1"))
            return xl_to_df(ws.Range("F1:H4")), df_numbers

        result_df, df_numbers = xl_open.run_in_sta(_do)
        assert result_df.shape == df_numbers.shape

    def test_roundtrip_columns(self, xl_open):
        """往復後の列名が一致"""
        from auto_office_editor_mcp.helpers import df_to_xl, xl_to_df

        def _do():
            ws = xl_open.xl_app.Workbooks(WORKBOOK_NAME).Worksheets("Sheet1")
            original_df = xl_to_df(ws.Range("A1:C4"))
            df_to_xl(original_df, ws.Range("F1"))
            return xl_to_df(ws.Range("F1:H4")), original_df

        result_df, original_df = xl_open.run_in_sta(_do)
        assert list(result_df.columns) == list(original_df.columns)

    def test_roundtrip_values(self, xl_open):
        """往復後の数値データが一致"""
        from auto_office_editor_mcp.helpers import df_to_xl, xl_to_df

        def _do():
            ws = xl_open.xl_app.Workbooks(WORKBOOK_NAME).Worksheets("Sheet1")
            original_df = xl_to_df(ws.Range("A1:C4"))
            df_to_xl(original_df, ws.Range("F1"))
            return xl_to_df(ws.Range("F1:H4")), original_df

        result_df, original_df = xl_open.run_in_sta(_do)
        assert list(result_df["年齢"]) == list(original_df["年齢"])
        assert list(result_df["売上"]) == list(original_df["売上"])

    def test_nan_handled(self, xl_open):
        """NaN を含む DataFrame を書き込んでもエラーにならない"""
        from auto_office_editor_mcp.helpers import df_to_xl
        import numpy as np

        df_with_nan = pd.DataFrame({"A": [1.0, float("nan"), 3.0]})
        xl_open.run_in_sta(
            lambda: df_to_xl(
                df_with_nan,
                xl_open.xl_app.Workbooks(WORKBOOK_NAME)
                .Worksheets("Sheet1")
                .Range("J1"),
            )
        )


# ===========================================================================
# execute_office_script : Excel 操作（Office 必要）
# ===========================================================================


class TestExecuteOfficeScriptExcel:
    def test_read_cell_value(self, xl_open):
        """execute_office_script でセル値を print できる"""
        from auto_office_editor_mcp.executor import execute_script

        code = f"wb = xl_app.Workbooks('{WORKBOOK_NAME}')\nws = wb.Worksheets('Sheet1')\nprint(ws.Range('A2').Value)"
        result = execute_script(code)
        assert "田中太郎" in result

    def test_chart_exists_on_sheet3(self, xl_open):
        """Sheet3 にグラフが1つあること"""
        from auto_office_editor_mcp.executor import execute_script

        code = f"wb = xl_app.Workbooks('{WORKBOOK_NAME}')\nws = wb.Worksheets('Sheet3')\nprint(ws.ChartObjects().Count)"
        result = execute_script(code)
        assert "1" in result

    def test_shape_exists_on_sheet3(self, xl_open):
        """Sheet3 に図形（星型含む）が2つあること（ChartObject + Shape）"""
        from auto_office_editor_mcp.executor import execute_script

        code = f"wb = xl_app.Workbooks('{WORKBOOK_NAME}')\nws = wb.Worksheets('Sheet3')\nprint(ws.Shapes.Count)"
        result = execute_script(code)
        assert "2" in result

    def test_write_and_read_cell(self, xl_open):
        """セルへの書き込みが反映されること"""
        from auto_office_editor_mcp.executor import execute_script

        code = (
            f"wb = xl_app.Workbooks('{WORKBOOK_NAME}')\n"
            "ws = wb.Worksheets('Sheet2')\n"
            "ws.Range('A1').Value = 'test_value'\n"
            "print(ws.Range('A1').Value)"
        )
        result = execute_script(code)
        assert "test_value" in result
