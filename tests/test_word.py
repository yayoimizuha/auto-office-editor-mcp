"""
Word 統合テスト。

前提: tests/fixtures/test.docx が存在すること。

test.docx の構成:
  段落1（見出し 1）: "売上レポート"
  段落2（標準）:     "本レポートは2024年度の売上実績をまとめたものです。"
  段落3（見出し 2）: "概要"
  段落4（標準）:     "全体の売上は前年比120%となりました。"
  段落5（見出し 2）: "部門別売上"
  表1（3行×3列）:
    ヘッダー: 部門 / 売上 / 前年比
    行1:      東京  / 1500000 / 120%
    行2:      大阪  / 980000  / 105%
  末尾空段落

注意: doc.Paragraphs は表セル内の段落もカウントするため、
  paragraphs_count は 19（本文5段落 + 表セル9 + その他）になる。

情報取得テストは wd_open フィクスチャ（保存なし）を使用。
編集テストは wd_editable フィクスチャを使用し、
  fixtures/editable/<クラス名>_<テスト名>.docx に保存する。
"""

from __future__ import annotations

import json

import pandas as pd
import pytest

pytestmark = pytest.mark.integration

DOC_NAME = "test.docx"

# test.docx の既知の値
FIRST_HEADING_TEXT = "売上レポート"
FIRST_HEADING_STYLE = "見出し 1"
BODY_PARAGRAPH_TEXT = "本レポートは2024年度の売上実績をまとめたものです。"
TABLES_COUNT = 1
TABLE1_ROWS = 3
TABLE1_COLS = 3
TABLE1_HEADERS = ["部門", "売上", "前年比"]
TABLE1_ROW1 = ["東京", "1500000", "120%"]
TABLE1_ROW2 = ["大阪", "980000", "105%"]


def _doc_name(dst) -> str:
    from pathlib import Path

    return Path(dst).name


def _doc(mgr, dst):
    """編集用ドキュメントオブジェクトを返す。STA スレッド内で使うこと。"""
    return mgr.wd_app.Documents(_doc_name(dst))


# ===========================================================================
# ファイルオープン / クローズ
# ===========================================================================


class TestOpenClose:
    def test_open_returns_correct_app(self, wd_open):
        """open_file が word アプリを返す"""
        from tests.conftest import WORD_PATH

        result = wd_open.open_file(str(WORD_PATH))
        assert result["app"] == "word"

    def test_open_returns_correct_name(self, wd_open):
        """open_file がファイル名を返す"""
        from tests.conftest import WORD_PATH

        result = wd_open.open_file(str(WORD_PATH))
        assert result["name"] == DOC_NAME

    def test_open_already_open_returns_flag(self, wd_open):
        """既に開いているファイルを再度開くと already_open=True"""
        from tests.conftest import WORD_PATH

        result = wd_open.open_file(str(WORD_PATH))
        assert result["already_open"] is True

    def test_open_nonexistent_raises(self, app_manager):
        """存在しないファイルを開こうとすると FileNotFoundError"""
        with pytest.raises(FileNotFoundError):
            app_manager.open_file("C:/nonexistent/hoge.docx")

    def test_open_unsupported_ext_raises(self, app_manager):
        """未対応拡張子は ValueError"""
        with pytest.raises(ValueError, match="未対応"):
            app_manager.open_file("C:/hoge.txt")


# ===========================================================================
# get_open_files
# ===========================================================================


class TestGetOpenFiles:
    def test_word_file_in_list(self, wd_open):
        """開いた後は get_open_files の word リストに含まれる"""
        files = wd_open.get_open_files()
        assert DOC_NAME in files["word"]

    def test_excel_list_empty(self, wd_open):
        """Excel ファイルは開いていない"""
        files = wd_open.get_open_files()
        assert files["excel"] == []

    def test_powerpoint_list_empty(self, wd_open):
        """PowerPoint の test.pptx は開いていない（他のファイルが開いていても構わない）"""
        files = wd_open.get_open_files()
        assert "test.pptx" not in files["powerpoint"]


# ===========================================================================
# get_wd_document_info / ドキュメント情報取得
# ===========================================================================


class TestGetWordDocumentInfo:
    def test_paragraphs_count_is_positive(self, wd_open):
        """paragraphs_count が正の整数であること"""
        from auto_office_editor_mcp.tools.info import get_wd_document_info

        result = json.loads(get_wd_document_info(DOC_NAME))
        assert result["paragraphs_count"] > 0

    def test_tables_count(self, wd_open):
        """表が1つあること"""
        from auto_office_editor_mcp.tools.info import get_wd_document_info

        result = json.loads(get_wd_document_info(DOC_NAME))
        assert result["tables_count"] == TABLES_COUNT

    def test_table1_rows(self, wd_open):
        """表1の行数が3であること"""
        from auto_office_editor_mcp.tools.info import get_wd_document_info

        result = json.loads(get_wd_document_info(DOC_NAME))
        assert result["tables"][0]["rows"] == TABLE1_ROWS

    def test_table1_columns(self, wd_open):
        """表1の列数が3であること"""
        from auto_office_editor_mcp.tools.info import get_wd_document_info

        result = json.loads(get_wd_document_info(DOC_NAME))
        assert result["tables"][0]["columns"] == TABLE1_COLS

    def test_first_paragraph_style(self, wd_open):
        """先頭段落のスタイルが '見出し 1' であること"""
        from auto_office_editor_mcp.tools.info import get_wd_document_info

        result = json.loads(get_wd_document_info(DOC_NAME))
        assert result["paragraphs"][0]["style"] == FIRST_HEADING_STYLE

    def test_first_paragraph_text(self, wd_open):
        """先頭段落のテキストが '売上レポート' であること"""
        from auto_office_editor_mcp.tools.info import get_wd_document_info

        result = json.loads(get_wd_document_info(DOC_NAME))
        assert result["paragraphs"][0]["text"] == FIRST_HEADING_TEXT

    def test_paragraphs_have_style_key(self, wd_open):
        """全段落エントリに style キーが存在すること"""
        from auto_office_editor_mcp.tools.info import get_wd_document_info

        result = json.loads(get_wd_document_info(DOC_NAME))
        for para in result["paragraphs"]:
            assert "style" in para

    def test_paragraphs_have_text_key(self, wd_open):
        """全段落エントリに text キーが存在すること"""
        from auto_office_editor_mcp.tools.info import get_wd_document_info

        result = json.loads(get_wd_document_info(DOC_NAME))
        for para in result["paragraphs"]:
            assert "text" in para

    def test_invalid_doc_name_returns_error(self, app_manager):
        """存在しないドキュメント名はエラーを返す"""
        from auto_office_editor_mcp.tools.info import get_wd_document_info

        result = json.loads(get_wd_document_info("nonexistent.docx"))
        assert "error" in result

    def test_result_has_name_key(self, wd_open):
        """結果に name キーが存在すること"""
        from auto_office_editor_mcp.tools.info import get_wd_document_info

        result = json.loads(get_wd_document_info(DOC_NAME))
        assert result["name"] == DOC_NAME


# ===========================================================================
# wd_table_to_df : Word 表 → DataFrame 変換
# ===========================================================================


class TestWdTableToDf:
    def test_shape(self, wd_open):
        """表1を DataFrame に変換すると 2行3列（ヘッダー除く）"""
        from auto_office_editor_mcp.helpers import wd_table_to_df

        df = wd_open.run_in_sta(
            lambda: wd_table_to_df(wd_open.wd_app.Documents(DOC_NAME).Tables(1))
        )
        assert df.shape == (2, 3)

    def test_headers(self, wd_open):
        """ヘッダー行が正しく取得される"""
        from auto_office_editor_mcp.helpers import wd_table_to_df

        df = wd_open.run_in_sta(
            lambda: wd_table_to_df(wd_open.wd_app.Documents(DOC_NAME).Tables(1))
        )
        assert list(df.columns) == TABLE1_HEADERS

    def test_first_row_values(self, wd_open):
        """1行目のデータが正しく取得される"""
        from auto_office_editor_mcp.helpers import wd_table_to_df

        df = wd_open.run_in_sta(
            lambda: wd_table_to_df(wd_open.wd_app.Documents(DOC_NAME).Tables(1))
        )
        assert list(df.iloc[0]) == TABLE1_ROW1

    def test_second_row_values(self, wd_open):
        """2行目のデータが正しく取得される"""
        from auto_office_editor_mcp.helpers import wd_table_to_df

        df = wd_open.run_in_sta(
            lambda: wd_table_to_df(wd_open.wd_app.Documents(DOC_NAME).Tables(1))
        )
        assert list(df.iloc[1]) == TABLE1_ROW2

    def test_control_chars_stripped(self, wd_open):
        """Word 特有の制御文字（\\r\\x07 等）が除去されている"""
        from auto_office_editor_mcp.helpers import wd_table_to_df

        df = wd_open.run_in_sta(
            lambda: wd_table_to_df(wd_open.wd_app.Documents(DOC_NAME).Tables(1))
        )
        for val in df.values.flatten():
            assert "\x07" not in str(val)
            assert "\r" not in str(val)


# ===========================================================================
# df_to_wd_table : DataFrame → Word 表追加（編集テスト）
# ===========================================================================


class TestDfToWdTable:
    def test_roundtrip_shape(self, wd_editable):
        """DataFrame → Word 表 → DataFrame で行列数が一致する"""
        from auto_office_editor_mcp.helpers import df_to_wd_table, wd_table_to_df

        mgr, dst = wd_editable
        df_src = pd.DataFrame(
            {"商品": ["りんご", "みかん"], "価格": ["150", "80"], "在庫": ["50", "120"]}
        )

        def _do():
            doc = _doc(mgr, dst)
            df_to_wd_table(df_src, doc)
            added_table = doc.Tables(doc.Tables.Count)
            return wd_table_to_df(added_table)

        df_result = mgr.run_in_sta(_do)
        assert df_result.shape == df_src.shape

    def test_roundtrip_columns(self, wd_editable):
        """往復後の列名が一致する"""
        from auto_office_editor_mcp.helpers import df_to_wd_table, wd_table_to_df

        mgr, dst = wd_editable
        df_src = pd.DataFrame({"A": ["a1", "a2"], "B": ["b1", "b2"]})

        def _do():
            doc = _doc(mgr, dst)
            df_to_wd_table(df_src, doc)
            return wd_table_to_df(doc.Tables(doc.Tables.Count))

        df_result = mgr.run_in_sta(_do)
        assert list(df_result.columns) == list(df_src.columns)

    def test_roundtrip_values(self, wd_editable):
        """往復後のデータ値が一致する"""
        from auto_office_editor_mcp.helpers import df_to_wd_table, wd_table_to_df

        mgr, dst = wd_editable
        df_src = pd.DataFrame({"名前": ["田中", "鈴木"], "点数": ["90", "75"]})

        def _do():
            doc = _doc(mgr, dst)
            df_to_wd_table(df_src, doc)
            return wd_table_to_df(doc.Tables(doc.Tables.Count))

        df_result = mgr.run_in_sta(_do)
        assert list(df_result["名前"]) == ["田中", "鈴木"]
        assert list(df_result["点数"]) == ["90", "75"]

    def test_table_count_increases(self, wd_editable):
        """表を追加すると Tables.Count が増える"""
        from auto_office_editor_mcp.helpers import df_to_wd_table

        mgr, dst = wd_editable
        df_src = pd.DataFrame({"X": ["1", "2"]})

        def _do():
            doc = _doc(mgr, dst)
            count_before = doc.Tables.Count
            df_to_wd_table(df_src, doc)
            return count_before, doc.Tables.Count

        count_before, count_after = mgr.run_in_sta(_do)
        assert count_after == count_before + 1

    def test_table_rows_and_cols(self, wd_editable):
        """追加した表の行数・列数が DataFrame と一致する（ヘッダー行含む）"""
        from auto_office_editor_mcp.helpers import df_to_wd_table

        mgr, dst = wd_editable
        df_src = pd.DataFrame({"Col1": ["v1", "v2", "v3"], "Col2": ["a", "b", "c"]})

        def _do():
            doc = _doc(mgr, dst)
            df_to_wd_table(df_src, doc)
            added = doc.Tables(doc.Tables.Count)
            return added.Rows.Count, added.Columns.Count

        rows, cols = mgr.run_in_sta(_do)
        # include_header=True なのでヘッダー1行 + データ3行 = 4行
        assert rows == len(df_src) + 1
        assert cols == len(df_src.columns)


# ===========================================================================
# execute_office_script 経由の Word 操作
# ===========================================================================


class TestExecuteOfficeScriptWord:
    def test_read_paragraph_text(self, wd_open):
        """execute_script でドキュメントの段落テキストを読み取れる"""
        from auto_office_editor_mcp.executor import execute_script

        code = (
            f"doc = wd_app.Documents('{DOC_NAME}')\n"
            "print(doc.Paragraphs(1).Range.Text.strip())"
        )
        result = execute_script(code)
        assert FIRST_HEADING_TEXT in result

    def test_add_paragraph(self, wd_editable):
        """execute_script で段落を追加できる"""
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = wd_editable
        doc_name = _doc_name(dst)
        count_before = mgr.run_in_sta(
            lambda: mgr.wd_app.Documents(doc_name).Paragraphs.Count
        )

        code = (
            f"doc = wd_app.Documents('{doc_name}')\n"
            "rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)\n"
            "rng.InsertParagraphAfter()\n"
            "p = doc.Paragraphs(doc.Paragraphs.Count)\n"
            "p.Range.Text = '追加された段落'\n"
            "print(doc.Paragraphs.Count)"
        )
        result = execute_script(code)
        # result は "実行成功\n\n--- 出力 ---\n<数値>" の形式
        output_line = [l for l in result.splitlines() if l.strip().isdigit()]
        assert output_line and int(output_line[0].strip()) > count_before

    def test_add_table_via_script(self, wd_editable):
        """execute_script で表を追加できる"""
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = wd_editable
        doc_name = _doc_name(dst)
        count_before = mgr.run_in_sta(
            lambda: mgr.wd_app.Documents(doc_name).Tables.Count
        )

        code = (
            f"doc = wd_app.Documents('{doc_name}')\n"
            "rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)\n"
            "table = doc.Tables.Add(Range=rng, NumRows=2, NumColumns=2)\n"
            "table.Cell(1,1).Range.Text = 'A'\n"
            "table.Cell(1,2).Range.Text = 'B'\n"
            "print(doc.Tables.Count)"
        )
        result = execute_script(code)
        # result は "実行成功\n\n--- 出力 ---\n<数値>" の形式
        assert str(count_before + 1) in result

    def test_bold_formatting(self, wd_editable):
        """execute_script で Bold 書式を設定・読み取れる"""
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = wd_editable
        doc_name = _doc_name(dst)

        code = (
            f"doc = wd_app.Documents('{doc_name}')\n"
            "doc.Paragraphs(1).Range.Bold = True\n"
            "print(doc.Paragraphs(1).Range.Bold)"
        )
        result = execute_script(code)
        assert "True" in result or "1" in result

    def test_font_size(self, wd_editable):
        """execute_script でフォントサイズを変更できる"""
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = wd_editable
        doc_name = _doc_name(dst)

        code = (
            f"doc = wd_app.Documents('{doc_name}')\n"
            "doc.Paragraphs(2).Range.Font.Size = 24\n"
            "print(doc.Paragraphs(2).Range.Font.Size)"
        )
        result = execute_script(code)
        assert "24" in result

    def test_font_color(self, wd_editable):
        """execute_script でフォント色を変更できる"""
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = wd_editable
        doc_name = _doc_name(dst)

        code = (
            f"doc = wd_app.Documents('{doc_name}')\n"
            "doc.Paragraphs(1).Range.Font.Color = rgb(255, 0, 0)\n"  # 赤
            "print(doc.Paragraphs(1).Range.Font.Color)"
        )
        result = execute_script(code)
        # OLE: rgb(255,0,0) = 255
        assert "255" in result

    def test_edit_table_cell(self, wd_editable):
        """execute_script で既存の表セルのテキストを編集できる"""
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = wd_editable
        doc_name = _doc_name(dst)

        code = (
            f"doc = wd_app.Documents('{doc_name}')\n"
            "table = doc.Tables(1)\n"
            "table.Cell(2,1).Range.Text = '名古屋'\n"
            "print(table.Cell(2,1).Range.Text.strip())"
        )
        result = execute_script(code)
        assert "名古屋" in result

    def test_read_table_cell(self, wd_open):
        """execute_script で既存の表セルのテキストを読み取れる"""
        from auto_office_editor_mcp.executor import execute_script

        code = (
            f"doc = wd_app.Documents('{DOC_NAME}')\n"
            "table = doc.Tables(1)\n"
            "text = table.Cell(1,1).Range.Text.replace('\\r\\x07','').replace('\\x07','').strip()\n"
            "print(text)"
        )
        result = execute_script(code)
        assert TABLE1_HEADERS[0] in result


# ===========================================================================
# 複合操作テスト（段落追加 + 表追加 + 書式設定）
# ===========================================================================


class TestCompositeEditWord:
    def test_full_document_edit(self, wd_editable):
        """
        段落追加・表追加・書式設定を一括で行う複合テスト。
        execute_script で全操作を実行し、結果を確認する。
        """
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = wd_editable
        doc_name = _doc_name(dst)

        code = (
            f"doc = wd_app.Documents('{doc_name}')\n"
            # 新しい見出し段落を追加
            "rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)\n"
            "rng.InsertParagraphAfter()\n"
            "p = doc.Paragraphs(doc.Paragraphs.Count)\n"
            "p.Style = doc.Styles('見出し 2')\n"
            "p.Range.Text = '新規セクション'\n"
            # さらに標準段落を追加
            "rng2 = doc.Range(doc.Content.End - 1, doc.Content.End - 1)\n"
            "rng2.InsertParagraphAfter()\n"
            "p2 = doc.Paragraphs(doc.Paragraphs.Count)\n"
            "p2.Style = doc.Styles('標準')\n"
            "p2.Range.Text = '追加されたコンテンツ'\n"
            # 表を追加
            "rng3 = doc.Range(doc.Content.End - 1, doc.Content.End - 1)\n"
            "tbl = doc.Tables.Add(Range=rng3, NumRows=2, NumColumns=3)\n"
            "tbl.Cell(1,1).Range.Text = '項目'\n"
            "tbl.Cell(1,2).Range.Text = '値'\n"
            "tbl.Cell(1,3).Range.Text = '備考'\n"
            "tbl.Cell(2,1).Range.Text = 'データA'\n"
            "tbl.Cell(2,2).Range.Text = '100'\n"
            "tbl.Cell(2,3).Range.Text = '正常'\n"
            # 見出し1を Bold + 赤色に
            "doc.Paragraphs(1).Range.Bold = True\n"
            "doc.Paragraphs(1).Range.Font.Color = rgb(255, 0, 0)\n"
            "print(doc.Tables.Count)\n"
            "print(doc.Paragraphs(1).Range.Bold)\n"
            "print(tbl.Cell(2,1).Range.Text.strip())"
        )
        result = execute_script(code)

        # 表が1つ増えているはず（元の1 + 追加の1 = 2）
        assert "2" in result
        # Bold が True
        assert "True" in result or "1" in result
        # セルのテキスト
        assert "データA" in result
