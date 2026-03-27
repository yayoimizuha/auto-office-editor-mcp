"""
Excel 編集操作の統合テスト。

前提: tests/fixtures/test.xlsx が存在すること。

各テストは xl_editable フィクスチャ（function スコープ）を使用する。
フィクスチャは毎回 test.xlsx を
    fixtures/editable/<クラス名>_<テスト名>.xlsx
にコピーして開き、テスト後に保存して閉じる。
各テストの編集結果が個別のファイルに残るため目視で確認できる。
次回実行時に同名ファイルがあれば上書きコピーしてクリーンな状態から開始する。

test_editable.xlsx の Sheet3 構成（コピー元 test.xlsx と同じ）:
    Shape 1: Chart 1   (Type=3, グラフ)
        Left: 205.1pt, Top: 198.0pt, Width: 397.9pt, Height: 321.0pt
        Bottom edge: 519.0pt
    Shape 2: 5-Point Star 2 (Type=1, オートシェイプ)
        Left: 437.2pt, Top: 31.2pt

テストする操作:
    - SmileyFace（顔アイコン）をグラフの下に追加
    - 塗りつぶし色を黄色・赤・その他に変更
    - 枠線の色を変更
    - 笑顔/しかめ面の Adjustments 変更
    - 位置・サイズの移動とリサイズ
    - 図形の削除
    - 既存の星型図形の編集
    - セルの書式設定（Bold、背景色）
"""

from __future__ import annotations

import pytest

from tests.utils import OLE_BLACK, OLE_BLUE, OLE_RED, OLE_YELLOW

pytestmark = pytest.mark.integration

SHEET3_NAME = "Sheet3"
SHEET1_NAME = "Sheet1"

# コピー元ファイルの Sheet3 の既存図形数（Chart + Star）
ORIGINAL_SHAPE_COUNT = 2

# グラフ（ChartObject）の既知の位置・サイズ
CHART_TOP = 198.0  # pt
CHART_LEFT = 205.1  # pt
CHART_HEIGHT = 321.0  # pt
CHART_BOTTOM = CHART_TOP + CHART_HEIGHT  # 519.0pt

# SmileyFace の AutoShapeType 定数
MSO_SHAPE_SMILEY_FACE = 98


def _wb_name(dst) -> str:
    """Path オブジェクトからファイル名（拡張子付き）を返す。"""
    from pathlib import Path

    return Path(dst).name


def _ws3(mgr, dst):
    """Sheet3 の Worksheet オブジェクトを返す。STA スレッド内で使うこと。"""
    return mgr.xl_app.Workbooks(_wb_name(dst)).Worksheets(SHEET3_NAME)


def _ws1(mgr, dst):
    """Sheet1 の Worksheet オブジェクトを返す。STA スレッド内で使うこと。"""
    return mgr.xl_app.Workbooks(_wb_name(dst)).Worksheets(SHEET1_NAME)


def _add_smiley(ws, left=None, top=None, width=80, height=80, name="SmileyFace_1"):
    """
    SmileyFace をグラフの直下に追加するヘルパー。STA スレッド内で使うこと。
    left/top 未指定時はグラフの左端・下端に配置する。
    """
    if left is None:
        left = CHART_LEFT
    if top is None:
        top = CHART_BOTTOM + 10  # グラフ下端 + 10pt のマージン
    shp = ws.Shapes.AddShape(MSO_SHAPE_SMILEY_FACE, left, top, width, height)
    shp.Name = name
    return shp


# ===========================================================================
# SmileyFace の追加と基本プロパティ
# ===========================================================================


class TestAddSmileyFace:
    def test_smiley_added_shape_count(self, xl_editable):
        """SmileyFace を追加すると Shapes.Count が 1 増える"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            _add_smiley(ws)
            return ws.Shapes.Count

        count = mgr.run_in_sta(_do)
        assert count == ORIGINAL_SHAPE_COUNT + 1

    def test_smiley_name(self, xl_editable):
        """追加した SmileyFace の名前が指定通りに設定される"""
        mgr, dst = xl_editable

        name = mgr.run_in_sta(
            lambda: (
                _add_smiley(_ws3(mgr, dst), name="TestFace")
                and _ws3(mgr, dst).Shapes("TestFace").Name
            )
        )

        # run_in_sta で2回 _ws3 を呼ぶのは非効率なので、まとめる
        def _do():
            ws = _ws3(mgr, dst)
            _add_smiley(ws, name="TestFace2")
            return ws.Shapes("TestFace2").Name

        assert mgr.run_in_sta(_do) == "TestFace2"

    def test_smiley_autoshape_type(self, xl_editable):
        """追加した図形の AutoShapeType が SmileyFace(98) である"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            return ws.Shapes(shp.Name).AutoShapeType

        assert mgr.run_in_sta(_do) == MSO_SHAPE_SMILEY_FACE

    def test_smiley_positioned_below_chart(self, xl_editable):
        """SmileyFace の Top がグラフの下端より大きい（グラフの下に配置される）"""
        mgr, dst = xl_editable

        top = mgr.run_in_sta(lambda: _add_smiley(_ws3(mgr, dst)).Top)

        # ただし _add_smiley 内では ws を 1回しか使わないので問題なし
        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            return ws.Shapes(shp.Name).Top

        assert mgr.run_in_sta(_do) > CHART_BOTTOM

    def test_smiley_left_aligned_with_chart(self, xl_editable):
        """SmileyFace の Left がグラフの左端と一致する"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            return ws.Shapes(shp.Name).Left

        assert abs(mgr.run_in_sta(_do) - CHART_LEFT) < 1.0

    def test_smiley_width(self, xl_editable):
        """SmileyFace の Width が指定値と一致する"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws, width=80, height=80)
            return ws.Shapes(shp.Name).Width

        assert abs(mgr.run_in_sta(_do) - 80) < 1.0

    def test_smiley_height(self, xl_editable):
        """SmileyFace の Height が指定値と一致する"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws, width=80, height=80)
            return ws.Shapes(shp.Name).Height

        assert abs(mgr.run_in_sta(_do) - 80) < 1.0

    def test_multiple_smileys(self, xl_editable):
        """SmileyFace を複数追加できる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            _add_smiley(ws, name="Face_A")
            _add_smiley(ws, left=CHART_LEFT + 100, name="Face_B")
            return ws.Shapes.Count

        assert mgr.run_in_sta(_do) == ORIGINAL_SHAPE_COUNT + 2


# ===========================================================================
# 塗りつぶし色の変更
# ===========================================================================


class TestSmileyFaceColor:
    def test_fill_color_yellow(self, xl_editable):
        """Fill.ForeColor.RGB を黄色に設定できる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            shp.Fill.ForeColor.RGB = OLE_YELLOW
            return ws.Shapes(shp.Name).Fill.ForeColor.RGB

        assert mgr.run_in_sta(_do) == OLE_YELLOW

    def test_fill_color_red(self, xl_editable):
        """Fill.ForeColor.RGB を赤に変更できる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            shp.Fill.ForeColor.RGB = OLE_YELLOW
            shp.Fill.ForeColor.RGB = OLE_RED
            return ws.Shapes(shp.Name).Fill.ForeColor.RGB

        assert mgr.run_in_sta(_do) == OLE_RED

    def test_fill_color_blue(self, xl_editable):
        """Fill.ForeColor.RGB を青に設定できる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            shp.Fill.ForeColor.RGB = OLE_BLUE
            return ws.Shapes(shp.Name).Fill.ForeColor.RGB

        assert mgr.run_in_sta(_do) == OLE_BLUE

    def test_fill_via_execute_script(self, xl_editable):
        """execute_office_script 経由で黄色に塗れる"""
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = xl_editable
        wb_name = _wb_name(dst)
        mgr.run_in_sta(lambda: _add_smiley(_ws3(mgr, dst), name="ScriptFace"))

        code = (
            f"wb = xl_app.Workbooks('{wb_name}')\n"
            f"ws = wb.Worksheets('{SHEET3_NAME}')\n"
            "shp = ws.Shapes('ScriptFace')\n"
            "shp.Fill.ForeColor.RGB = rgb(255, 255, 0)\n"
            "print(shp.Fill.ForeColor.RGB)"
        )
        result = execute_script(code)
        assert str(OLE_YELLOW) in result

    def test_line_color_black(self, xl_editable):
        """Line.ForeColor.RGB を黒に設定できる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            shp.Line.ForeColor.RGB = OLE_BLACK
            return ws.Shapes(shp.Name).Line.ForeColor.RGB

        assert mgr.run_in_sta(_do) == OLE_BLACK

    def test_line_color_red(self, xl_editable):
        """Line.ForeColor.RGB を赤に設定できる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            shp.Line.ForeColor.RGB = OLE_RED
            return ws.Shapes(shp.Name).Line.ForeColor.RGB

        assert mgr.run_in_sta(_do) == OLE_RED

    def test_fill_invisible(self, xl_editable):
        """Fill.Visible = False で塗りつぶしを非表示にできる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            shp.Fill.Visible = False
            return ws.Shapes(shp.Name).Fill.Visible

        assert mgr.run_in_sta(_do) == 0  # False = 0


# ===========================================================================
# 笑顔 / しかめ面の Adjustments
# ===========================================================================


class TestSmileyFaceExpression:
    """
    SmileyFace の Adjustments:
        Adjustments(1): 目の縦位置（デフォルト ~0.167）
        Adjustments(2): 口の形（-1=満面の笑み, 0=フラット, +1=しかめ面）
    gencache.EnsureDispatch が必要なため Adjustments.SetItem() を使う。
    """

    def test_default_adjustments(self, xl_editable):
        """デフォルトの Adjustments(2) は 0.5（やや笑顔）"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            return ws.Shapes(shp.Name).Adjustments(2)

        assert abs(mgr.run_in_sta(_do) - 0.5) < 0.01

    def test_set_smile(self, xl_editable):
        """Adjustments.SetItem(2, -0.5) で笑顔にできる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            shp.Adjustments.SetItem(2, -0.5)
            return ws.Shapes(shp.Name).Adjustments(2)

        assert abs(mgr.run_in_sta(_do) - (-0.5)) < 0.01

    def test_set_frown(self, xl_editable):
        """Adjustments.SetItem(2, 0.8) でしかめ面にできる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            shp.Adjustments.SetItem(2, 0.8)
            return ws.Shapes(shp.Name).Adjustments(2)

        assert abs(mgr.run_in_sta(_do) - 0.8) < 0.01

    def test_set_flat(self, xl_editable):
        """Adjustments.SetItem(2, 0) でフラットな表情にできる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            shp.Adjustments.SetItem(2, 0.0)
            return ws.Shapes(shp.Name).Adjustments(2)

        assert abs(mgr.run_in_sta(_do) - 0.0) < 0.01

    def test_eye_position_adjustment(self, xl_editable):
        """Adjustments.SetItem(1, 0.3) で目の位置を変更できる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            shp.Adjustments.SetItem(1, 0.3)
            return ws.Shapes(shp.Name).Adjustments(1)

        assert abs(mgr.run_in_sta(_do) - 0.3) < 0.01

    def test_smile_via_execute_script(self, xl_editable):
        """execute_office_script 経由で笑顔に変更できる"""
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = xl_editable
        wb_name = _wb_name(dst)
        mgr.run_in_sta(lambda: _add_smiley(_ws3(mgr, dst), name="ExprFace"))

        code = (
            f"wb = xl_app.Workbooks('{wb_name}')\n"
            f"ws = wb.Worksheets('{SHEET3_NAME}')\n"
            "shp = ws.Shapes('ExprFace')\n"
            "shp.Adjustments.SetItem(2, -0.5)\n"
            "print(shp.Adjustments(2))"
        )
        result = execute_script(code)
        assert "-0.5" in result


# ===========================================================================
# 位置・サイズの変更
# ===========================================================================


class TestSmileyFaceMove:
    def test_move_left(self, xl_editable):
        """Left を変更すると読み戻し値が一致する"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            shp.Left = 300.0
            return ws.Shapes(shp.Name).Left

        assert abs(mgr.run_in_sta(_do) - 300.0) < 1.0

    def test_move_top(self, xl_editable):
        """Top を変更すると読み戻し値が一致する"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            shp.Top = 600.0
            return ws.Shapes(shp.Name).Top

        assert abs(mgr.run_in_sta(_do) - 600.0) < 1.0

    def test_resize_width(self, xl_editable):
        """Width を変更すると読み戻し値が一致する"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            shp.Width = 120.0
            return ws.Shapes(shp.Name).Width

        assert abs(mgr.run_in_sta(_do) - 120.0) < 1.0

    def test_resize_height(self, xl_editable):
        """Height を変更すると読み戻し値が一致する"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            shp.Height = 120.0
            return ws.Shapes(shp.Name).Height

        assert abs(mgr.run_in_sta(_do) - 120.0) < 1.0

    def test_cm_to_pt_positioning(self, xl_editable):
        """cm_to_pt() を使った位置指定が正しく適用される（5cm = 141.7pt）"""
        from auto_office_editor_mcp.helpers import cm_to_pt

        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            shp.Left = cm_to_pt(5.0)
            return shp.Left / 28.3465

        actual_cm = mgr.run_in_sta(_do)
        assert abs(actual_cm - 5.0) < 0.05

    def test_move_via_execute_script(self, xl_editable):
        """execute_office_script 経由で位置を変更できる"""
        from auto_office_editor_mcp.executor import execute_script
        from auto_office_editor_mcp.helpers import cm_to_pt

        mgr, dst = xl_editable
        wb_name = _wb_name(dst)
        mgr.run_in_sta(lambda: _add_smiley(_ws3(mgr, dst), name="MoveFace"))

        code = (
            f"wb = xl_app.Workbooks('{wb_name}')\n"
            f"ws = wb.Worksheets('{SHEET3_NAME}')\n"
            "shp = ws.Shapes('MoveFace')\n"
            "shp.Left = cm_to_pt(3.0)\n"
            "shp.Top = cm_to_pt(20.0)\n"
            "print(round(shp.Left, 1))"
        )
        result = execute_script(code)
        expected = round(cm_to_pt(3.0), 1)
        assert str(expected) in result


# ===========================================================================
# 図形の削除
# ===========================================================================


class TestSmileyFaceDelete:
    def test_delete_restores_count(self, xl_editable):
        """SmileyFace を追加して削除すると元の Shapes.Count に戻る"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            shp = _add_smiley(ws)
            name = shp.Name
            ws.Shapes(name).Delete()
            return ws.Shapes.Count

        assert mgr.run_in_sta(_do) == ORIGINAL_SHAPE_COUNT

    def test_delete_via_execute_script(self, xl_editable):
        """execute_office_script 経由で削除できる"""
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = xl_editable
        wb_name = _wb_name(dst)
        mgr.run_in_sta(lambda: _add_smiley(_ws3(mgr, dst), name="DelFace"))

        code = (
            f"wb = xl_app.Workbooks('{wb_name}')\n"
            f"ws = wb.Worksheets('{SHEET3_NAME}')\n"
            "ws.Shapes('DelFace').Delete()\n"
            "print(ws.Shapes.Count)"
        )
        result = execute_script(code)
        assert str(ORIGINAL_SHAPE_COUNT) in result

    def test_add_and_delete_multiple(self, xl_editable):
        """複数追加して全て削除すると元の数に戻る"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            names = ["Face_X", "Face_Y", "Face_Z"]
            for i, name in enumerate(names):
                _add_smiley(ws, left=CHART_LEFT + i * 100, name=name)
            count_added = ws.Shapes.Count
            for name in names:
                ws.Shapes(name).Delete()
            return count_added, ws.Shapes.Count

        count_added, count_after = mgr.run_in_sta(_do)
        assert count_added == ORIGINAL_SHAPE_COUNT + 3
        assert count_after == ORIGINAL_SHAPE_COUNT


# ===========================================================================
# 既存の星型図形の編集
# ===========================================================================


class TestEditExistingStar:
    STAR_NAME = "5-Point Star 2"

    def test_star_exists(self, xl_editable):
        """既存の星型図形が存在する"""
        mgr, dst = xl_editable
        star = mgr.run_in_sta(lambda: _ws3(mgr, dst).Shapes(self.STAR_NAME))
        assert star is not None

    def test_star_fill_color_change(self, xl_editable):
        """星型図形の塗りつぶし色を黄色に変更できる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            ws.Shapes(self.STAR_NAME).Fill.ForeColor.RGB = OLE_YELLOW
            return ws.Shapes(self.STAR_NAME).Fill.ForeColor.RGB

        assert mgr.run_in_sta(_do) == OLE_YELLOW

    def test_star_fill_color_blue(self, xl_editable):
        """星型図形の塗りつぶし色を青に変更できる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            ws.Shapes(self.STAR_NAME).Fill.ForeColor.RGB = OLE_BLUE
            return ws.Shapes(self.STAR_NAME).Fill.ForeColor.RGB

        assert mgr.run_in_sta(_do) == OLE_BLUE

    def test_star_move(self, xl_editable):
        """星型図形を移動できる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            ws.Shapes(self.STAR_NAME).Left = 50.0
            ws.Shapes(self.STAR_NAME).Top = 50.0
            return ws.Shapes(self.STAR_NAME).Left, ws.Shapes(self.STAR_NAME).Top

        left, top = mgr.run_in_sta(_do)
        assert abs(left - 50.0) < 1.0
        assert abs(top - 50.0) < 1.0

    def test_star_resize(self, xl_editable):
        """星型図形のサイズを変更できる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws3(mgr, dst)
            ws.Shapes(self.STAR_NAME).Width = 100.0
            ws.Shapes(self.STAR_NAME).Height = 100.0
            return ws.Shapes(self.STAR_NAME).Width, ws.Shapes(self.STAR_NAME).Height

        width, height = mgr.run_in_sta(_do)
        assert abs(width - 100.0) < 1.0
        assert abs(height - 100.0) < 1.0

    def test_star_edit_via_execute_script(self, xl_editable):
        """execute_office_script 経由で星の色と位置を変更できる"""
        from auto_office_editor_mcp.executor import execute_script
        from auto_office_editor_mcp.helpers import cm_to_pt

        mgr, dst = xl_editable
        wb_name = _wb_name(dst)

        code = (
            f"wb = xl_app.Workbooks('{wb_name}')\n"
            f"ws = wb.Worksheets('{SHEET3_NAME}')\n"
            f"star = ws.Shapes('{self.STAR_NAME}')\n"
            "star.Fill.ForeColor.RGB = rgb(255, 255, 0)\n"
            "star.Left = cm_to_pt(2.0)\n"
            "print(star.Fill.ForeColor.RGB, round(star.Left, 1))"
        )
        result = execute_script(code)
        assert str(OLE_YELLOW) in result
        assert str(round(cm_to_pt(2.0), 1)) in result


# ===========================================================================
# セルの書式設定
# ===========================================================================

SHEET2_NAME = "Sheet2"


def _ws2(mgr, dst):
    """Sheet2 の Worksheet オブジェクトを返す。STA スレッド内で使うこと。"""
    return mgr.xl_app.Workbooks(_wb_name(dst)).Worksheets(SHEET2_NAME)


def _write_sample_data(ws):
    """
    目視確認用のサンプルデータを Sheet2 の A1:C3 に書き込む。STA スレッド内で使うこと。

        A1: "見出しA"  B1: "見出しB"  C1: "見出しC"
        A2: "データ1"  B2: 100        C2: 200
        A3: "データ2"  B3: 300        C3: 400
    """
    ws.Range("A1").Value = "見出しA"
    ws.Range("B1").Value = "見出しB"
    ws.Range("C1").Value = "見出しC"
    ws.Range("A2").Value = "データ1"
    ws.Range("B2").Value = 100
    ws.Range("C2").Value = 200
    ws.Range("A3").Value = "データ2"
    ws.Range("B3").Value = 300
    ws.Range("C3").Value = 400


class TestCellFormatting:
    """
    各テストは Sheet2（空シート）にサンプルデータを書き込んでから
    書式を設定する。元が空なので変更前後の差が一目でわかる。
    """

    def test_cell_bold(self, xl_editable):
        """見出し行（A1:C1）を Bold に設定できる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws2(mgr, dst)
            _write_sample_data(ws)
            ws.Range("A1:C1").Font.Bold = True
            return [ws.Range(f"{col}1").Font.Bold for col in ["A", "B", "C"]]

        results = mgr.run_in_sta(_do)
        for bold in results:
            assert bold is True

    def test_cell_font_size(self, xl_editable):
        """見出し行のフォントサイズを 18pt に変更できる（デフォルト 11pt から大幅に変化）"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws2(mgr, dst)
            _write_sample_data(ws)
            ws.Range("A1:C1").Font.Size = 18
            return ws.Range("A1").Font.Size, ws.Range("B1").Font.Size

        a1_size, b1_size = mgr.run_in_sta(_do)
        assert a1_size == 18
        assert b1_size == 18

    def test_cell_background_color_yellow(self, xl_editable):
        """見出し行の背景色を黄色に設定できる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws2(mgr, dst)
            _write_sample_data(ws)
            ws.Range("A1:C1").Interior.Color = OLE_YELLOW
            return ws.Range("A1").Interior.Color, ws.Range("C1").Interior.Color

        a1_color, c1_color = mgr.run_in_sta(_do)
        assert a1_color == OLE_YELLOW
        assert c1_color == OLE_YELLOW

    def test_cell_background_color_red(self, xl_editable):
        """データ行の背景色を赤に設定できる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws2(mgr, dst)
            _write_sample_data(ws)
            ws.Range("A2:C2").Interior.Color = OLE_RED
            return ws.Range("A2").Interior.Color, ws.Range("C2").Interior.Color

        a2_color, c2_color = mgr.run_in_sta(_do)
        assert a2_color == OLE_RED
        assert c2_color == OLE_RED

    def test_cell_font_color(self, xl_editable):
        """見出し行のフォント色を青に設定できる"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws2(mgr, dst)
            _write_sample_data(ws)
            ws.Range("A1:C1").Font.Color = OLE_BLUE
            return ws.Range("A1").Font.Color

        assert mgr.run_in_sta(_do) == OLE_BLUE

    def test_range_bold(self, xl_editable):
        """見出し行を Bold、データ行を通常のまま → 視覚的な差が明確"""
        mgr, dst = xl_editable

        def _do():
            ws = _ws2(mgr, dst)
            _write_sample_data(ws)
            ws.Range("A1:C1").Font.Bold = True
            header_bolds = [ws.Range(f"{col}1").Font.Bold for col in ["A", "B", "C"]]
            data_bolds = [ws.Range(f"{col}2").Font.Bold for col in ["A", "B", "C"]]
            return header_bolds, data_bolds

        header_bolds, data_bolds = mgr.run_in_sta(_do)
        for bold in header_bolds:
            assert bold is True
        for bold in data_bolds:
            assert bold is not True

    def test_cell_formatting_via_execute_script(self, xl_editable):
        """execute_office_script 経由で複数の書式を一括設定できる"""
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = xl_editable
        wb_name = _wb_name(dst)
        mgr.run_in_sta(lambda: _write_sample_data(_ws2(mgr, dst)))

        code = (
            f"wb = xl_app.Workbooks('{wb_name}')\n"
            f"ws = wb.Worksheets('{SHEET2_NAME}')\n"
            # 見出し行：黄色背景 + Bold + 14pt
            "ws.Range('A1:C1').Interior.Color = rgb(255, 255, 0)\n"
            "ws.Range('A1:C1').Font.Bold = True\n"
            "ws.Range('A1:C1').Font.Size = 14\n"
            # データ行：交互に色付け
            "ws.Range('A2:C2').Interior.Color = rgb(204, 229, 255)\n"
            "ws.Range('A3:C3').Interior.Color = rgb(240, 240, 240)\n"
            "print(ws.Range('A1').Interior.Color)\n"
            "print(ws.Range('A1').Font.Bold)\n"
            "print(ws.Range('A1').Font.Size)"
        )
        result = execute_script(code)
        assert str(OLE_YELLOW) in result
        assert "True" in result
        assert "14" in result


# ===========================================================================
# 複合操作テスト（SmileyFace + 色 + 表情 + 位置を一括設定）
# ===========================================================================


class TestCompositeEdit:
    def test_full_smiley_setup(self, xl_editable):
        """
        グラフの下に黄色い笑顔の SmileyFace を追加する複合操作テスト。
        execute_office_script で一括実行し、全プロパティを確認する。
        """
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = xl_editable
        wb_name = _wb_name(dst)

        code = (
            f"wb = xl_app.Workbooks('{wb_name}')\n"
            f"ws = wb.Worksheets('{SHEET3_NAME}')\n"
            "chart = ws.ChartObjects(1)\n"
            "top = chart.Top + chart.Height + 10\n"
            "shp = ws.Shapes.AddShape(98, chart.Left, top, 80, 80)\n"
            "shp.Name = 'HappyFace'\n"
            "shp.Fill.ForeColor.RGB = rgb(255, 255, 0)\n"
            "shp.Line.ForeColor.RGB = rgb(0, 0, 0)\n"
            "shp.Adjustments.SetItem(2, -0.5)\n"
            "print(shp.Name)\n"
            "print(shp.Fill.ForeColor.RGB)\n"
            "print(shp.Adjustments(2))\n"
            "print(ws.Shapes.Count)"
        )
        result = execute_script(code)

        assert "HappyFace" in result
        assert str(OLE_YELLOW) in result
        assert "-0.5" in result
        assert str(ORIGINAL_SHAPE_COUNT + 1) in result

    def test_smiley_and_star_together(self, xl_editable):
        """SmileyFace を追加しつつ既存の星の色も変更する複合テスト"""
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = xl_editable
        wb_name = _wb_name(dst)

        code = (
            f"wb = xl_app.Workbooks('{wb_name}')\n"
            f"ws = wb.Worksheets('{SHEET3_NAME}')\n"
            "chart = ws.ChartObjects(1)\n"
            "shp = ws.Shapes.AddShape(98, chart.Left, chart.Top + chart.Height + 20, 60, 60)\n"
            "shp.Name = 'ComboFace'\n"
            "shp.Fill.ForeColor.RGB = rgb(255, 255, 0)\n"
            "star = ws.Shapes('5-Point Star 2')\n"
            "star.Fill.ForeColor.RGB = rgb(0, 0, 255)\n"
            "print(ws.Shapes.Count)\n"
            "print(star.Fill.ForeColor.RGB)"
        )
        result = execute_script(code)

        assert str(ORIGINAL_SHAPE_COUNT + 1) in result
        assert str(OLE_BLUE) in result
