"""
PowerPoint 統合テスト。

前提: tests/fixtures/test.pptx が存在すること。

test.pptx の構成:
  スライド1（Slide1, layout=1）:
    Shape[1] "Title 1"    (type=14, has_text=True): "四半期売上報告"
    Shape[2] "Subtitle 2" (type=14, has_text=True): "2024年度 第4四半期"

  スライド2（Slide2, layout=2）:
    Shape[1] "Title 1"  (type=14, has_text=True): "部門別売上"
    Shape[2] "Table 3"  (type=19, has_table=True): 3行×3列の表

  スライド3（Slide3, layout=12）:
    Shape[1] "TextBox 1" (type=17, has_text=True): "グラフスライド"
    Shape[2] "SampleRect" (type=1, has_text=True): left≈3cm, top≈5cm

表2 のデータ:
    ヘッダー: 部門 / 売上 / 前年比
    行1:      東京  / 1500000 / 120%
    行2:      大阪  / 980000  / 105%

情報取得テストは ppt_open フィクスチャ（保存なし）を使用。
編集テストは ppt_editable フィクスチャを使用し、
  fixtures/editable/<クラス名>_<テスト名>.pptx に保存する。
"""

from __future__ import annotations

import json

import pandas as pd
import pytest

pytestmark = pytest.mark.integration

PRES_NAME = "test.pptx"

# test.pptx の既知の値
SLIDE_COUNT = 3
SLIDE1_TITLE = "四半期売上報告"
SLIDE1_SUBTITLE = "2024年度 第4四半期"
SLIDE2_TITLE = "部門別売上"
SLIDE3_TEXTBOX_TEXT = "グラフスライド"
SAMPLE_RECT_NAME = "SampleRect"
SAMPLE_RECT_LEFT_CM = 3.0
SAMPLE_RECT_TOP_CM = 5.0

TABLE_HEADERS = ["部門", "売上", "前年比"]
TABLE_ROW1 = ["東京", "1500000", "120%"]
TABLE_ROW2 = ["大阪", "980000", "105%"]

# 各スライドの図形数
SLIDE1_SHAPE_COUNT = 2
SLIDE2_SHAPE_COUNT = 2
SLIDE3_SHAPE_COUNT = 2


# 色定数（OLE BGR 順）
def ole_yellow():
    return 255 + (255 * 256) + (0 * 65536)


def ole_red():
    return 255 + (0 * 256) + (0 * 65536)


def ole_blue():
    return 0 + (0 * 256) + (255 * 65536)


def _pres_name(dst) -> str:
    from pathlib import Path

    return Path(dst).name


def _pres(mgr, dst):
    """編集用プレゼンテーションオブジェクトを返す。STA スレッド内で使うこと。"""
    return mgr.ppt_app.Presentations(_pres_name(dst))


def _slide(mgr, dst, index: int):
    """スライドオブジェクトを返す（1始まり）。STA スレッド内で使うこと。"""
    return _pres(mgr, dst).Slides(index)


# ===========================================================================
# ファイルオープン / クローズ
# ===========================================================================


class TestOpenClose:
    def test_open_returns_correct_app(self, ppt_open):
        """open_file が powerpoint アプリを返す"""
        from tests.conftest import PPT_PATH

        result = ppt_open.open_file(str(PPT_PATH))
        assert result["app"] == "powerpoint"

    def test_open_returns_correct_name(self, ppt_open):
        """open_file がファイル名を返す"""
        from tests.conftest import PPT_PATH

        result = ppt_open.open_file(str(PPT_PATH))
        assert result["name"] == PRES_NAME

    def test_open_already_open_returns_flag(self, ppt_open):
        """既に開いているファイルを再度開くと already_open=True"""
        from tests.conftest import PPT_PATH

        result = ppt_open.open_file(str(PPT_PATH))
        assert result["already_open"] is True

    def test_open_nonexistent_raises(self, app_manager):
        """存在しないファイルを開こうとすると FileNotFoundError"""
        with pytest.raises(FileNotFoundError):
            app_manager.open_file("C:/nonexistent/hoge.pptx")

    def test_open_unsupported_ext_raises(self, app_manager):
        """未対応拡張子は ValueError"""
        with pytest.raises(ValueError, match="未対応"):
            app_manager.open_file("C:/hoge.txt")


# ===========================================================================
# get_open_files
# ===========================================================================


class TestGetOpenFiles:
    def test_ppt_file_in_list(self, ppt_open):
        """開いた後は get_open_files の powerpoint リストに含まれる"""
        files = ppt_open.get_open_files()
        assert PRES_NAME in files["powerpoint"]

    def test_excel_list_empty(self, ppt_open):
        """Excel ファイルは開いていない"""
        files = ppt_open.get_open_files()
        assert files["excel"] == []

    def test_word_list_empty(self, ppt_open):
        """Word の test.docx は開いていない（他のファイルが開いていても構わない）"""
        files = ppt_open.get_open_files()
        assert "test.docx" not in files["word"]


# ===========================================================================
# get_ppt_slides_info / スライド一覧取得
# ===========================================================================


class TestGetPptSlidesInfo:
    def test_slide_count(self, ppt_open):
        """スライド数が3であること"""
        from auto_office_editor_mcp.tools.info import get_ppt_slides_info

        result = json.loads(get_ppt_slides_info(PRES_NAME))
        assert len(result) == SLIDE_COUNT

    def test_slide_indices(self, ppt_open):
        """インデックスが 1始まりで連番であること"""
        from auto_office_editor_mcp.tools.info import get_ppt_slides_info

        result = json.loads(get_ppt_slides_info(PRES_NAME))
        assert [s["index"] for s in result] == [1, 2, 3]

    def test_slide_has_name(self, ppt_open):
        """全スライドに name キーが存在すること"""
        from auto_office_editor_mcp.tools.info import get_ppt_slides_info

        result = json.loads(get_ppt_slides_info(PRES_NAME))
        for s in result:
            assert "name" in s

    def test_slide_has_layout(self, ppt_open):
        """全スライドに layout キーが存在すること"""
        from auto_office_editor_mcp.tools.info import get_ppt_slides_info

        result = json.loads(get_ppt_slides_info(PRES_NAME))
        for s in result:
            assert "layout" in s

    def test_slide_has_shapes_count(self, ppt_open):
        """全スライドに shapes_count キーが存在すること"""
        from auto_office_editor_mcp.tools.info import get_ppt_slides_info

        result = json.loads(get_ppt_slides_info(PRES_NAME))
        for s in result:
            assert "shapes_count" in s

    def test_slide1_shapes_count(self, ppt_open):
        """スライド1の shapes_count が 2 であること"""
        from auto_office_editor_mcp.tools.info import get_ppt_slides_info

        result = json.loads(get_ppt_slides_info(PRES_NAME))
        assert result[0]["shapes_count"] == SLIDE1_SHAPE_COUNT

    def test_slide2_shapes_count(self, ppt_open):
        """スライド2の shapes_count が 2 であること"""
        from auto_office_editor_mcp.tools.info import get_ppt_slides_info

        result = json.loads(get_ppt_slides_info(PRES_NAME))
        assert result[1]["shapes_count"] == SLIDE2_SHAPE_COUNT

    def test_slide3_shapes_count(self, ppt_open):
        """スライド3の shapes_count が 2 であること"""
        from auto_office_editor_mcp.tools.info import get_ppt_slides_info

        result = json.loads(get_ppt_slides_info(PRES_NAME))
        assert result[2]["shapes_count"] == SLIDE3_SHAPE_COUNT

    def test_invalid_pres_name_returns_error(self, app_manager):
        """存在しないプレゼンテーション名はエラーを返す"""
        from auto_office_editor_mcp.tools.info import get_ppt_slides_info

        result = json.loads(get_ppt_slides_info("nonexistent.pptx"))
        assert "error" in result


# ===========================================================================
# get_ppt_shapes_info / 図形一覧取得
# ===========================================================================


class TestGetPptShapesInfo:
    def test_slide1_shape_count(self, ppt_open):
        """スライド1の図形数が2であること"""
        from auto_office_editor_mcp.tools.info import get_ppt_shapes_info

        result = json.loads(get_ppt_shapes_info(PRES_NAME, 1))
        assert len(result) == SLIDE1_SHAPE_COUNT

    def test_shape_has_required_keys(self, ppt_open):
        """各図形に必須キーが存在すること"""
        from auto_office_editor_mcp.tools.info import get_ppt_shapes_info

        result = json.loads(get_ppt_shapes_info(PRES_NAME, 1))
        required_keys = {
            "index",
            "name",
            "type",
            "has_text",
            "text",
            "left_cm",
            "top_cm",
            "width_cm",
            "height_cm",
            "has_table",
        }
        for shape in result:
            assert required_keys.issubset(shape.keys())

    def test_slide1_title_text(self, ppt_open):
        """スライド1のタイトルテキストが '四半期売上報告' であること"""
        from auto_office_editor_mcp.tools.info import get_ppt_shapes_info

        result = json.loads(get_ppt_shapes_info(PRES_NAME, 1))
        texts = [s["text"] for s in result if s["has_text"]]
        assert SLIDE1_TITLE in texts

    def test_slide1_subtitle_text(self, ppt_open):
        """スライド1のサブタイトルテキストが含まれること"""
        from auto_office_editor_mcp.tools.info import get_ppt_shapes_info

        result = json.loads(get_ppt_shapes_info(PRES_NAME, 1))
        texts = [s["text"] for s in result if s["has_text"]]
        assert SLIDE1_SUBTITLE in texts

    def test_slide2_has_table_shape(self, ppt_open):
        """スライド2に has_table=True の図形が存在すること"""
        from auto_office_editor_mcp.tools.info import get_ppt_shapes_info

        result = json.loads(get_ppt_shapes_info(PRES_NAME, 2))
        assert any(s["has_table"] for s in result)

    def test_slide3_textbox_has_text(self, ppt_open):
        """スライド3のテキストボックスが has_text=True であること"""
        from auto_office_editor_mcp.tools.info import get_ppt_shapes_info

        result = json.loads(get_ppt_shapes_info(PRES_NAME, 3))
        text_shapes = [s for s in result if s["has_text"] and s["text"]]
        assert any(SLIDE3_TEXTBOX_TEXT in s["text"] for s in text_shapes)

    def test_sample_rect_position(self, ppt_open):
        """SampleRect の left_cm が約 3.0cm であること"""
        from auto_office_editor_mcp.tools.info import get_ppt_shapes_info

        result = json.loads(get_ppt_shapes_info(PRES_NAME, 3))
        rect_shapes = [s for s in result if s["name"] == SAMPLE_RECT_NAME]
        assert len(rect_shapes) == 1
        assert abs(rect_shapes[0]["left_cm"] - SAMPLE_RECT_LEFT_CM) < 0.1

    def test_position_in_cm(self, ppt_open):
        """全図形の left_cm/top_cm が数値であること"""
        from auto_office_editor_mcp.tools.info import get_ppt_shapes_info

        result = json.loads(get_ppt_shapes_info(PRES_NAME, 1))
        for s in result:
            assert isinstance(s["left_cm"], (int, float))
            assert isinstance(s["top_cm"], (int, float))

    def test_size_in_cm(self, ppt_open):
        """全図形の width_cm/height_cm が正の数値であること"""
        from auto_office_editor_mcp.tools.info import get_ppt_shapes_info

        result = json.loads(get_ppt_shapes_info(PRES_NAME, 1))
        for s in result:
            assert s["width_cm"] > 0
            assert s["height_cm"] > 0

    def test_invalid_slide_index_returns_error(self, ppt_open):
        """存在しないスライドインデックスはエラーを返す"""
        from auto_office_editor_mcp.tools.info import get_ppt_shapes_info

        result = json.loads(get_ppt_shapes_info(PRES_NAME, 999))
        assert "error" in result


# ===========================================================================
# ppt_table_to_df : PPT 表 → DataFrame 変換
# ===========================================================================


class TestPptTableToDf:
    def _get_table_shape(self, mgr):
        """スライド2の表 Shape を返す。STA スレッド内で使うこと。"""
        pres = mgr.ppt_app.Presentations(PRES_NAME)
        slide2 = pres.Slides(2)
        for i in range(1, slide2.Shapes.Count + 1):
            sh = slide2.Shapes(i)
            try:
                if sh.HasTable:
                    return sh
            except Exception:
                pass
        raise RuntimeError("表が見つかりません")

    def test_shape(self, ppt_open):
        """スライド2の表を DataFrame に変換すると 2行3列（ヘッダー除く）"""
        from auto_office_editor_mcp.helpers import ppt_table_to_df

        df = ppt_open.run_in_sta(
            lambda: ppt_table_to_df(self._get_table_shape(ppt_open))
        )
        assert df.shape == (2, 3)

    def test_headers(self, ppt_open):
        """ヘッダー行が正しく取得される"""
        from auto_office_editor_mcp.helpers import ppt_table_to_df

        df = ppt_open.run_in_sta(
            lambda: ppt_table_to_df(self._get_table_shape(ppt_open))
        )
        assert list(df.columns) == TABLE_HEADERS

    def test_first_row_values(self, ppt_open):
        """1行目のデータが正しく取得される"""
        from auto_office_editor_mcp.helpers import ppt_table_to_df

        df = ppt_open.run_in_sta(
            lambda: ppt_table_to_df(self._get_table_shape(ppt_open))
        )
        assert list(df.iloc[0]) == TABLE_ROW1

    def test_second_row_values(self, ppt_open):
        """2行目のデータが正しく取得される"""
        from auto_office_editor_mcp.helpers import ppt_table_to_df

        df = ppt_open.run_in_sta(
            lambda: ppt_table_to_df(self._get_table_shape(ppt_open))
        )
        assert list(df.iloc[1]) == TABLE_ROW2

    def test_non_table_shape_raises(self, ppt_open):
        """HasTable=False の図形を渡すと ValueError が発生する"""
        from auto_office_editor_mcp.helpers import ppt_table_to_df

        def _do():
            pres = ppt_open.ppt_app.Presentations(PRES_NAME)
            # スライド1のタイトル（テキストボックス）を渡す
            title_shape = pres.Slides(1).Shapes(1)
            ppt_table_to_df(title_shape)

        with pytest.raises(ValueError):
            ppt_open.run_in_sta(_do)


# ===========================================================================
# df_to_ppt_table : DataFrame → PPT 表追加（編集テスト）
# ===========================================================================


class TestDfToPptTable:
    def test_roundtrip_shape(self, ppt_editable):
        """DataFrame → PPT 表 → DataFrame で行列数が一致する"""
        from auto_office_editor_mcp.helpers import df_to_ppt_table, ppt_table_to_df

        mgr, dst = ppt_editable
        df_src = pd.DataFrame(
            {"商品": ["りんご", "みかん"], "価格": ["150", "80"], "在庫": ["50", "120"]}
        )

        def _do():
            slide3 = _slide(mgr, dst, 3)
            shape = df_to_ppt_table(
                df_src, slide3, left_cm=1, top_cm=12, width_cm=18, height_cm=4
            )
            return ppt_table_to_df(shape)

        df_result = mgr.run_in_sta(_do)
        assert df_result.shape == df_src.shape

    def test_roundtrip_columns(self, ppt_editable):
        """往復後の列名が一致する"""
        from auto_office_editor_mcp.helpers import df_to_ppt_table, ppt_table_to_df

        mgr, dst = ppt_editable
        df_src = pd.DataFrame({"A": ["a1", "a2"], "B": ["b1", "b2"]})

        def _do():
            slide3 = _slide(mgr, dst, 3)
            shape = df_to_ppt_table(
                df_src, slide3, left_cm=1, top_cm=14, width_cm=10, height_cm=3
            )
            return ppt_table_to_df(shape)

        df_result = mgr.run_in_sta(_do)
        assert list(df_result.columns) == list(df_src.columns)

    def test_roundtrip_values(self, ppt_editable):
        """往復後のデータ値が一致する"""
        from auto_office_editor_mcp.helpers import df_to_ppt_table, ppt_table_to_df

        mgr, dst = ppt_editable
        df_src = pd.DataFrame({"名前": ["田中", "鈴木"], "点数": ["90", "75"]})

        def _do():
            slide3 = _slide(mgr, dst, 3)
            shape = df_to_ppt_table(
                df_src, slide3, left_cm=1, top_cm=12, width_cm=12, height_cm=3
            )
            return ppt_table_to_df(shape)

        df_result = mgr.run_in_sta(_do)
        assert list(df_result["名前"]) == ["田中", "鈴木"]
        assert list(df_result["点数"]) == ["90", "75"]

    def test_shape_count_increases(self, ppt_editable):
        """表を追加すると Shapes.Count が増える"""
        from auto_office_editor_mcp.helpers import df_to_ppt_table

        mgr, dst = ppt_editable
        df_src = pd.DataFrame({"X": ["1", "2"]})

        def _do():
            slide3 = _slide(mgr, dst, 3)
            count_before = slide3.Shapes.Count
            df_to_ppt_table(
                df_src, slide3, left_cm=1, top_cm=12, width_cm=8, height_cm=3
            )
            return count_before, slide3.Shapes.Count

        count_before, count_after = mgr.run_in_sta(_do)
        assert count_after == count_before + 1

    def test_position_params(self, ppt_editable):
        """left_cm/top_cm が図形の位置に反映されること（±0.5cm の許容誤差）"""
        from auto_office_editor_mcp.helpers import df_to_ppt_table, pt_to_cm

        mgr, dst = ppt_editable
        df_src = pd.DataFrame({"Col": ["v1"]})

        def _do():
            slide3 = _slide(mgr, dst, 3)
            shape = df_to_ppt_table(
                df_src, slide3, left_cm=4.0, top_cm=6.0, width_cm=10, height_cm=2
            )
            return pt_to_cm(shape.Left), pt_to_cm(shape.Top)

        left_cm, top_cm = mgr.run_in_sta(_do)
        assert abs(left_cm - 4.0) < 0.5
        assert abs(top_cm - 6.0) < 0.5


# ===========================================================================
# 図形の追加・編集・削除（スライド3の SampleRect を対象）
# ===========================================================================


class TestAddShape:
    def test_add_rectangle_count(self, ppt_editable):
        """長方形を追加すると Shapes.Count が増える"""
        from auto_office_editor_mcp.helpers import cm_to_pt

        mgr, dst = ppt_editable

        def _do():
            slide3 = _slide(mgr, dst, 3)
            count_before = slide3.Shapes.Count
            # msoShapeRectangle = 1
            slide3.Shapes.AddShape(
                1, cm_to_pt(1), cm_to_pt(1), cm_to_pt(5), cm_to_pt(3)
            )
            return count_before, slide3.Shapes.Count

        count_before, count_after = mgr.run_in_sta(_do)
        assert count_after == count_before + 1

    def test_add_shape_name(self, ppt_editable):
        """追加した図形の名前を設定できる"""
        from auto_office_editor_mcp.helpers import cm_to_pt

        mgr, dst = ppt_editable

        def _do():
            slide3 = _slide(mgr, dst, 3)
            shp = slide3.Shapes.AddShape(
                1, cm_to_pt(1), cm_to_pt(1), cm_to_pt(5), cm_to_pt(3)
            )
            shp.Name = "NewRect"
            return slide3.Shapes("NewRect").Name

        assert mgr.run_in_sta(_do) == "NewRect"

    def test_add_textbox(self, ppt_editable):
        """テキストボックスを追加できる"""
        from auto_office_editor_mcp.helpers import cm_to_pt

        mgr, dst = ppt_editable

        def _do():
            slide3 = _slide(mgr, dst, 3)
            count_before = slide3.Shapes.Count
            tbx = slide3.Shapes.AddTextbox(
                1, cm_to_pt(2), cm_to_pt(10), cm_to_pt(15), cm_to_pt(2)
            )
            tbx.TextFrame.TextRange.Text = "新しいテキストボックス"
            return (
                slide3.Shapes.Count,
                "新しいテキストボックス" in tbx.TextFrame.TextRange.Text,
            )

        count, has_text = mgr.run_in_sta(_do)
        assert count > 1
        assert has_text


class TestEditExistingShape:
    """既存の SampleRect の編集テスト。"""

    def test_sample_rect_exists(self, ppt_editable):
        """SampleRect が存在すること"""
        mgr, dst = ppt_editable
        shape = mgr.run_in_sta(lambda: _slide(mgr, dst, 3).Shapes(SAMPLE_RECT_NAME))
        assert shape is not None

    def test_fill_color_yellow(self, ppt_editable):
        """SampleRect の塗りつぶし色を黄色に設定できる"""
        mgr, dst = ppt_editable

        def _do():
            slide3 = _slide(mgr, dst, 3)
            slide3.Shapes(SAMPLE_RECT_NAME).Fill.ForeColor.RGB = ole_yellow()
            return slide3.Shapes(SAMPLE_RECT_NAME).Fill.ForeColor.RGB

        assert mgr.run_in_sta(_do) == ole_yellow()

    def test_fill_color_red(self, ppt_editable):
        """SampleRect の塗りつぶし色を赤に設定できる"""
        mgr, dst = ppt_editable

        def _do():
            slide3 = _slide(mgr, dst, 3)
            slide3.Shapes(SAMPLE_RECT_NAME).Fill.ForeColor.RGB = ole_red()
            return slide3.Shapes(SAMPLE_RECT_NAME).Fill.ForeColor.RGB

        assert mgr.run_in_sta(_do) == ole_red()

    def test_fill_color_blue(self, ppt_editable):
        """SampleRect の塗りつぶし色を青に設定できる"""
        mgr, dst = ppt_editable

        def _do():
            slide3 = _slide(mgr, dst, 3)
            slide3.Shapes(SAMPLE_RECT_NAME).Fill.ForeColor.RGB = ole_blue()
            return slide3.Shapes(SAMPLE_RECT_NAME).Fill.ForeColor.RGB

        assert mgr.run_in_sta(_do) == ole_blue()

    def test_move_shape(self, ppt_editable):
        """SampleRect の位置を変更できる"""
        from auto_office_editor_mcp.helpers import cm_to_pt

        mgr, dst = ppt_editable

        def _do():
            slide3 = _slide(mgr, dst, 3)
            slide3.Shapes(SAMPLE_RECT_NAME).Left = cm_to_pt(1.0)
            slide3.Shapes(SAMPLE_RECT_NAME).Top = cm_to_pt(2.0)
            return (
                slide3.Shapes(SAMPLE_RECT_NAME).Left,
                slide3.Shapes(SAMPLE_RECT_NAME).Top,
            )

        left, top = mgr.run_in_sta(_do)
        assert abs(left - cm_to_pt(1.0)) < 1.0
        assert abs(top - cm_to_pt(2.0)) < 1.0

    def test_resize_shape(self, ppt_editable):
        """SampleRect のサイズを変更できる"""
        from auto_office_editor_mcp.helpers import cm_to_pt

        mgr, dst = ppt_editable

        def _do():
            slide3 = _slide(mgr, dst, 3)
            slide3.Shapes(SAMPLE_RECT_NAME).Width = cm_to_pt(15.0)
            slide3.Shapes(SAMPLE_RECT_NAME).Height = cm_to_pt(6.0)
            return (
                slide3.Shapes(SAMPLE_RECT_NAME).Width,
                slide3.Shapes(SAMPLE_RECT_NAME).Height,
            )

        width, height = mgr.run_in_sta(_do)
        assert abs(width - cm_to_pt(15.0)) < 1.0
        assert abs(height - cm_to_pt(6.0)) < 1.0

    def test_line_color(self, ppt_editable):
        """SampleRect の枠線色を設定できる"""
        mgr, dst = ppt_editable

        def _do():
            slide3 = _slide(mgr, dst, 3)
            slide3.Shapes(SAMPLE_RECT_NAME).Line.ForeColor.RGB = ole_red()
            return slide3.Shapes(SAMPLE_RECT_NAME).Line.ForeColor.RGB

        assert mgr.run_in_sta(_do) == ole_red()

    def test_set_text(self, ppt_editable):
        """SampleRect にテキストを設定できる"""
        mgr, dst = ppt_editable

        def _do():
            slide3 = _slide(mgr, dst, 3)
            slide3.Shapes(SAMPLE_RECT_NAME).TextFrame.TextRange.Text = "編集済み"
            return slide3.Shapes(SAMPLE_RECT_NAME).TextFrame.TextRange.Text

        assert "編集済み" in mgr.run_in_sta(_do)


class TestDeleteShape:
    def test_delete_restores_count(self, ppt_editable):
        """図形を追加して削除すると元の Shapes.Count に戻る"""
        from auto_office_editor_mcp.helpers import cm_to_pt

        mgr, dst = ppt_editable

        def _do():
            slide3 = _slide(mgr, dst, 3)
            count_before = slide3.Shapes.Count
            shp = slide3.Shapes.AddShape(
                1, cm_to_pt(1), cm_to_pt(1), cm_to_pt(3), cm_to_pt(2)
            )
            shp.Name = "TempRect"
            count_added = slide3.Shapes.Count
            slide3.Shapes("TempRect").Delete()
            return count_before, count_added, slide3.Shapes.Count

        count_before, count_added, count_after = mgr.run_in_sta(_do)
        assert count_added == count_before + 1
        assert count_after == count_before

    def test_delete_sample_rect(self, ppt_editable):
        """SampleRect を削除するとスライド3の図形数が1つ減る"""
        mgr, dst = ppt_editable

        def _do():
            slide3 = _slide(mgr, dst, 3)
            count_before = slide3.Shapes.Count
            slide3.Shapes(SAMPLE_RECT_NAME).Delete()
            return count_before, slide3.Shapes.Count

        count_before, count_after = mgr.run_in_sta(_do)
        assert count_after == count_before - 1


# ===========================================================================
# execute_office_script 経由の PowerPoint 操作
# ===========================================================================


class TestExecuteOfficeScriptPowerPoint:
    def test_read_slide_title(self, ppt_open):
        """execute_script でスライド1のタイトルを読み取れる"""
        from auto_office_editor_mcp.executor import execute_script

        code = (
            f"pres = ppt_app.Presentations('{PRES_NAME}')\n"
            "slide = pres.Slides(1)\n"
            "print(slide.Shapes.Title.TextFrame.TextRange.Text)"
        )
        result = execute_script(code)
        assert SLIDE1_TITLE in result

    def test_change_title_text(self, ppt_editable):
        """execute_script でスライドタイトルを変更できる"""
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = ppt_editable
        pres_name = _pres_name(dst)

        code = (
            f"pres = ppt_app.Presentations('{pres_name}')\n"
            "slide = pres.Slides(1)\n"
            "slide.Shapes.Title.TextFrame.TextRange.Text = '変更後タイトル'\n"
            "print(slide.Shapes.Title.TextFrame.TextRange.Text)"
        )
        result = execute_script(code)
        assert "変更後タイトル" in result

    def test_add_shape_via_script(self, ppt_editable):
        """execute_script で図形を追加できる"""
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = ppt_editable
        pres_name = _pres_name(dst)
        count_before = mgr.run_in_sta(lambda: _pres(mgr, dst).Slides(3).Shapes.Count)

        code = (
            f"pres = ppt_app.Presentations('{pres_name}')\n"
            "slide = pres.Slides(3)\n"
            "shp = slide.Shapes.AddShape(1, cm_to_pt(1), cm_to_pt(8), cm_to_pt(6), cm_to_pt(3))\n"
            "shp.Name = 'ScriptRect'\n"
            "print(slide.Shapes.Count)"
        )
        result = execute_script(code)
        # result は "実行成功\n\n--- 出力 ---\n<数値>" の形式
        assert str(count_before + 1) in result

    def test_shape_fill_color_via_script(self, ppt_editable):
        """execute_script で図形の塗りつぶし色を設定できる"""
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = ppt_editable
        pres_name = _pres_name(dst)

        code = (
            f"pres = ppt_app.Presentations('{pres_name}')\n"
            "slide = pres.Slides(3)\n"
            f"shp = slide.Shapes('{SAMPLE_RECT_NAME}')\n"
            "shp.Fill.ForeColor.RGB = rgb(255, 255, 0)\n"
            "print(shp.Fill.ForeColor.RGB)"
        )
        result = execute_script(code)
        assert str(ole_yellow()) in result

    def test_shape_position_via_script(self, ppt_editable):
        """execute_script で図形の位置を変更できる"""
        from auto_office_editor_mcp.executor import execute_script
        from auto_office_editor_mcp.helpers import cm_to_pt

        mgr, dst = ppt_editable
        pres_name = _pres_name(dst)

        code = (
            f"pres = ppt_app.Presentations('{pres_name}')\n"
            "slide = pres.Slides(3)\n"
            f"shp = slide.Shapes('{SAMPLE_RECT_NAME}')\n"
            "shp.Left = cm_to_pt(2.0)\n"
            "shp.Top = cm_to_pt(4.0)\n"
            "print(round(shp.Left, 1))"
        )
        result = execute_script(code)
        assert str(round(cm_to_pt(2.0), 1)) in result

    def test_delete_shape_via_script(self, ppt_editable):
        """execute_script で図形を削除できる"""
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = ppt_editable
        pres_name = _pres_name(dst)
        count_before = mgr.run_in_sta(lambda: _pres(mgr, dst).Slides(3).Shapes.Count)

        code = (
            f"pres = ppt_app.Presentations('{pres_name}')\n"
            "slide = pres.Slides(3)\n"
            f"slide.Shapes('{SAMPLE_RECT_NAME}').Delete()\n"
            "print(slide.Shapes.Count)"
        )
        result = execute_script(code)
        # result は "実行成功\n\n--- 出力 ---\n<数値>" の形式
        assert str(count_before - 1) in result


# ===========================================================================
# 複合操作テスト
# ===========================================================================


class TestCompositeEditPowerPoint:
    def test_full_slide_edit(self, ppt_editable):
        """
        スライド3に複数の図形追加・色設定・テキスト設定を一括で行う複合テスト。
        execute_script で全操作を実行し、結果を確認する。
        """
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = ppt_editable
        pres_name = _pres_name(dst)
        count_before = mgr.run_in_sta(lambda: _pres(mgr, dst).Slides(3).Shapes.Count)

        code = (
            f"pres = ppt_app.Presentations('{pres_name}')\n"
            "slide = pres.Slides(3)\n"
            # 楕円を追加（msoShapeOval = 9）
            "oval = slide.Shapes.AddShape(9, cm_to_pt(2), cm_to_pt(8), cm_to_pt(6), cm_to_pt(4))\n"
            "oval.Name = 'CompositeOval'\n"
            "oval.Fill.ForeColor.RGB = rgb(255, 255, 0)\n"
            "oval.Line.ForeColor.RGB = rgb(0, 0, 0)\n"
            "oval.TextFrame.TextRange.Text = '楕円'\n"
            # SampleRect の色も変更
            f"rect = slide.Shapes('{SAMPLE_RECT_NAME}')\n"
            "rect.Fill.ForeColor.RGB = rgb(0, 0, 255)\n"
            "print(slide.Shapes.Count)\n"
            "print(oval.Fill.ForeColor.RGB)\n"
            "print(rect.Fill.ForeColor.RGB)\n"
            "print(oval.TextFrame.TextRange.Text)"
        )
        result = execute_script(code)

        assert str(count_before + 1) in result
        assert str(ole_yellow()) in result
        assert str(ole_blue()) in result
        assert "楕円" in result

    def test_add_table_to_new_slide_via_script(self, ppt_editable):
        """
        execute_script で新しいスライドを追加し、表を挿入する複合テスト。
        """
        from auto_office_editor_mcp.executor import execute_script

        mgr, dst = ppt_editable
        pres_name = _pres_name(dst)
        slides_before = mgr.run_in_sta(lambda: _pres(mgr, dst).Slides.Count)

        code = (
            f"pres = ppt_app.Presentations('{pres_name}')\n"
            # 空白スライドを末尾に追加（ppLayoutBlank = 12）
            "new_slide = pres.Slides.Add(pres.Slides.Count + 1, 12)\n"
            # 表を追加
            "tbl_shape = new_slide.Shapes.AddTable(3, 3, cm_to_pt(2), cm_to_pt(3), cm_to_pt(20), cm_to_pt(6))\n"
            "tbl = tbl_shape.Table\n"
            "tbl.Cell(1,1).Shape.TextFrame.TextRange.Text = '列A'\n"
            "tbl.Cell(1,2).Shape.TextFrame.TextRange.Text = '列B'\n"
            "tbl.Cell(1,3).Shape.TextFrame.TextRange.Text = '列C'\n"
            "tbl.Cell(2,1).Shape.TextFrame.TextRange.Text = 'データ1'\n"
            "print(pres.Slides.Count)\n"
            "print(tbl.Cell(1,1).Shape.TextFrame.TextRange.Text)"
        )
        result = execute_script(code)

        assert str(slides_before + 1) in result
        assert "列A" in result
