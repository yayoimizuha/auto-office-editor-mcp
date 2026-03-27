"""
preview 統合テスト。

get_word_page_image / get_ppt_slide_image の
PDF エクスポート → PNG 変換フローを end-to-end で検証する。

テスト対象:
    src/auto_office_editor_mcp/tools/preview.py
        get_word_page_image(doc_name, pages, dpi)
        get_ppt_slide_image(pres_name, pages, dpi)

検証する観点:
    1. 返り値の構造  : list[tuple[int, bytes]] であること
    2. ページ番号    : タプルの第1要素が 1始まりの正しい番号であること
    3. PNG バイト列  : PNG シグネチャ（\\x89PNG）で始まること
    4. 画像サイズ    : PyMuPDF で開けて width/height が正値であること
    5. ページ指定    : None / int / list[int] の各指定が正しく機能すること
    6. DPI の反映    : dpi=300 は dpi=75 より大きい画像になること
    7. エラー系      : 未開封ファイル・範囲外ページで適切な例外が発生すること
    8. ファイル保存  : PDF・PNG を fixtures/editable/ に書き出して目視確認できること

test.docx の構成（pages=1 の画像が返れば十分）:
    1ページ以上のドキュメント（見出し・段落・表を含む）

test.pptx の構成:
    3スライド
"""

from __future__ import annotations

from pathlib import Path

import pytest

# editable ディレクトリ（Word/PPT テストと共用）
EDITABLE_DIR = Path(__file__).parent / "fixtures" / "editable"

pytestmark = pytest.mark.integration

# PNG ファイルシグネチャ（先頭8バイト）
PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"

DOC_NAME = "test.docx"
PRES_NAME = "test.pptx"
EXPECTED_SLIDE_COUNT = 3


# ===========================================================================
# ヘルパー
# ===========================================================================


def _is_valid_png(data: bytes) -> bool:
    """PNG シグネチャで始まるかチェック。"""
    return data[:8] == PNG_SIGNATURE


def _image_size(data: bytes) -> tuple[int, int]:
    """PyMuPDF で PNG バイト列を開いて (width, height) を返す。"""
    import fitz

    doc = fitz.open(stream=data, filetype="png")
    page = doc[0]
    w = int(page.rect.width)
    h = int(page.rect.height)
    doc.close()
    return w, h


# ===========================================================================
# Word プレビュー
# ===========================================================================


class TestGetWordPageImage:
    """get_word_page_image の統合テスト。"""

    # ------------------------------------------------------------------
    # 返り値の構造
    # ------------------------------------------------------------------

    def test_returns_list(self, wd_open):
        """返り値がリストであること"""
        from auto_office_editor_mcp.tools.preview import get_word_page_image

        result = get_word_page_image(DOC_NAME, pages=1)
        assert isinstance(result, list)

    def test_single_page_returns_one_item(self, wd_open):
        """単一ページ指定で要素数が1であること"""
        from auto_office_editor_mcp.tools.preview import get_word_page_image

        result = get_word_page_image(DOC_NAME, pages=1)
        assert len(result) == 1

    def test_item_is_tuple(self, wd_open):
        """各要素が (page_number, bytes) のタプルであること"""
        from auto_office_editor_mcp.tools.preview import get_word_page_image

        result = get_word_page_image(DOC_NAME, pages=1)
        page_num, png_bytes = result[0]
        assert isinstance(page_num, int)
        assert isinstance(png_bytes, bytes)

    # ------------------------------------------------------------------
    # ページ番号
    # ------------------------------------------------------------------

    def test_page_number_is_1based(self, wd_open):
        """ページ番号が 1始まりであること"""
        from auto_office_editor_mcp.tools.preview import get_word_page_image

        result = get_word_page_image(DOC_NAME, pages=1)
        page_num, _ = result[0]
        assert page_num == 1

    def test_list_page_numbers(self, wd_open):
        """list 指定時に正しいページ番号が返ること"""
        from auto_office_editor_mcp.tools.preview import get_word_page_image

        result = get_word_page_image(DOC_NAME, pages=[1])
        assert result[0][0] == 1

    # ------------------------------------------------------------------
    # PNG バイト列の妥当性
    # ------------------------------------------------------------------

    def test_bytes_is_valid_png(self, wd_open):
        """返されたバイト列が PNG 形式であること"""
        from auto_office_editor_mcp.tools.preview import get_word_page_image

        _, png_bytes = get_word_page_image(DOC_NAME, pages=1)[0]
        assert _is_valid_png(png_bytes)

    def test_png_is_openable(self, wd_open):
        """PyMuPDF で開けること（width > 0 かつ height > 0）"""
        from auto_office_editor_mcp.tools.preview import get_word_page_image

        _, png_bytes = get_word_page_image(DOC_NAME, pages=1)[0]
        w, h = _image_size(png_bytes)
        assert w > 0
        assert h > 0

    def test_png_is_not_empty(self, wd_open):
        """PNG バイト列が最低限のサイズ（1KB 以上）であること"""
        from auto_office_editor_mcp.tools.preview import get_word_page_image

        _, png_bytes = get_word_page_image(DOC_NAME, pages=1)[0]
        assert len(png_bytes) > 1024

    # ------------------------------------------------------------------
    # DPI の反映
    # ------------------------------------------------------------------

    def test_higher_dpi_produces_larger_image(self, wd_open):
        """dpi=300 の画像は dpi=75 より大きいこと"""
        from auto_office_editor_mcp.tools.preview import get_word_page_image

        _, low_dpi = get_word_page_image(DOC_NAME, pages=1, dpi=75)[0]
        _, high_dpi = get_word_page_image(DOC_NAME, pages=1, dpi=300)[0]

        w_low, h_low = _image_size(low_dpi)
        w_high, h_high = _image_size(high_dpi)

        assert w_high > w_low
        assert h_high > h_low

    # ------------------------------------------------------------------
    # ページ指定のバリエーション
    # ------------------------------------------------------------------

    def test_none_returns_all_pages(self, wd_open):
        """pages=None で全ページが返ること（20ページ以下なら OK）"""
        from auto_office_editor_mcp.tools.preview import get_word_page_image

        result = get_word_page_image(DOC_NAME, pages=None)
        assert len(result) >= 1

    def test_list_spec_returns_correct_count(self, wd_open):
        """list 指定で指定した枚数分だけ返ること"""
        from auto_office_editor_mcp.tools.preview import get_word_page_image

        result = get_word_page_image(DOC_NAME, pages=[1])
        assert len(result) == 1

    # ------------------------------------------------------------------
    # エラー系
    # ------------------------------------------------------------------

    def test_not_opened_doc_raises(self, app_manager):
        """開いていないドキュメント名で FileNotFoundError が発生すること"""
        from auto_office_editor_mcp.tools.preview import get_word_page_image

        with pytest.raises(FileNotFoundError):
            get_word_page_image("nonexistent.docx", pages=1)

    def test_out_of_range_page_raises(self, wd_open):
        """存在しないページ番号で ValueError が発生すること"""
        from auto_office_editor_mcp.tools.preview import get_word_page_image

        with pytest.raises(ValueError, match="範囲外"):
            get_word_page_image(DOC_NAME, pages=9999)

    def test_zero_page_raises(self, wd_open):
        """0ページ指定（1始まりに反する）で ValueError が発生すること"""
        from auto_office_editor_mcp.tools.preview import get_word_page_image

        with pytest.raises(ValueError, match="範囲外"):
            get_word_page_image(DOC_NAME, pages=0)


# ===========================================================================
# PowerPoint プレビュー
# ===========================================================================


class TestGetPptSlideImage:
    """get_ppt_slide_image の統合テスト。"""

    # ------------------------------------------------------------------
    # 返り値の構造
    # ------------------------------------------------------------------

    def test_returns_list(self, ppt_open):
        """返り値がリストであること"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        result = get_ppt_slide_image(PRES_NAME, pages=1)
        assert isinstance(result, list)

    def test_single_slide_returns_one_item(self, ppt_open):
        """単一スライド指定で要素数が1であること"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        result = get_ppt_slide_image(PRES_NAME, pages=1)
        assert len(result) == 1

    def test_item_is_tuple(self, ppt_open):
        """各要素が (slide_number, bytes) のタプルであること"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        result = get_ppt_slide_image(PRES_NAME, pages=1)
        slide_num, png_bytes = result[0]
        assert isinstance(slide_num, int)
        assert isinstance(png_bytes, bytes)

    # ------------------------------------------------------------------
    # スライド番号
    # ------------------------------------------------------------------

    def test_slide_number_is_1based(self, ppt_open):
        """スライド番号が 1始まりであること"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        result = get_ppt_slide_image(PRES_NAME, pages=1)
        slide_num, _ = result[0]
        assert slide_num == 1

    def test_slide2_number(self, ppt_open):
        """スライド2を指定すると番号が2で返ること"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        result = get_ppt_slide_image(PRES_NAME, pages=2)
        assert result[0][0] == 2

    def test_list_slide_numbers_preserved(self, ppt_open):
        """list 指定時に各スライド番号が正しく保持されること"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        result = get_ppt_slide_image(PRES_NAME, pages=[1, 3])
        numbers = [r[0] for r in result]
        assert numbers == [1, 3]

    # ------------------------------------------------------------------
    # PNG バイト列の妥当性
    # ------------------------------------------------------------------

    def test_bytes_is_valid_png(self, ppt_open):
        """返されたバイト列が PNG 形式であること"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        _, png_bytes = get_ppt_slide_image(PRES_NAME, pages=1)[0]
        assert _is_valid_png(png_bytes)

    def test_png_is_openable(self, ppt_open):
        """PyMuPDF で開けること（width > 0 かつ height > 0）"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        _, png_bytes = get_ppt_slide_image(PRES_NAME, pages=1)[0]
        w, h = _image_size(png_bytes)
        assert w > 0
        assert h > 0

    def test_png_is_not_empty(self, ppt_open):
        """PNG バイト列が最低限のサイズ（1KB 以上）であること"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        _, png_bytes = get_ppt_slide_image(PRES_NAME, pages=1)[0]
        assert len(png_bytes) > 1024

    # ------------------------------------------------------------------
    # DPI の反映
    # ------------------------------------------------------------------

    def test_higher_dpi_produces_larger_image(self, ppt_open):
        """dpi=300 の画像は dpi=75 より大きいこと"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        _, low_dpi = get_ppt_slide_image(PRES_NAME, pages=1, dpi=75)[0]
        _, high_dpi = get_ppt_slide_image(PRES_NAME, pages=1, dpi=300)[0]

        w_low, h_low = _image_size(low_dpi)
        w_high, h_high = _image_size(high_dpi)

        assert w_high > w_low
        assert h_high > h_low

    # ------------------------------------------------------------------
    # スライド指定のバリエーション
    # ------------------------------------------------------------------

    def test_none_returns_all_slides(self, ppt_open):
        """pages=None で全スライドが返ること"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        result = get_ppt_slide_image(PRES_NAME, pages=None)
        assert len(result) == EXPECTED_SLIDE_COUNT

    def test_all_slides_are_valid_png(self, ppt_open):
        """全スライドが有効な PNG であること"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        result = get_ppt_slide_image(PRES_NAME, pages=None)
        for _, png_bytes in result:
            assert _is_valid_png(png_bytes)

    def test_list_returns_correct_count(self, ppt_open):
        """list 指定で指定した枚数分だけ返ること"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        result = get_ppt_slide_image(PRES_NAME, pages=[1, 2])
        assert len(result) == 2

    def test_single_last_slide(self, ppt_open):
        """最終スライドを単体で指定できること"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        result = get_ppt_slide_image(PRES_NAME, pages=EXPECTED_SLIDE_COUNT)
        assert len(result) == 1
        assert result[0][0] == EXPECTED_SLIDE_COUNT

    # ------------------------------------------------------------------
    # スライドごとに異なる画像が生成されること
    # ------------------------------------------------------------------

    def test_different_slides_produce_different_images(self, ppt_open):
        """スライド1とスライド3の PNG バイト列が異なること（中身が違う）"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        _, png_slide1 = get_ppt_slide_image(PRES_NAME, pages=1)[0]
        _, png_slide3 = get_ppt_slide_image(PRES_NAME, pages=3)[0]
        assert png_slide1 != png_slide3

    # ------------------------------------------------------------------
    # エラー系
    # ------------------------------------------------------------------

    def test_not_opened_pres_raises(self, app_manager):
        """開いていないプレゼンテーション名で FileNotFoundError が発生すること"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        with pytest.raises(FileNotFoundError):
            get_ppt_slide_image("nonexistent.pptx", pages=1)

    def test_out_of_range_slide_raises(self, ppt_open):
        """存在しないスライド番号で ValueError が発生すること"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        with pytest.raises(ValueError, match="範囲外"):
            get_ppt_slide_image(PRES_NAME, pages=9999)

    def test_zero_slide_raises(self, ppt_open):
        """0スライド指定（1始まりに反する）で ValueError が発生すること"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        with pytest.raises(ValueError, match="範囲外"):
            get_ppt_slide_image(PRES_NAME, pages=0)


# ===========================================================================
# PDF・PNG を editable/ に保存して目視確認できるテスト
# ===========================================================================


def _save_results(results: list[tuple[int, bytes]], stem: str) -> None:
    """
    get_*_page_image の返り値を fixtures/editable/ に PNG として書き出す。

    ファイル名: <stem>_page<N>.png  （N は 1始まりのページ/スライド番号）
    """
    EDITABLE_DIR.mkdir(exist_ok=True)
    for page_num, png_bytes in results:
        out = EDITABLE_DIR / f"{stem}_page{page_num}.png"
        out.write_bytes(png_bytes)


class TestSaveWordPdf:
    """Word の全ページを PNG に変換して editable/ に保存する。"""

    def test_save_all_pages(self, wd_open):
        """test.docx の全ページを PNG に変換して editable/ に書き出す"""
        from auto_office_editor_mcp.tools.preview import get_word_page_image

        results = get_word_page_image(DOC_NAME, pages=None, dpi=150)
        _save_results(results, "word_test")

        # 保存されたファイルが存在し、サイズが 0 より大きいこと
        for page_num, _ in results:
            saved = EDITABLE_DIR / f"word_test_page{page_num}.png"
            assert saved.exists()
            assert saved.stat().st_size > 0

    def test_save_page1_high_dpi(self, wd_open):
        """test.docx の 1ページ目を dpi=300 で PNG に変換して editable/ に書き出す"""
        from auto_office_editor_mcp.tools.preview import get_word_page_image

        results = get_word_page_image(DOC_NAME, pages=1, dpi=300)
        _save_results(results, "word_test_300dpi")

        saved = EDITABLE_DIR / "word_test_300dpi_page1.png"
        assert saved.exists()
        assert saved.stat().st_size > 0


class TestSavePptPdf:
    """PowerPoint の全スライドを PNG に変換して editable/ に保存する。"""

    def test_save_all_slides(self, ppt_open):
        """test.pptx の全スライドを PNG に変換して editable/ に書き出す"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        results = get_ppt_slide_image(PRES_NAME, pages=None, dpi=150)
        _save_results(results, "ppt_test")

        assert len(results) == EXPECTED_SLIDE_COUNT
        for page_num, _ in results:
            saved = EDITABLE_DIR / f"ppt_test_page{page_num}.png"
            assert saved.exists()
            assert saved.stat().st_size > 0

    def test_save_slide1_high_dpi(self, ppt_open):
        """test.pptx のスライド1を dpi=300 で PNG に変換して editable/ に書き出す"""
        from auto_office_editor_mcp.tools.preview import get_ppt_slide_image

        results = get_ppt_slide_image(PRES_NAME, pages=1, dpi=300)
        _save_results(results, "ppt_test_300dpi")

        saved = EDITABLE_DIR / "ppt_test_300dpi_page1.png"
        assert saved.exists()
        assert saved.stat().st_size > 0
