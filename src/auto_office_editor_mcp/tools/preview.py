"""
Word / PowerPoint のページ・スライドを画像として返すプレビューツール。

処理フロー:
    1. win32com で対象ファイルを PDF にエクスポート（一時ファイル）← STA ワーカー内
    2. PyMuPDF で指定ページを PNG に変換                         ← STA ワーカー外
    3. PNG バイト列のリストで返す（server.py で FastMCP の Image 型に変換）
    4. 一時 PDF ファイルを削除

ページ指定:
    - None        → 全ページ（20ページ超はエラー）
    - int         → 単一ページ（1始まり）
    - list[int]   → 複数ページ指定（1始まり）
"""

from __future__ import annotations

import tempfile
from pathlib import Path

import fitz  # PyMuPDF

from ..app_manager import manager

# ページ数の上限（未指定時に全ページ変換する場合の上限）
MAX_PAGES = 20

# PDF エクスポート用の win32com 定数
_WD_EXPORT_FORMAT_PDF = 17  # wdExportFormatPDF
_PP_SAVE_AS_PDF = 32  # ppSaveAsPDF（SaveAs の FileFormat 引数）


def _resolve_pages(
    total: int,
    pages: int | list[int] | None,
    label: str,
) -> list[int]:
    """
    ページ指定を 0始まりのインデックスリストに正規化する。

    Args:
        total:  総ページ数
        pages:  None / int / list[int]（いずれも 1始まり）
        label:  エラーメッセージ用のラベル（"ページ" or "スライド"）

    Returns:
        0始まりのインデックスリスト
    """
    if pages is None:
        if total > MAX_PAGES:
            raise ValueError(
                f"{label}数が {total} あります。"
                f"{MAX_PAGES} {label}以下のファイルか、{label}番号を指定してください。"
            )
        return list(range(total))

    if isinstance(pages, int):
        pages = [pages]

    result = []
    for p in pages:
        if p < 1 or p > total:
            raise ValueError(f"{label}番号 {p} は範囲外です（1〜{total}）。")
        result.append(p - 1)  # 0始まりに変換
    return result


def _pdf_pages_to_images(
    pdf_path: str,
    page_indices: list[int],
    dpi: int,
) -> list[tuple[int, bytes]]:
    """
    PDF の指定ページを PNG バイト列に変換する。

    Returns:
        [(page_number_1based, png_bytes), ...]
    """
    doc = fitz.open(pdf_path)
    results = []
    zoom = dpi / 72.0  # PyMuPDF のデフォルトは 72 DPI
    mat = fitz.Matrix(zoom, zoom)

    try:
        for idx in page_indices:
            page = doc[idx]
            pix = page.get_pixmap(matrix=mat, alpha=False)
            png_bytes: bytes = pix.tobytes("png")
            results.append((idx + 1, png_bytes))
    finally:
        doc.close()

    return results


def get_word_page_image(
    doc_name: str,
    pages: int | list[int] | None = None,
    dpi: int = 150,
) -> list[tuple[int, bytes]]:
    """
    Word ドキュメントの指定ページを PNG バイト列に変換して返す。

    Args:
        doc_name: ドキュメント名（例: "report.docx"）
        pages:    ページ番号（1始まり）。
                  None → 全ページ（20ページ超はエラー）
                  int  → 単一ページ
                  list[int] → 複数ページ
        dpi:      出力解像度（デフォルト: 150）

    Returns:
        [(page_number, png_bytes), ...]
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = str(Path(tmpdir) / "export.pdf")

        # Word → PDF エクスポート（COM 操作は STA ワーカーで実行）
        def _export():
            wd_app = manager.wd_app
            if wd_app is None:
                raise RuntimeError("Word が起動していません。")
            try:
                doc = wd_app.Documents(doc_name)
            except Exception:
                raise FileNotFoundError(f"'{doc_name}' が開いていません。")
            doc.ExportAsFixedFormat(
                OutputFileName=pdf_path,
                ExportFormat=_WD_EXPORT_FORMAT_PDF,
            )

        manager.run_in_sta(_export)

        # 総ページ数を PDF から取得（Word の BuiltinDocumentProperties より正確）
        pdf_doc = fitz.open(pdf_path)
        total_pages = pdf_doc.page_count
        pdf_doc.close()

        page_indices = _resolve_pages(total_pages, pages, "ページ")
        return _pdf_pages_to_images(pdf_path, page_indices, dpi)


def get_ppt_slide_image(
    pres_name: str,
    pages: int | list[int] | None = None,
    dpi: int = 150,
) -> list[tuple[int, bytes]]:
    """
    PowerPoint プレゼンテーションの指定スライドを PNG バイト列に変換して返す。

    Args:
        pres_name: プレゼンテーション名（例: "slide.pptx"）
        pages:     スライド番号（1始まり）。
                   None → 全スライド（20枚超はエラー）
                   int  → 単一スライド
                   list[int] → 複数スライド
        dpi:       出力解像度（デフォルト: 150）

    Returns:
        [(slide_number, png_bytes), ...]
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = str(Path(tmpdir) / "export.pdf")

        # PowerPoint → PDF エクスポート（COM 操作は STA ワーカーで実行）
        def _export():
            ppt_app = manager.ppt_app
            if ppt_app is None:
                raise RuntimeError("PowerPoint が起動していません。")
            try:
                pres = ppt_app.Presentations(pres_name)
            except Exception:
                raise FileNotFoundError(f"'{pres_name}' が開いていません。")
            total = pres.Slides.Count
            # PowerPoint → PDF エクスポート（ppSaveAsPDF = 32）
            pres.SaveAs(pdf_path, _PP_SAVE_AS_PDF)
            return total

        total_slides = manager.run_in_sta(_export)

        page_indices = _resolve_pages(total_slides, pages, "スライド")
        return _pdf_pages_to_images(pdf_path, page_indices, dpi)
