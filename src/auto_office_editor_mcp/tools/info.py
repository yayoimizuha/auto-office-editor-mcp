"""
Office オブジェクトの情報取得ツール（LLM の「目」）。

LLM がコードを書く前に「今どこに何があるか」を把握するための
情報取得関数群。すべて JSON 文字列を返す。

全ての COM 操作は manager.run_in_sta() 経由で STA ワーカースレッドに委譲する。
"""

from __future__ import annotations

import json
from typing import Callable, Any

from ..app_manager import manager
from ..helpers import pt_to_cm


def _run_sta_as_json(func: Callable[[], Any]) -> str:
    """func を STA スレッドで実行し、結果を JSON 文字列として返す。"""
    return json.dumps(manager.run_in_sta(func), ensure_ascii=False)


# ------------------------------------------------------------------
# Excel
# ------------------------------------------------------------------


def get_xl_sheets_info(wb_name: str) -> str:
    """
    Excel ワークブックのシート一覧を返す。

    Args:
        wb_name: ワークブック名（例: "hoge.xlsx"）

    Returns:
        JSON 文字列:
        [
            {"index": 1, "name": "Sheet1", "visible": true},
            ...
        ]
    """

    def _impl():
        xl_app = manager.xl_app
        if xl_app is None:
            return {"error": "Excel が起動していません。"}

        try:
            wb = xl_app.Workbooks(wb_name)
        except Exception:
            return {"error": f"'{wb_name}' が開いていません。"}

        sheets = []
        for i in range(1, wb.Sheets.Count + 1):
            sh = wb.Sheets(i)
            sheets.append(
                {
                    "index": i,
                    "name": sh.Name,
                    "visible": sh.Visible == -1,  # xlSheetVisible = -1
                }
            )
        return sheets

    return _run_sta_as_json(_impl)


def get_xl_range_info(wb_name: str, sheet_name: str, range_address: str) -> str:
    """
    Excel の指定セル範囲の情報を返す。

    Args:
        wb_name: ワークブック名（例: "hoge.xlsx"）
        sheet_name: シート名（例: "Sheet1"）
        range_address: セル範囲（例: "A1:D10"、"A1" など）

    Returns:
        JSON 文字列:
        {
            "address": "$A$1:$D$10",
            "rows": 10,
            "columns": 4,
            "values": [[...], ...]   # 先頭 20 行まで
        }
    """

    def _impl():
        xl_app = manager.xl_app
        if xl_app is None:
            return {"error": "Excel が起動していません。"}

        try:
            wb = xl_app.Workbooks(wb_name)
            ws = wb.Worksheets(sheet_name)
            rng = ws.Range(range_address)
        except Exception as e:
            return {"error": str(e)}

        raw = rng.Value
        if raw is None:
            values = []
        elif not isinstance(raw, tuple):
            values = [[raw]]
        elif raw and not isinstance(raw[0], tuple):
            values = [list(raw)]
        else:
            values = [list(row) for row in raw]

        # COM が返す datetime は Python の datetime に変換できないことがあるので文字列化
        def _safe(v):
            if v is None:
                return None
            try:
                return v if isinstance(v, (int, float, str, bool)) else str(v)
            except Exception:
                return str(v)

        safe_values = [[_safe(v) for v in row] for row in values[:20]]

        return {
            "address": rng.Address,
            "rows": rng.Rows.Count,
            "columns": rng.Columns.Count,
            "values": safe_values,
            "note": "values は先頭 20 行のみ表示しています。"
            if len(values) > 20
            else "",
        }

    return _run_sta_as_json(_impl)


# ------------------------------------------------------------------
# PowerPoint
# ------------------------------------------------------------------


def get_ppt_slides_info(pres_name: str) -> str:
    """
    PowerPoint プレゼンテーションのスライド一覧を返す。

    Args:
        pres_name: プレゼンテーション名（例: "slide.pptx"）

    Returns:
        JSON 文字列:
        [
            {"index": 1, "name": "スライド1", "layout": "タイトルスライド", "shapes_count": 3},
            ...
        ]
    """

    def _impl():
        ppt_app = manager.ppt_app
        if ppt_app is None:
            return {"error": "PowerPoint が起動していません。"}

        try:
            pres = ppt_app.Presentations(pres_name)
        except Exception:
            return {"error": f"'{pres_name}' が開いていません。"}

        slides = []
        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(i)
            try:
                layout_name = slide.Layout  # 数値
            except Exception:
                layout_name = None
            slides.append(
                {
                    "index": i,
                    "name": slide.Name,
                    "layout": layout_name,
                    "shapes_count": slide.Shapes.Count,
                }
            )
        return slides

    return _run_sta_as_json(_impl)


def get_ppt_shapes_info(pres_name: str, slide_index: int) -> str:
    """
    PowerPoint スライド上の全図形の情報を返す。LLM が「どの図形を操作するか」
    を判断するための認知ツール。

    Args:
        pres_name: プレゼンテーション名（例: "slide.pptx"）
        slide_index: スライド番号（1始まり）

    Returns:
        JSON 文字列:
        [
            {
                "index": 1,
                "name": "TextBox 1",
                "type": 17,
                "has_text": true,
                "text": "売上推移",
                "left_cm": 2.5,
                "top_cm": 1.2,
                "width_cm": 10.0,
                "height_cm": 2.0,
                "has_table": false
            },
            ...
        ]
    """

    def _impl():
        ppt_app = manager.ppt_app
        if ppt_app is None:
            return {"error": "PowerPoint が起動していません。"}

        try:
            pres = ppt_app.Presentations(pres_name)
            slide = pres.Slides(slide_index)
        except Exception as e:
            return {"error": str(e)}

        shapes = []
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)

            text = ""
            has_text = False
            try:
                if shape.HasTextFrame:
                    has_text = True
                    text = shape.TextFrame.TextRange.Text
            except Exception:
                pass

            has_table = False
            try:
                has_table = bool(shape.HasTable)
            except Exception:
                pass

            shapes.append(
                {
                    "index": i,
                    "name": shape.Name,
                    "type": shape.Type,
                    "has_text": has_text,
                    "text": text[:200],  # 長すぎる場合は切り詰める
                    "left_cm": round(pt_to_cm(shape.Left), 2),
                    "top_cm": round(pt_to_cm(shape.Top), 2),
                    "width_cm": round(pt_to_cm(shape.Width), 2),
                    "height_cm": round(pt_to_cm(shape.Height), 2),
                    "has_table": has_table,
                }
            )
        return shapes

    return _run_sta_as_json(_impl)


# ------------------------------------------------------------------
# Word
# ------------------------------------------------------------------


def get_wd_document_info(doc_name: str) -> str:
    """
    Word ドキュメントの構造情報を返す。

    段落数・表の数・スタイル情報を含む概要を返す。
    LLM が「どこに何があるか」を把握するための認知ツール。

    Args:
        doc_name: ドキュメント名（例: "report.docx"）

    Returns:
        JSON 文字列:
        {
            "name": "report.docx",
            "paragraphs_count": 42,
            "tables_count": 3,
            "paragraphs": [
                {"index": 1, "style": "見出し 1", "text": "はじめに"},
                ...
            ],
            "tables": [
                {"index": 1, "rows": 5, "columns": 3},
                ...
            ]
        }
    """

    def _impl():
        wd_app = manager.wd_app
        if wd_app is None:
            return {"error": "Word が起動していません。"}

        try:
            doc = wd_app.Documents(doc_name)
        except Exception:
            return {"error": f"'{doc_name}' が開いていません。"}

        # 段落情報（先頭 50 段落まで）
        paragraphs = []
        para_count = doc.Paragraphs.Count
        for i in range(1, min(para_count + 1, 51)):
            para = doc.Paragraphs(i)
            text = para.Range.Text.replace("\r", "").replace("\n", "").strip()
            paragraphs.append(
                {
                    "index": i,
                    "style": para.Style.NameLocal,
                    "text": text[:100],
                }
            )

        # 表情報
        tables = []
        for i in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(i)
            tables.append(
                {
                    "index": i,
                    "rows": tbl.Rows.Count,
                    "columns": tbl.Columns.Count,
                }
            )

        return {
            "name": doc_name,
            "paragraphs_count": para_count,
            "tables_count": doc.Tables.Count,
            "paragraphs": paragraphs,
            "note_paragraphs": "段落は先頭 50 件のみ表示しています。"
            if para_count > 50
            else "",
            "tables": tables,
        }

    return _run_sta_as_json(_impl)
