"""
FastMCP サーバー定義。

公開するツール一覧:
    execute_office_script   - win32com コードを exec() で実行（コアツール）
    open_office_file        - ファイルを開く
    close_office_file       - ファイルを閉じる
    list_open_files         - 現在開いているファイルの一覧
    get_excel_sheets        - Excel シート一覧
    get_excel_range         - Excel セル範囲の値と情報（先頭20行）
    get_powerpoint_slides   - PowerPoint スライド一覧
    get_powerpoint_shapes   - PowerPoint スライドの図形一覧と座標
    get_word_document       - Word ドキュメントの段落・表の構造情報
    preview_word_pages      - Word ページを PNG 画像として返す
    preview_ppt_slides      - PowerPoint スライドを PNG 画像として返す
"""

from __future__ import annotations

import argparse
import logging
from contextlib import asynccontextmanager
from typing import AsyncIterator, Sequence

from fastmcp import FastMCP
from fastmcp.utilities.types import Image

from .app_manager import manager
from .executor import execute_script
from .tools.file_ops import close_file, get_open_files, open_file
from .tools.info import (
    get_ppt_shapes_info,
    get_ppt_slides_info,
    get_wd_document_info,
    get_xl_range_info,
    get_xl_sheets_info,
)
from .tools.preview import get_ppt_slide_image, get_word_page_image

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
logger = logging.getLogger(__name__)


# ------------------------------------------------------------------
# ライフスパン：サーバー起動時に Office アプリを起動し、
#              サーバー停止時に全アプリを終了する
# ------------------------------------------------------------------


@asynccontextmanager
async def lifespan(server: FastMCP) -> AsyncIterator[None]:
    logger.info("Starting Office apps...")
    manager.start()
    try:
        yield
    finally:
        logger.info("Shutting down Office apps...")
        manager.shutdown()


# ------------------------------------------------------------------
# FastMCP インスタンス
# ------------------------------------------------------------------

mcp = FastMCP(
    name="auto-office-editor",
    instructions=(
        "Win32com を使って Excel / Word / PowerPoint をバックグラウンドで操作するツール群です。\n"
        "MCPサーバー起動時に3アプリが自動的に起動し、停止時に開いているファイルを全て保存して終了します。\n"
        "\n"
        "## 基本的な使い方\n"
        "1. open_office_file でファイルを開く\n"
        "2. get_excel_sheets / get_powerpoint_shapes / get_word_document などで現在の状態を確認する\n"
        "3. execute_office_script で操作コードを実行する\n"
        "4. 必要に応じて preview_word_pages / preview_ppt_slides で結果を目視確認する\n"
        "5. close_office_file で保存して閉じる（省略時はサーバー停止時に自動保存）\n"
        "\n"
        "注入済み変数・ヘルパー関数・注意事項の詳細は execute_office_script の説明を参照してください。\n"
    ),
    lifespan=lifespan,
)


# ------------------------------------------------------------------
# ツール登録
# ------------------------------------------------------------------


@mcp.tool()
def execute_office_script(code: str) -> str:
    """
    Python コード（win32com を使った Office 操作）を実行する。

    コードは以下の変数・モジュール・関数が注入済みの環境で実行される。
    コード内で print() した内容は戻り値に含まれる。
    エラー発生時は Traceback を返すので、内容を読んでコードを修正して再実行すること。

    ## 注入済み変数

    ### Office アプリオブジェクト
    - xl_app  : Excel.Application
    - wd_app  : Word.Application
    - ppt_app : PowerPoint.Application
    - const   : win32com.client.constants（xlUp, msoTrue などの VBA 定数）

    ### 標準モジュール
    - pd, np, os, re, sys, json, datetime, Path, win32com, pythoncom

    ### ヘルパー関数
    - rgb(r, g, b)             : RGB → OLE カラー整数 ※色指定は必ずこれを使うこと
    - hex_to_rgb_ole('#FF0000') : 16進カラー → OLE カラー整数
    - cm_to_pt(cm)             : cm → ポイント（Office の座標単位）
    - inch_to_pt(inch)         : インチ → ポイント
    - pt_to_cm(pt)             : ポイント → cm
    - df_to_xl(df, start_cell) : DataFrame を Excel のセル範囲に一括書き込み
    - xl_to_df(range_obj)      : Excel のセル範囲を DataFrame に変換（1行目をヘッダーとして扱う）
    - df_to_ppt_table(df, slide, left_cm=2, top_cm=5, width_cm=20, height_cm=8)
                               : DataFrame を PPT 表として追加
    - ppt_table_to_df(shape)   : PPT の表 Shape を DataFrame に変換
    - df_to_wd_table(df, doc, include_header=True, range_obj=None)
                               : DataFrame を Word 表として追加
    - wd_table_to_df(table)    : Word の表を DataFrame に変換

    ## ファイルへのアクセス方法
    ファイルは open_office_file で開いた後、ファイル名で明示的に指定する。
    ActiveWorkbook / ActiveDocument などは使用しないこと。

        wb    = xl_app.Workbooks('hoge.xlsx')
        ws    = wb.Worksheets('Sheet1')
        doc   = wd_app.Documents('report.docx')
        pres  = ppt_app.Presentations('slide.pptx')
        slide = pres.Slides(1)  # 1始まり

    ## コード例

        # Excel: セルに値を書き込んで書式を設定する
        wb = xl_app.Workbooks('hoge.xlsx')
        ws = wb.Worksheets('Sheet1')
        ws.Range('A1').Value = '完了'
        ws.Range('A1').Font.Bold = True
        ws.Range('A1').Interior.Color = rgb(255, 255, 0)  # 黄色

        # Excel: セル範囲を DataFrame に読み込む
        df = xl_to_df(ws.Range('A1:D10'))

        # PowerPoint: 図形の位置を変更する
        slide = ppt_app.Presentations('slide.pptx').Slides(1)
        slide.Shapes('TextBox 1').Left = cm_to_pt(3.0)
        slide.Shapes('TextBox 1').Fill.ForeColor.RGB = rgb(255, 0, 0)

        # Word: 段落スタイルを変更する
        doc = wd_app.Documents('report.docx')
        doc.Paragraphs(1).Style = doc.Styles('見出し 1')

    ## win32com 共通の注意点（VBA との違い）

    ### コレクションのインデックスは 1始まり
    全コレクション（Worksheets / Slides / Shapes / Paragraphs 等）は 1始まり。
    名前でのアクセスも VBA と同様に使える:
        slide.Shapes('SampleRect')   # 名前で図形を取得

    ### Boolean / None
    - True/False はそのまま使う: range.Bold = True
    - 空セルは None で返る（VBA の Empty とは異なる）:
        if ws.Range('A1').Value is not None:

    ### 存在しない名前・インデックスは com_error
    事前に get_powerpoint_shapes / get_word_document 等で名前を確認してからアクセスすること。

    ## Word 固有の注意点

    ### 表セルのテキストに末尾制御文字が付く
    table.Cell(1,1).Range.Text の末尾には '\\r\\x07' が付く。必ずトリムすること:
        text = table.Cell(1,1).Range.Text.replace('\\r\\x07','').replace('\\x07','').strip()

    ### doc.Paragraphs.Count は表セル内段落も含む
    表セルも段落としてカウントされるため本文段落数とは一致しない。

    ### ドキュメント末尾への挿入
    doc.Content.End をそのまま使うと範囲外になる。必ず -1 する:
        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        rng.InsertParagraphAfter()

    ### 表セルへのテキスト設定（Word）
        table.Cell(row, col).Range.Text = '値'   # Range 経由

    ## PowerPoint 固有の注意点

    ### ExportAsFixedFormat は使用禁止
    win32com の gencache 環境で TypeError になる。PDF 出力は SaveAs を使うこと:
        pres.SaveAs('C:/path/to/out.pdf', 32)   # 32 = ppSaveAsPDF

    ### 表セルへのテキスト設定（PowerPoint）
    Word と異なり Cell に .Range はない。.Shape.TextFrame.TextRange を経由すること:
        table.Cell(row, col).Shape.TextFrame.TextRange.Text = '値'

    ### よく使うスライドレイアウト番号
        1  : タイトルスライド, 2  : タイトルとコンテンツ, 12 : 空白

    ### よく使う図形タイプ番号（AddShape の Type 引数）
        1 : 長方形, 9 : 楕円  ※テキストボックスは AddShape ではなく AddTextbox を使う

    Args:
        code: 実行する Python コード文字列

    Returns:
        成功時: "実行成功"（print() の出力があれば "--- 出力 ---" セクションに含まれる）
        失敗時: Traceback 全文の文字列
    """
    return execute_script(code)


@mcp.tool()
def open_office_file(filepath: str) -> str:
    """
    Office ファイルを開く。

    拡張子から自動的にアプリ（Excel / Word / PowerPoint）を判定してファイルを開く。
    対応拡張子: .xlsx .xlsm .xlsb .xls / .docx .doc .dotx / .pptx .ppt .potx

    同名ファイルの扱い:
    - 同一パスのファイルが既に開いている → already_open=true で返す（再オープンしない）
    - 同名だが別パスのファイルが開いている → エラーを返す
      （Office の仕様上、同名ファイルを別パスから同時に開くことはできない）

    注意: インターネットからダウンロードした Word ファイルは「保護されたビュー」で
    開かれる場合があります。その場合は自動的に編集モードに切り替えます。
    それでも応答がない場合は、Word の Trust Center で「保護されたビュー」を
    手動で無効化してから再試行してください。

    Args:
        filepath: 開くファイルの絶対パス（例: "C:/Users/me/report.xlsx"）

    Returns:
        JSON: {
            "app": "excel",
            "name": "report.xlsx",
            "path": "C:/Users/me/report.xlsx",
            "already_open": false
        }
    """
    return open_file(filepath)


@mcp.tool()
def close_office_file(filepath: str, save: bool = True) -> str:
    """
    開いている Office ファイルを閉じる。

    Args:
        filepath: 閉じるファイルのパスまたはファイル名（例: "report.xlsx"）
        save: True の場合は上書き保存してから閉じる（デフォルト: True）

    Returns:
        JSON: {"closed": "report.xlsx", "saved": true}
    """
    return close_file(filepath, save=save)


@mcp.tool()
def list_open_files() -> str:
    """
    現在開いている Office ファイルの一覧を返す。

    Returns:
        JSON: {"excel": [...], "word": [...], "powerpoint": [...]}
    """
    return get_open_files()


@mcp.tool()
def get_excel_sheets(wb_name: str) -> str:
    """
    Excel ワークブックのシート一覧を返す。

    Args:
        wb_name: ワークブック名（例: "hoge.xlsx"）

    Returns:
        JSON: [{"index": 1, "name": "Sheet1", "visible": true}, ...]
    """
    return get_xl_sheets_info(wb_name)


@mcp.tool()
def get_excel_range(wb_name: str, sheet_name: str, range_address: str) -> str:
    """
    Excel の指定セル範囲の値と情報を返す。値は先頭 20 行まで表示。

    20 行を超える範囲を扱う場合は execute_office_script で
    xl_to_df(ws.Range("A1:D100")) を直接呼び出してください。

    Args:
        wb_name: ワークブック名（例: "hoge.xlsx"）
        sheet_name: シート名（例: "Sheet1"）
        range_address: セル範囲（例: "A1:D10"）

    Returns:
        JSON: {"address": "$A$1:$D$10", "rows": 10, "columns": 4, "values": [[...]]}
    """
    return get_xl_range_info(wb_name, sheet_name, range_address)


@mcp.tool()
def get_powerpoint_slides(pres_name: str) -> str:
    """
    PowerPoint プレゼンテーションのスライド一覧を返す。

    Args:
        pres_name: プレゼンテーション名（例: "slide.pptx"）

    Returns:
        JSON: [{"index": 1, "name": "スライド1", "shapes_count": 3}, ...]
    """
    return get_ppt_slides_info(pres_name)


@mcp.tool()
def get_powerpoint_shapes(pres_name: str, slide_index: int) -> str:
    """
    PowerPoint スライド上の全図形の情報を返す。

    図形の名前・座標・サイズ・テキスト内容を含む。
    execute_office_script で図形を操作する前に、対象図形の名前を確認するために使う。

    Args:
        pres_name: プレゼンテーション名（例: "slide.pptx"）
        slide_index: スライド番号（1 始まり）

    Returns:
        JSON: [{"name": "TextBox 1", "left_cm": 2.5, "top_cm": 1.2, ...}, ...]
    """
    return get_ppt_shapes_info(pres_name, slide_index)


@mcp.tool()
def get_word_document(doc_name: str) -> str:
    """
    Word ドキュメントの段落・表の構造情報を返す。段落は先頭 50 件まで。

    Args:
        doc_name: ドキュメント名（例: "report.docx"）

    Returns:
        JSON: {"paragraphs_count": 42, "tables_count": 3, "paragraphs": [...], "tables": [...]}
    """
    return get_wd_document_info(doc_name)


def _to_images(results) -> list[Image]:
    """(page_index, png_bytes) のリストを Image オブジェクトのリストに変換する。"""
    return [Image(data=png_bytes, format="png") for _, png_bytes in results]


@mcp.tool()
def preview_word_pages(
    doc_name: str,
    pages: int | list[int] | None = None,
    dpi: int = 150,
) -> list[Image]:
    """
    Word ドキュメントの指定ページを画像として返す。

    PDF 経由でレンダリングするため、実際の印刷イメージに近い結果が得られる。
    体裁確認・レイアウト確認・編集後の目視チェックに使う。

    Args:
        doc_name: ドキュメント名（例: "report.docx"）
        pages:    ページ番号（1始まり）。
                  未指定 → 全ページ（20ページ超はエラー）
                  int    → 単一ページ（例: 1）
                  list   → 複数ページ（例: [1, 3, 5]）
        dpi:      出力解像度（デフォルト: 150）。高品質が必要なら 300 を指定。

    Returns:
        PNG 画像のリスト（各ページ1枚）
    """
    results = get_word_page_image(doc_name, pages=pages, dpi=dpi)
    return _to_images(results)


@mcp.tool()
def preview_ppt_slides(
    pres_name: str,
    pages: int | list[int] | None = None,
    dpi: int = 150,
) -> list[Image]:
    """
    PowerPoint の指定スライドを画像として返す。

    PDF 経由でレンダリングするため、フォント・図形・グラデーションを含む
    実際の表示に近い結果が得られる。スライドのレイアウト確認・編集後チェックに使う。

    Args:
        pres_name: プレゼンテーション名（例: "slide.pptx"）
        pages:     スライド番号（1始まり）。
                   未指定 → 全スライド（20枚超はエラー）
                   int    → 単一スライド（例: 2）
                   list   → 複数スライド（例: [1, 2, 3]）
        dpi:       出力解像度（デフォルト: 150）。高品質が必要なら 300 を指定。

    Returns:
        PNG 画像のリスト（各スライド1枚）
    """
    results = get_ppt_slide_image(pres_name, pages=pages, dpi=dpi)
    return _to_images(results)


# ------------------------------------------------------------------
# エントリーポイント
# ------------------------------------------------------------------


def _parse_args(argv: Sequence[str] | None = None):
    parser = argparse.ArgumentParser(description="Auto Office Editor MCP server")
    parser.add_argument(
        "--visible",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Show Office application windows (use --no-visible to hide).",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> None:
    args = _parse_args(argv)
    manager.set_visibility(bool(args.visible))
    mcp.run()


if __name__ == "__main__":
    main()
