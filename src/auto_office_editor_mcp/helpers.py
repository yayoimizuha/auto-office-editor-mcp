"""
Win32com 操作で頻出するユーティリティ関数群。
executor.py が exec_globals に注入し、LLM が生成するコード内で直接呼び出せる。
"""

from __future__ import annotations

import pandas as pd
import win32com.client


# ------------------------------------------------------------------
# 色変換
# ------------------------------------------------------------------


def rgb(r: int, g: int, b: int) -> int:
    """
    RGB値を Office の OLE カラー整数（BGR順）に変換する。

    Office の VBA/COM では色を BGR 順の整数で指定する必要がある。
    LLMが直感的に rgb(255, 0, 0) と書けるようにラップする。

    例:
        shape.Fill.ForeColor.RGB = rgb(255, 0, 0)  # 赤
    """
    return r + (g * 256) + (b * 65536)


def hex_to_rgb_ole(hex_color: str) -> int:
    """
    "#FF0000" 形式の16進数カラーを OLE カラー整数に変換する。

    例:
        shape.Fill.ForeColor.RGB = hex_to_rgb_ole("#FF0000")  # 赤
    """
    hex_color = hex_color.lstrip("#")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return rgb(r, g, b)


# ------------------------------------------------------------------
# 単位変換
# ------------------------------------------------------------------


def cm_to_pt(cm: float) -> float:
    """
    センチメートルをポイント（pt）に変換する。
    Office の座標・サイズはすべてポイント単位。
    1 cm = 28.3465 pt
    """
    return cm * 28.3465


def inch_to_pt(inch: float) -> float:
    """インチをポイントに変換する。1 inch = 72 pt"""
    return inch * 72.0


def pt_to_cm(pt: float) -> float:
    """ポイントをセンチメートルに変換する。"""
    return pt / 28.3465


# ------------------------------------------------------------------
# DataFrame ↔ Excel
# ------------------------------------------------------------------


def df_to_xl(df: pd.DataFrame, start_cell) -> None:
    """
    DataFrame を Excel のセル範囲に一括書き込みする。

    Args:
        df: 書き込む DataFrame
        start_cell: 書き込み開始セル（ws.Range("B2") など）

    例:
        wb = xl_app.Workbooks("hoge.xlsx")
        ws = wb.Worksheets("Sheet1")
        df_to_xl(df, ws.Range("A1"))
    """
    # NaN は空文字に変換（COM が None/NaN を扱えないため）
    # pandas 3.0 以降は StringDtype が使われる場合があるため
    # object 型に変換してから処理する
    df_clean = df.fillna("").astype(object)

    # ヘッダー行 + データ行を結合したタプルのタプル
    # columns も明示的に str 変換して None が混入しないようにする
    headers = tuple(str(c) if c is not None else "" for c in df_clean.columns)
    data_rows = tuple(tuple(row) for row in df_clean.itertuples(index=False, name=None))
    data = (headers,) + data_rows

    rows = len(data)
    cols = len(data[0]) if data else 0
    if rows == 0 or cols == 0:
        return

    # Range(start, end) で書き込み先範囲を確定して Value に一括代入
    # Resize は gencache なし環境で正常に動作しないため使用しない
    ws = start_cell.Parent
    start_row = start_cell.Row
    start_col = start_cell.Column
    end_cell = ws.Cells(start_row + rows - 1, start_col + cols - 1)
    ws.Range(start_cell, end_cell).Value = data


def xl_to_df(range_obj) -> pd.DataFrame:
    """
    Excel のセル範囲（Range オブジェクト）を DataFrame に変換する。

    Args:
        range_obj: ws.Range("A1:D10") などの Range オブジェクト

    Returns:
        1行目をヘッダーとした DataFrame

    例:
        wb = xl_app.Workbooks("hoge.xlsx")
        ws = wb.Worksheets("Sheet1")
        df = xl_to_df(ws.Range("A1:D10"))
    """
    raw = range_obj.Value

    if raw is None:
        return pd.DataFrame()

    # 単一セルの場合はタプルではなくスカラーが返る
    if not isinstance(raw, tuple):
        raw = ((raw,),)

    # 1行だけの場合はタプルのタプルではなくタプルが返る
    if raw and not isinstance(raw[0], tuple):
        raw = (raw,)

    rows = [list(row) for row in raw]
    if len(rows) < 2:
        return pd.DataFrame(rows)

    headers = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=headers)


# ------------------------------------------------------------------
# DataFrame ↔ PowerPoint 表
# ------------------------------------------------------------------


def _prepare_table_data(df: pd.DataFrame, include_header: bool) -> list[list]:
    """
    DataFrame をセルデータのリスト（行×列）に変換する共通前処理。

    Args:
        df: 変換対象の DataFrame
        include_header: True の場合、列名を先頭行として追加する

    Returns:
        セルデータの2次元リスト

    Raises:
        ValueError: DataFrame が空の場合
    """
    if include_header:
        data = [df.columns.tolist()] + df.fillna("").values.tolist()
    else:
        data = df.fillna("").values.tolist()

    num_rows = len(data)
    num_cols = len(data[0]) if data else 0
    if num_rows == 0 or num_cols == 0:
        raise ValueError("DataFrame が空です。")

    return data


def df_to_ppt_table(
    df: pd.DataFrame,
    slide,
    left_cm: float = 2.0,
    top_cm: float = 5.0,
    width_cm: float = 20.0,
    height_cm: float = 8.0,
    include_header: bool = True,
) -> object:
    """
    DataFrame を PowerPoint スライドに表として追加する。

    Args:
        df: 書き込む DataFrame
        slide: 対象スライドオブジェクト（ppt_app.ActivePresentation.Slides(1) など）
        left_cm, top_cm: 表の左上位置（cm）
        width_cm, height_cm: 表のサイズ（cm）
        include_header: True の場合、DataFrame の列名を1行目として追加する

    Returns:
        追加された表の Shape オブジェクト

    例:
        slide = ppt_app.Presentations("slide.pptx").Slides(1)
        df_to_ppt_table(df, slide)
    """
    data = _prepare_table_data(df, include_header)
    num_rows = len(data)
    num_cols = len(data[0])

    shape = slide.Shapes.AddTable(
        NumRows=num_rows,
        NumColumns=num_cols,
        Left=cm_to_pt(left_cm),
        Top=cm_to_pt(top_cm),
        Width=cm_to_pt(width_cm),
        Height=cm_to_pt(height_cm),
    )
    table = shape.Table

    for r, row in enumerate(data):
        for c, val in enumerate(row):
            table.Cell(r + 1, c + 1).Shape.TextFrame.TextRange.Text = str(val)

    return shape


def ppt_table_to_df(shape) -> pd.DataFrame:
    """
    PowerPoint の表 Shape を DataFrame に変換する。

    Args:
        shape: HasTable が True の Shape オブジェクト

    Returns:
        1行目をヘッダーとした DataFrame

    例:
        for shape in ppt_app.Presentations("slide.pptx").Slides(1).Shapes:
            if shape.HasTable:
                df = ppt_table_to_df(shape)
                break
    """
    if not shape.HasTable:
        raise ValueError("指定した Shape は表ではありません。")

    table = shape.Table
    data = []
    for r in range(1, table.Rows.Count + 1):
        row = []
        for c in range(1, table.Columns.Count + 1):
            text = table.Cell(r, c).Shape.TextFrame.TextRange.Text
            row.append(text.strip())
        data.append(row)

    if len(data) < 2:
        return pd.DataFrame(data)

    return pd.DataFrame(data[1:], columns=data[0])


# ------------------------------------------------------------------
# DataFrame ↔ Word 表
# ------------------------------------------------------------------


def df_to_wd_table(
    df: pd.DataFrame,
    doc,
    include_header: bool = True,
    range_obj=None,
) -> object:
    """
    DataFrame を Word ドキュメントに表として追加する。

    Args:
        df: 書き込む DataFrame
        doc: Word ドキュメントオブジェクト（wd_app.Documents("report.docx") など）
        include_header: True の場合、DataFrame の列名を1行目として追加する
        range_obj: 挿入位置の Range。None の場合はドキュメント末尾に追加する

    Returns:
        追加された Word Table オブジェクト

    例:
        doc = wd_app.Documents("report.docx")
        df_to_wd_table(df, doc)
    """
    data = _prepare_table_data(df, include_header)
    num_rows = len(data)
    num_cols = len(data[0])

    insert_range = (
        range_obj
        if range_obj is not None
        else doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    )

    table = doc.Tables.Add(Range=insert_range, NumRows=num_rows, NumColumns=num_cols)

    for r, row in enumerate(data):
        for c, val in enumerate(row):
            table.Cell(r + 1, c + 1).Range.Text = str(val)

    return table


def wd_table_to_df(table) -> pd.DataFrame:
    """
    Word の Table オブジェクトを DataFrame に変換する。

    Word のセルテキストには末尾に制御文字（\\r\\x07 など）が付くため除去する。

    Args:
        table: doc.Tables(1) などの Word Table オブジェクト

    Returns:
        1行目をヘッダーとした DataFrame

    例:
        doc = wd_app.Documents("report.docx")
        df = wd_table_to_df(doc.Tables(1))
    """
    data = []
    for row in table.Rows:
        row_data = []
        for cell in row.Cells:
            # Word のセルテキストは末尾に \r\x07 または \x07 が付く
            text = cell.Range.Text.replace("\r\x07", "").replace("\x07", "").strip()
            row_data.append(text)
        data.append(row_data)

    if len(data) < 2:
        return pd.DataFrame(data)

    return pd.DataFrame(data[1:], columns=data[0])
