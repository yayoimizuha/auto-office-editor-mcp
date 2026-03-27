"""
tests/fixtures/test.docx を win32com で生成するスクリプト。

構成:
  - 見出し1: "売上レポート"
  - 標準段落: "本レポートは2024年度の売上実績をまとめたものです。"
  - 見出し2: "概要"
  - 標準段落: "全体の売上は前年比120%となりました。"
  - 見出し2: "部門別売上"
  - 表（3行×3列）: 部門/売上/前年比 + データ2行
  - 末尾空段落（df_to_wd_table の挿入位置）

注意: Word の doc.Paragraphs は表セル内の段落もカウントする。
  本文段落 5 + 表セル 9(3行×3列) + 末尾空段落 1 + その他 = 実測値で確認すること。

テストから期待する値（実行後の出力で確認）:
  tables_count: 1
  tables[0]: rows=3, columns=3
  paragraphs[0].style: "見出し 1"
  paragraphs[0].text:  "売上レポート"
"""

from __future__ import annotations

import sys
from pathlib import Path

import win32com.client

OUTPUT_PATH = Path(__file__).parent.parent / "tests" / "fixtures" / "test.docx"


def main() -> None:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        doc = word.Documents.Add()

        # ----------------------------------------------------------------
        # 段落1: 見出し1
        # ----------------------------------------------------------------
        para = doc.Paragraphs(1)
        para.Style = doc.Styles("見出し 1")
        para.Range.Text = "売上レポート"

        # ----------------------------------------------------------------
        # 段落2: 標準テキスト
        # ----------------------------------------------------------------
        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        rng.InsertParagraphAfter()
        para2 = doc.Paragraphs(doc.Paragraphs.Count)
        para2.Style = doc.Styles("標準")
        para2.Range.Text = "本レポートは2024年度の売上実績をまとめたものです。"

        # ----------------------------------------------------------------
        # 段落3: 見出し2
        # ----------------------------------------------------------------
        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        rng.InsertParagraphAfter()
        para3 = doc.Paragraphs(doc.Paragraphs.Count)
        para3.Style = doc.Styles("見出し 2")
        para3.Range.Text = "概要"

        # ----------------------------------------------------------------
        # 段落4: 標準テキスト
        # ----------------------------------------------------------------
        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        rng.InsertParagraphAfter()
        para4 = doc.Paragraphs(doc.Paragraphs.Count)
        para4.Style = doc.Styles("標準")
        para4.Range.Text = "全体の売上は前年比120%となりました。"

        # ----------------------------------------------------------------
        # 段落5: 見出し2
        # ----------------------------------------------------------------
        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        rng.InsertParagraphAfter()
        para5 = doc.Paragraphs(doc.Paragraphs.Count)
        para5.Style = doc.Styles("見出し 2")
        para5.Range.Text = "部門別売上"

        # ----------------------------------------------------------------
        # 表（3行×3列）をドキュメント末尾に追加
        # ----------------------------------------------------------------
        insert_rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        table = doc.Tables.Add(Range=insert_rng, NumRows=3, NumColumns=3)

        table_data = [
            ["部門", "売上", "前年比"],
            ["東京", "1500000", "120%"],
            ["大阪", "980000", "105%"],
        ]
        for r, row in enumerate(table_data):
            for c, val in enumerate(row):
                table.Cell(r + 1, c + 1).Range.Text = val

        # ----------------------------------------------------------------
        # 末尾に空段落（df_to_wd_table の挿入位置確認用）
        # ----------------------------------------------------------------
        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        rng.InsertParagraphAfter()

        # ----------------------------------------------------------------
        # 保存
        # ----------------------------------------------------------------
        OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
        doc.SaveAs2(str(OUTPUT_PATH.resolve()))
        print(f"保存完了: {OUTPUT_PATH.resolve()}")

        # 段落数・表数を確認
        print(f"  paragraphs_count : {doc.Paragraphs.Count}")
        print(f"  tables_count     : {doc.Tables.Count}")
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            text = p.Range.Text.replace("\r", "").strip()
            style = p.Style.NameLocal
            print(f"  para[{i}] style={style!r:20s} text={text!r}")

    finally:
        doc.Close(SaveChanges=False)
        word.Quit()


if __name__ == "__main__":
    main()
