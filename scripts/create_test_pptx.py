"""
tests/fixtures/test.pptx を win32com で生成するスクリプト。

構成:
  スライド1（レイアウト1: タイトルスライド）
    - タイトルプレースホルダー: "四半期売上報告"
    - サブタイトルプレースホルダー: "2024年度 第4四半期"

  スライド2（レイアウト2: タイトルとコンテンツ）
    - タイトルプレースホルダー: "部門別売上"
    - 表（3行×3列）: 部門/売上/前年比 + データ2行

  スライド3（レイアウト6: 空白）
    - テキストボックス: "グラフスライド"
    - 長方形AutoShape（名前: "SampleRect"）: left=3cm, top=5cm, width=8cm, height=4cm

テストから期待する値:
  slides_count: 3
  slide1.shapes_count: 2
  slide1 title text: "四半期売上報告"
  slide2.shapes_count: 2  （タイトル + 表）
  slide2 に has_table=True の図形が存在する
  slide3.shapes_count: 2  （テキストボックス + 長方形）
  slide3 に has_text=True の図形が存在する
  SampleRect の left_cm ≈ 3.0
"""

from __future__ import annotations

from pathlib import Path

import win32com.client

OUTPUT_PATH = Path(__file__).parent.parent / "tests" / "fixtures" / "test.pptx"

# ポイント変換（1cm = 28.3465pt）
CM = 28.3465


def cm(v: float) -> float:
    return v * CM


def main() -> None:
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = True  # PowerPoint は Visible=True が必要な場合がある

    try:
        pres = ppt.Presentations.Add(WithWindow=True)

        # スライドのデフォルトサイズ確認
        slide_width = pres.PageSetup.SlideWidth  # 通常 720pt (25.4cm)
        slide_height = pres.PageSetup.SlideHeight  # 通常 540pt (19.05cm)
        print(f"スライドサイズ: {slide_width:.1f}pt × {slide_height:.1f}pt")

        # ----------------------------------------------------------------
        # スライド1: タイトルスライド（レイアウト1）
        # ----------------------------------------------------------------
        slide1 = pres.Slides.Add(1, 1)  # ppLayoutTitle = 1

        # タイトル（プレースホルダー1）
        slide1.Shapes.Title.TextFrame.TextRange.Text = "四半期売上報告"

        # サブタイトル（プレースホルダー2）
        slide1.Shapes(2).TextFrame.TextRange.Text = "2024年度 第4四半期"

        print(f"スライド1 shapes_count: {slide1.Shapes.Count}")

        # ----------------------------------------------------------------
        # スライド2: タイトル + 表（レイアウト2: タイトルとコンテンツ）
        # ----------------------------------------------------------------
        slide2 = pres.Slides.Add(2, 2)  # ppLayoutText = 2

        # タイトル
        slide2.Shapes.Title.TextFrame.TextRange.Text = "部門別売上"

        # コンテンツプレースホルダーを削除して表を追加
        # （プレースホルダー2を削除してから表を直接追加）
        for i in range(slide2.Shapes.Count, 0, -1):
            sh = slide2.Shapes(i)
            try:
                if sh.PlaceholderFormat.Type != 1:  # タイトル以外を削除
                    sh.Delete()
            except Exception:
                pass

        # 表（3行×3列）を追加
        table_shape = slide2.Shapes.AddTable(
            NumRows=3,
            NumColumns=3,
            Left=cm(2),
            Top=cm(4),
            Width=cm(20),
            Height=cm(6),
        )
        table = table_shape.Table

        table_data = [
            ["部門", "売上", "前年比"],
            ["東京", "1500000", "120%"],
            ["大阪", "980000", "105%"],
        ]
        for r, row in enumerate(table_data):
            for c_idx, val in enumerate(row):
                table.Cell(r + 1, c_idx + 1).Shape.TextFrame.TextRange.Text = val

        print(f"スライド2 shapes_count: {slide2.Shapes.Count}")

        # ----------------------------------------------------------------
        # スライド3: 空白スライド（レイアウト12: 空白）
        # ----------------------------------------------------------------
        slide3 = pres.Slides.Add(3, 12)  # ppLayoutBlank = 12

        # テキストボックス
        tbx = slide3.Shapes.AddTextbox(
            Orientation=1,  # msoTextOrientationHorizontal
            Left=cm(2),
            Top=cm(2),
            Width=cm(15),
            Height=cm(2),
        )
        tbx.TextFrame.TextRange.Text = "グラフスライド"

        # 長方形AutoShape（名前: "SampleRect"）
        # msoShapeRectangle = 1
        rect = slide3.Shapes.AddShape(
            Type=1,
            Left=cm(3),
            Top=cm(5),
            Width=cm(8),
            Height=cm(4),
        )
        rect.Name = "SampleRect"
        rect.TextFrame.TextRange.Text = ""

        print(f"スライド3 shapes_count: {slide3.Shapes.Count}")

        # ----------------------------------------------------------------
        # 保存
        # ----------------------------------------------------------------
        OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
        pres.SaveAs(str(OUTPUT_PATH.resolve()))
        print(f"保存完了: {OUTPUT_PATH.resolve()}")

        # 確認出力
        print(f"  slides_count: {pres.Slides.Count}")
        for si in range(1, pres.Slides.Count + 1):
            sl = pres.Slides(si)
            print(f"  slide[{si}] name={sl.Name!r} shapes_count={sl.Shapes.Count}")
            for shi in range(1, sl.Shapes.Count + 1):
                sh = sl.Shapes(shi)
                has_text = False
                text = ""
                try:
                    if sh.HasTextFrame:
                        has_text = True
                        text = sh.TextFrame.TextRange.Text[:40]
                except Exception:
                    pass
                has_table = False
                try:
                    has_table = bool(sh.HasTable)
                except Exception:
                    pass
                print(
                    f"    shape[{shi}] name={sh.Name!r:20s} "
                    f"type={sh.Type} has_text={has_text} has_table={has_table} "
                    f"left={sh.Left / CM:.2f}cm top={sh.Top / CM:.2f}cm "
                    f"text={text!r}"
                )

    finally:
        pres.Close()
        ppt.Quit()


if __name__ == "__main__":
    main()
