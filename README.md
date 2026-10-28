# auto-office-editor-mcp

Windows 上の Excel / Word / PowerPoint を win32com でバックグラウンド操作する MCP サーバーです。
LLM が自然言語の指示から Office ファイルを直接編集・プレビューできます。

## 特徴

- **Excel / Word / PowerPoint を統一インターフェースで操作**
- **`execute_office_script` でほぼ全ての VBA 操作を Python から実行可能**
- **PDF 経由の PNG プレビュー**（編集結果をその場で目視確認）
- **DataFrame ↔ Office 表の双方向変換**ヘルパー内蔵

## 動作要件

| 要件 | 詳細 |
|---|---|
| OS | Windows 10 / 11 |
| Python | 3.11 以上 |
| Microsoft Office | Excel / Word / PowerPoint がインストール済みであること |
| パッケージ管理 | [uv](https://docs.astral.sh/uv/) を推奨 |

## セットアップ

```bash
git clone https://github.com/yayoimizuha/auto-office-editor-mcp.git
cd auto-office-editor-mcp
uv sync
```

`uv sync` は開発・テスト用です。
MCP クライアントから使うときは、通常は `uvx` でこのパッケージを起動します。

## MCP クライアントへの登録

`claude_desktop_config.json` などに以下を追加します。

### 1. ローカル checkout をそのまま使う場合

```json
{
  "mcpServers": {
    "auto-office-editor": {
      "command": "uvx",
      "args": [
        "--from",
        "C:/path/to/auto-office-editor-mcp",
        "auto-office-editor-mcp"
      ]
    }
  }
}
```

### 2. Git リポジトリ URL から直接起動する場合

```json
{
  "mcpServers": {
    "auto-office-editor": {
      "command": "uvx",
      "args": [
        "--from",
        "git+https://github.com/your-org/auto-office-editor-mcp.git",
        "auto-office-editor-mcp"
      ]
    }
  }
}
```

どちらの方法でも、`uvx` は指定した配布元から一時環境を作成し、
`pyproject.toml` の `project.scripts` に定義された
`auto-office-editor-mcp` コマンドを実行します。

手元で単体起動して確認したい場合は、次でも起動できます。

```bash
uvx --from . auto-office-editor-mcp
```

サーバー起動時は COM 用の STA ワーカーだけを起動します。
Excel / Word / PowerPoint は `open_office_file` で対象ファイルを最初に開くときに起動します。
サーバー停止時には起動済みアプリで開いているファイルを全て保存して終了します。

### ウィンドウの表示/非表示を切り替える

デフォルトでは Office アプリのウィンドウを表示します。バックグラウンドで動かしたい場合は `--no-visible` を付けて起動してください。

```bash
uvx --from . auto-office-editor-mcp --no-visible
```

`--visible / --no-visible` で表示状態を切り替えられます。

## 公開ツール一覧

| ツール | 説明 |
|---|---|
| `execute_office_script(code)` | Python コード（win32com）を実行するコアツール |
| `open_office_file(filepath)` | ファイルを開く（拡張子からアプリを自動判定） |
| `close_office_file(filepath, save)` | ファイルを閉じる（デフォルトで保存） |
| `list_open_files()` | 現在開いているファイルの一覧を返す |
| `get_excel_sheets(wb_name)` | Excel シート一覧 |
| `get_excel_range(wb_name, sheet_name, range_address)` | Excel セル範囲の値（先頭 20 行） |
| `get_powerpoint_slides(pres_name)` | PowerPoint スライド一覧 |
| `get_powerpoint_shapes(pres_name, slide_index)` | スライド上の図形一覧・座標 |
| `get_word_document(doc_name)` | Word ドキュメントの段落・表の構造情報 |
| `preview_word_pages(doc_name, pages, dpi)` | Word ページを PNG 画像として返す |
| `preview_ppt_slides(pres_name, pages, dpi)` | PowerPoint スライドを PNG 画像として返す |

## 基本的な使い方

```
1. open_office_file でファイルを開く
2. get_excel_sheets / get_powerpoint_shapes / get_word_document で現在の状態を確認する
3. execute_office_script で操作コードを実行する
4. preview_word_pages / preview_ppt_slides で結果を目視確認する（任意）
5. close_office_file で保存して閉じる
```

## execute_office_script の使い方

`execute_office_script` に渡すコードには、以下の変数・関数が注入済みです。

### 注入済み変数

```python
# Office アプリオブジェクト
xl_app   # Excel.Application（未起動なら None）
wd_app   # Word.Application（未起動なら None）
ppt_app  # PowerPoint.Application（未起動なら None）
const    # win32com.client.constants（xlUp, msoTrue 等の VBA 定数）

# 標準モジュール
pd, np, os, re, sys, json, datetime, Path, win32com, pythoncom
```

### ヘルパー関数

```python
rgb(r, g, b)                # RGB → OLE カラー整数（色指定は必ずこれを使う）
hex_to_rgb_ole('#FF0000')   # 16進カラー → OLE カラー整数
cm_to_pt(cm)                # cm → ポイント
inch_to_pt(inch)            # インチ → ポイント
pt_to_cm(pt)                # ポイント → cm

df_to_xl(df, start_cell)    # DataFrame を Excel のセル範囲に一括書き込み
xl_to_df(range_obj)         # Excel セル範囲 → DataFrame（1行目をヘッダーとして扱う）

df_to_ppt_table(df, slide, left_cm, top_cm, width_cm, height_cm)  # DataFrame → PPT 表
ppt_table_to_df(shape)      # PPT 表 Shape → DataFrame

df_to_wd_table(df, doc, include_header=True, range_obj=None)  # DataFrame → Word 表
wd_table_to_df(table)       # Word 表 → DataFrame
```

### コード例

`execute_office_script` は未起動アプリを自動起動しません。必要なファイルは先に `open_office_file` で開いてください。

```python
# Excel: セルに値を書き込んで書式を設定する
wb = xl_app.Workbooks('hoge.xlsx')
ws = wb.Worksheets('Sheet1')
ws.Range('A1').Value = '完了'
ws.Range('A1').Font.Bold = True
ws.Range('A1').Interior.Color = rgb(255, 255, 0)  # 黄色

# Excel: セル範囲を DataFrame に読み込む
df = xl_to_df(ws.Range('A1:D10'))

# PowerPoint: 図形のテキストと位置を変更する
slide = ppt_app.Presentations('slide.pptx').Slides(1)
slide.Shapes('TextBox 1').TextFrame.TextRange.Text = '新しいテキスト'
slide.Shapes('TextBox 1').Left = cm_to_pt(3.0)
slide.Shapes('TextBox 1').Fill.ForeColor.RGB = rgb(255, 0, 0)

# Word: 段落スタイルを変更して表を追加する
doc = wd_app.Documents('report.docx')
doc.Paragraphs(1).Style = doc.Styles('見出し 1')
df_to_wd_table(df, doc)
```

## VBA と win32com の主な違い（注意点）

LLM が誤りやすい差異を以下にまとめます。

### 共通

| 項目 | 説明 |
|---|---|
| **インデックスは 1始まり** | `Worksheets(1)`, `Slides(1)`, `Shapes(1)` 等、全コレクションは VBA と同様に 1始まり |
| **名前アクセス** | `slide.Shapes('SampleRect')` のように名前でもアクセス可能（VBA と同じ） |
| **Boolean** | VBA の `True/False` はそのまま Python の `True/False` を使う |
| **空セルは `None`** | VBA の `Empty` に相当するセル値は Python では `None` になる |
| **存在しない名前・インデックス** | `com_error` が発生する。事前に `get_powerpoint_shapes` 等で名前を確認すること |

### Word 固有

| 項目 | 説明 |
|---|---|
| **表セルに末尾制御文字** | `table.Cell(1,1).Range.Text` の末尾に `\r\x07` が付く。必ず除去する: `.replace('\r\x07','').replace('\x07','').strip()` |
| **`Paragraphs.Count` は表セル内も含む** | 表のセルも段落としてカウントされるため、本文の段落数とは一致しない |
| **ドキュメント末尾への挿入** | `doc.Content.End` をそのまま使うと範囲外になる。`doc.Content.End - 1` を使う: `doc.Range(doc.Content.End - 1, doc.Content.End - 1)` |
| **表セルへのテキスト設定** | `table.Cell(row, col).Range.Text = '値'` |

### PowerPoint 固有

| 項目 | 説明 |
|---|---|
| **`ExportAsFixedFormat` は使用禁止** | win32com の gencache 環境で `TypeError` が発生する。PDF 出力には `pres.SaveAs(path, 32)` を使う（`32 = ppSaveAsPDF`） |
| **表セルへのテキスト設定** | Word と異なり `.Range` はない。`.Shape.TextFrame.TextRange.Text` を経由する: `table.Cell(row, col).Shape.TextFrame.TextRange.Text = '値'` |
| **スライドレイアウト番号** | `1` = タイトルスライド / `2` = タイトルとコンテンツ / `12` = 空白 |
| **図形タイプ番号** | `1` = 長方形 / `9` = 楕円 / テキストボックスは `AddShape` ではなく `AddTextbox` を使う |

### Excel 固有

| 項目 | 説明 |
|---|---|
| **セル値の種別** | `.Value` = 計算済みの値 / `.Formula` = 数式文字列（例: `"=SUM(A1:A10)"`） |
| **一括書き込み** | 行列データは `df_to_xl(df, ws.Range('A1'))` ヘルパーを使うと効率的 |

## プロジェクト構成

```
auto-office-editor-mcp/
├── src/auto_office_editor_mcp/
│   ├── server.py        # FastMCP サーバー定義・ツール登録
│   ├── app_manager.py   # Office アプリの起動・終了・ファイル開閉管理
│   ├── executor.py      # execute_script の exec 実行環境
│   ├── helpers.py       # 色変換・単位変換・DataFrame ↔ Office 表変換
│   └── tools/
│       ├── file_ops.py  # open / close / list_open_files
│       ├── info.py      # get_xl_sheets / get_ppt_shapes / get_wd_document 等
│       └── preview.py   # Word/PPT → PDF → PNG 変換
├── tests/
│   ├── conftest.py      # フィクスチャ（app_manager / xl_open / wd_open / ppt_open 等）
│   ├── fixtures/
│   │   ├── test.xlsx    # Excel テスト用フィクスチャ
│   │   ├── test.docx    # Word テスト用フィクスチャ
│   │   └── test.pptx    # PowerPoint テスト用フィクスチャ
│   ├── test_helpers.py      # 純粋ユニットテスト（Office 不要）
│   ├── test_excel.py        # Excel 統合テスト
│   ├── test_excel_edit.py   # Excel 編集テスト
│   ├── test_word.py         # Word 統合テスト
│   ├── test_powerpoint.py   # PowerPoint 統合テスト
│   └── test_preview.py      # プレビュー統合テスト
└── scripts/
    ├── create_test_docx.py  # test.docx を win32com で生成するスクリプト
    └── create_test_pptx.py  # test.pptx を win32com で生成するスクリプト
```

## テスト

Office アプリがインストールされた環境で以下を実行します。

```bash
# 統合テスト（Excel / Word / PowerPoint が必要）
uv run pytest -m integration -v

# ユニットテストのみ（Office 不要）
uv run pytest tests/test_helpers.py -v

# 全テスト
uv run pytest -v
```

### テスト用フィクスチャファイルの再生成

```bash
uv run python scripts/create_test_docx.py
uv run python scripts/create_test_pptx.py
```

## トラブルシューティング

### `win32com.gen_py... has no attribute 'CLSIDToClassMap'` / `CLSIDToPackageMap`

`pywin32` の COM 型ライブラリキャッシュ (`gen_py`) が壊れていると発生します。
このリポジトリでは起動時に毎回 `gen_py` を削除してから再生成します。
既存の壊れた `uvx` キャッシュ環境が残っている場合は、手動削除も有効です。

PowerShell:

```powershell
Remove-Item -Recurse -Force $env:TEMP\gen_py -ErrorAction SilentlyContinue
uv cache clean
uvx --from git+https://github.com/yayoimizuha/auto-office-editor-mcp.git auto-office-editor-mcp
```

それでも再現する場合は、まずローカル checkout で `uv run python -m auto_office_editor_mcp.server`
を試して、`uvx` の一時環境依存かどうかを切り分けてください。

## ライセンス

MIT
