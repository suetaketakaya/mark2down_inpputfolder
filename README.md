# Microsoft Office ファイル to Markdown 変換ツール

MicrosoftのOfficeファイル（Word、Excel、PowerPoint、PDF）をMarkdownに変換するPythonツールです。

## 対応ファイル形式

- `.docx` - Microsoft Word
- `.pptx` - Microsoft PowerPoint
- `.xlsx` - Microsoft Excel
- `.pdf` - PDF

## 必要な環境

- Python 3.9以上
- pip (Pythonパッケージマネージャー)

## インストール

### 1. 依存関係のインストール

```bash
pip install python-pptx mammoth openpyxl pdfminer.six markdownify
```

## 使い方

### 基本的な使用方法

1. 変換したいファイルを `./input` フォルダに配置
2. 以下のコマンドを実行:

```bash
python convert_to_markdown.py
```

3. 変換されたMarkdownファイルは `./output` フォルダに保存されます

### カスタムディレクトリを指定する場合

```bash
python convert_to_markdown.py <入力ディレクトリ> <出力ディレクトリ>
```

例:
```bash
python convert_to_markdown.py ./my_files ./converted_files
```

## プロジェクト構造

```
microsoft_markit2down/
├── convert_to_markdown.py  # メイン変換スクリプト
├── input/                   # 変換元ファイルを配置するフォルダ
├── output/                  # 変換後のMarkdownファイルが保存されるフォルダ
├── markitdown/             # Microsoft MarkItDown リポジトリ（参考用）
└── README.md               # このファイル
```

## 機能

### 自動整形機能

変換後のMarkdownファイルを自動的に整形します：

1. **1文字ずつ改行された文字の結合**
   - PDFから抽出される際に1文字ずつ改行される問題を自動修正
   - 例: `©` `2` `0` `2` `5` → `©2025`

2. **スペース区切り文字の結合**
   - 見出しなどでスペース区切りになっている文字を自動結合
   - 例: `自 己 紹 介` → `自己紹介`
   - 例: `A I 領 域` → `AI領域`

3. **過剰な空行の削減**
   - 3行以上の連続する空行を2行以内に制限
   - ドキュメントの可読性を向上

4. **前後の空白削除**
   - 各行の先頭・末尾の不要な空白を削除

## 変換例

### Word文書 (.docx)
- テキスト、見出し、リストなどをMarkdown形式に変換
- HTMLを経由してMarkdownに変換するため、書式が保持されます
- 変換後に自動整形が適用されます

### PowerPoint (.pptx)
- 各スライドを見出しとして扱い、テキストを抽出
- スライド番号が自動的に付与されます
- 変換後に自動整形が適用されます

### Excel (.xlsx)
- 各シートをMarkdownテーブルとして変換
- セルの値を保持してテーブル形式で出力
- テーブル構造を保持したまま整形されます

### PDF (.pdf)
- テキストを抽出してMarkdown形式で出力
- レイアウト情報も可能な限り保持
- **自動整形機能がPDFで特に効果的です**
  - 1文字ずつ改行される問題を解決
  - スペース区切りの見出しを自動修正

## トラブルシューティング

### Python 3.10以上を使用したい場合

Microsoft MarkItDownの最新版（Python 3.10以上が必要）を使用する場合:

```bash
# 仮想環境を作成（推奨）
python3.10 -m venv .venv
source .venv/bin/activate

# MarkItDownをインストール
pip install 'markitdown[all]'

# MarkItDownのCLIを使用
markitdown input/example.docx -o output/example.md
```

### エラーが発生した場合

1. 依存関係が正しくインストールされているか確認:
   ```bash
   pip list | grep -E "mammoth|python-pptx|openpyxl|pdfminer"
   ```

2. ファイルが破損していないか確認

3. ファイルが読み取り可能か確認

## 参考

このツールは [Microsoft MarkItDown](https://github.com/microsoft/markitdown) のコンセプトに基づいています。

## ライセンス

このプロジェクトはMITライセンスの下で公開されています。
