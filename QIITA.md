# Microsoft OfficeファイルをMarkdownに一括変換！自動整形機能付きPythonツールを作った

## はじめに

PDFやWordファイルの内容をMarkdownに変換したいことってありませんか？
特にPDFからテキストを抽出すると、1文字ずつ改行されたり、スペース区切りになったりして使い物にならない...そんな経験ありますよね。

そこで、**Microsoft OfficeファイルをMarkdownに変換し、自動で整形してくれるツール**を作りました！

https://github.com/suetaketakaya/mark2down_inpputfolder

## 特徴

### 対応ファイル形式

- `.docx` - Microsoft Word
- `.pptx` - Microsoft PowerPoint
- `.xlsx` - Microsoft Excel
- `.pdf` - PDF

### 自動整形機能が優秀

変換後のMarkdownを自動的に整形してくれます：

1. **1文字ずつ改行された文字を結合**
   - PDFでよくある問題を解決
   - `©` `2` `0` `2` `5` → `©2025`

2. **スペース区切り文字を結合**
   - 見出しなどの問題を自動修正
   - `自 己 紹 介` → `自己紹介`
   - `A I 領 域` → `AI領域`

3. **過剰な空行を削減**
   - 3行以上の連続する空行を2行以内に制限

4. **各行の前後の空白を削除**

## インストール

```bash
# 依存関係のインストール
pip install python-pptx mammoth openpyxl pdfminer.six markdownify
```

## 使い方

```bash
# 1. リポジトリをクローン
git clone https://github.com/suetaketakaya/mark2down_inpputfolder.git
cd mark2down_inpputfolder

# 2. 変換したいファイルを input フォルダに配置
cp your_file.pdf ./input/

# 3. 変換実行
python convert_to_markdown.py

# 4. output フォルダに変換結果が保存されます
ls output/
```

### カスタムディレクトリ指定

```bash
python convert_to_markdown.py ./my_files ./converted_files
```

## 実装のポイント

### 整形機能の実装

PDFからの変換で最も問題になるのが、1文字ずつ改行される現象です。
これを解決するために、以下のようなロジックを実装しました：

```python
class MarkdownFormatter:
    """Markdownテキストを整形するクラス"""

    @staticmethod
    def remove_single_char_spaces(text: str) -> str:
        """1文字ずつスペースで区切られた文字列を結合する"""
        lines = text.split('\n')
        result_lines = []

        for line in lines:
            # スペース区切りの文字数をカウント
            words = line.split()
            if len(words) >= 3:  # 3つ以上のトークンがある場合
                # すべてのトークンが1-2文字かチェック
                if all(len(word) <= 2 for word in words):
                    # スペースを除去して結合
                    line = ''.join(words)
            result_lines.append(line)

        return '\n'.join(result_lines)
```

### ファイル形式別の変換

各ファイル形式に応じて適切なライブラリを使用：

```python
converters = {
    '.docx': self.convert_docx,  # mammothを使用
    '.pptx': self.convert_pptx,  # python-pptxを使用
    '.xlsx': self.convert_xlsx,  # openpyxlを使用
    '.pdf': self.convert_pdf,    # pdfminer.sixを使用
}
```

## 実際の変換例

### 変換前（PDFから抽出）

```
©
︎
2
0
2
5
V
e
r
i
S
e
r
v
e

自 己 紹 介

A I 領 域 に お け る R & D 活 動
```

### 変換後（自動整形済み）

```
©︎2025VeriServe

自己紹介

AI領域におけるR&D活動
```

見違えるほど読みやすくなりました！

## プロジェクト構造

```
.
├── convert_to_markdown.py  # メイン変換スクリプト
├── README.md               # 使用方法
├── .gitignore             # Git除外設定
├── input/                  # 変換元ファイル
└── output/                 # 変換後のMarkdown
```

## パフォーマンス

実際のPDFファイル（2.7MB）を変換した結果：

- **変換前の生データ**: 22KB（可読性低）
- **整形後のMarkdown**: 19KB（可読性高）
- **削減率**: 約14%
- **行数**: 1059行

整形により、ファイルサイズも削減されつつ、可読性が大幅に向上しました。

## 応用例

### ドキュメント管理

```bash
# 複数のPDFを一括変換
cp *.pdf ./input/
python convert_to_markdown.py
```

### CI/CDへの組み込み

```yaml
# GitHub Actionsの例
- name: Convert documents to Markdown
  run: |
    pip install python-pptx mammoth openpyxl pdfminer.six markdownify
    python convert_to_markdown.py ./docs ./output
```

## まとめ

Microsoft OfficeファイルをMarkdownに変換するツールを作成しました。

**良かった点：**
- PDFの1文字改行問題を自動解決
- スペース区切り文字の自動結合
- 複数ファイル形式に対応
- シンプルで使いやすいCLI

**今後の改善案：**
- GUI版の開発
- より高度なレイアウト保持
- 画像の抽出・埋め込み
- バッチ処理の最適化

ドキュメント管理やナレッジベース構築に役立つツールになれば幸いです！

## リポジトリ

https://github.com/suetaketakaya/mark2down_inpputfolder

スターやコントリビューションお待ちしています！

## 参考

このツールは [Microsoft MarkItDown](https://github.com/microsoft/markitdown) のコンセプトに基づいています。
