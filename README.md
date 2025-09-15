# Office Ruby Remover

Office Ruby Removerは、Microsoft WordドキュメントとExcelファイルからルビ（ふりがな）を削除するPythonライブラリです。

## 特徴

- Word文書（.docx）からのルビ削除
- Excelファイル（.xlsx）からのルビ削除
- クロスプラットフォーム対応（Windows/Linux/Mac）

## インストール

```bash
pip install office-ruby-remover
```

## 使用方法

### Wordドキュメントのルビ削除

```python
from word.remove_word_ruby import remove_word_ruby

# 基本的な使用方法
result = remove_word_ruby(
    input_path="input.docx",
    output_path="output.docx"
)

# 上書きを許可する場合
result = remove_word_ruby(
    input_path="input.docx",
    output_path="output.docx",
    overwrite=True
)
```

### Excelファイルのルビ削除

```python
from excel.remove_excel_ruby import remove_excel_ruby

# 基本的な使用方法
result = remove_excel_ruby(
    input_path="input.xlsx",
    output_path="output.xlsx"
)

# 上書きを許可する場合
result = remove_excel_ruby(
    input_path="input.xlsx",
    output_path="output.xlsx",
    overwrite=True
)
```

## パラメータ

### 共通パラメータ

- `input_path` (str | Path): 入力ファイルのパス
- `output_path` (str | Path): 出力ファイルのパス
- `overwrite` (bool): 出力ファイルが存在する場合に上書きを許可するかどうか（デフォルト: False）

### 戻り値

- `0`: 成功
- `1`: 失敗（ファイルが見つからない、権限エラーなど）

## エラーハンドリング

```python
from excel.remove_excel_ruby import remove_excel_ruby
import logging

# ログレベルの設定
logging.basicConfig(level=logging.INFO)

try:
    result = remove_excel_ruby("input.xlsx", "output.xlsx")
    if result == 0:
        print("ルビの削除が成功しました")
    else:
        print("ルビの削除に失敗しました")
except Exception as e:
    print(f"エラーが発生しました: {e}")
```

## 制限事項

- Word文書: .docx形式のみ対応（.doc形式は非対応）
- Excelファイル: .xlsx形式のみ対応（.xls形式は非対応）
- 保護されたドキュメントは処理できない場合があります

## ライセンス

本プロジェクトはMITライセンスの下で公開されています。詳細は[LICENSE](LICENSE)ファイルをご覧ください。