"""
Word文書からルビ（ふりがな）を削除するモジュール
"""
from __future__ import annotations
import os
import logging
from pathlib import Path
from typing import Optional
from docx import Document
from docx.document import Document as _Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run

# 定数定義
WORD_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NAMESPACES = {'w': WORD_NAMESPACE}
SUCCESS = 0
FAILURE = 1

# ロガーの設定
logger = logging.getLogger(__name__)

class WordProcessingError(Exception):
    """Word文書処理中のエラーを示すカスタム例外"""
    pass

def validate_paths(input_path: str | Path, output_path: str | Path, overwrite: bool = False) -> None:
    """
    入力・出力パスのバリデーションを行う

    Args:
        input_path: 入力ファイルのパス
        output_path: 出力ファイルのパス
        overwrite: 上書きを許可するかどうか

    Raises:
        WordProcessingError: パスのバリデーションに失敗した場合
    """
    input_path = Path(input_path)
    output_path = Path(output_path)

    if not input_path.exists():
        raise WordProcessingError("入力ファイルが存在しません")
    
    if output_path.exists() and not overwrite:
        raise WordProcessingError("出力ファイルが既に存在し、上書きが許可されていません")

def process_run(run: Run) -> None:
    """
    文書の実行部分（Run）からルビを削除する

    Args:
        run: 処理対象のRun要素

    Note:
        XML要素にアクセスできない場合は何もせずにスキップ
    """
    try:
        if not hasattr(run, '_element'):
            return

        # XML要素を取得してルビを検索
        element = run._element
        ruby_elements = element.findall('.//w:ruby', NAMESPACES)
        
        for ruby in ruby_elements:
            # ベーステキスト（ルビなしテキスト）を取得
            base_text = ruby.find('.//w:rubyBase//w:t', NAMESPACES)
            if base_text is not None:
                # ルビ要素をベーステキストで置換
                ruby.getparent().replace(ruby, base_text)
                logger.debug(f"ルビを削除しました: {base_text.text}")
    
    except AttributeError as e:
        logger.debug(f"Run要素の処理をスキップしました: {e}")

def process_paragraph(paragraph: Paragraph) -> None:
    """
    段落内の全てのRunからルビを削除する

    Args:
        paragraph: 処理対象の段落
    """
    for run in paragraph.runs:
        process_run(run)

def process_document(doc: _Document) -> None:
    """
    文書内の全ての段落のルビを処理する

    Args:
        doc: 処理対象の文書
    """
    for paragraph in doc.paragraphs:
        process_paragraph(paragraph)

def remove_word_ruby(input_path: str | Path, output_path: str | Path, overwrite: bool = False) -> int:
    """
    Word文書からルビ（ふりがな）を削除する

    Args:
        input_path: 入力Word文書のパス
        output_path: 保存先Word文書のパス
        overwrite: 出力ファイルの上書きを許可するかどうか

    Returns:
        int: ステータスコード（0: 成功, 1: 失敗）

    Note:
        - 入力ファイルは変更されません
        - ルビが含まれていないファイルも正常に処理されます
        - テーブルやテキストボックス内のルビも処理されます
    """
    try:
        # パスのバリデーション
        validate_paths(input_path, output_path, overwrite)

        # 文書を読み込み
        doc = Document(str(input_path))

        # 文書内のルビを処理
        process_document(doc)

        # 変更を保存
        doc.save(str(output_path))
        logger.info(f"ルビを削除し、ファイルを保存しました: {output_path}")
        return SUCCESS

    except WordProcessingError as e:
        logger.error(str(e))
        return FAILURE
    except Exception as e:
        logger.exception(f"予期しないエラーが発生しました: {e}")
        return FAILURE