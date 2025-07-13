"""
pptx翻訳スクリプト
日本語のスライドショーを、レイアウト等は変えずに文字だけ翻訳する
"""

import logging
import sys
import argparse
from pathlib import Path
from typing import List, Optional
import shutil
import json
from dataclasses import dataclass
from pptx import Presentation
from google import genai


@dataclass
class RunStyleInfo:
    """runのスタイル情報を格納するデータクラス"""
    text: str
    font_name: Optional[str] = None
    font_size: Optional[int] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    color_rgb: Optional[tuple] = None
    color_theme: Optional[int] = None
    color_brightness: Optional[float] = None

@dataclass
class TextPosition:
    """テキストの位置情報を格納するデータクラス"""
    slide_idx: int
    shape_idx: int
    para_idx: int
    original_text: str
    run_styles: List[RunStyleInfo]

@dataclass
class TextPositionPair:
    """テキストと位置情報のペアを格納するデータクラス"""
    text: str
    position: TextPosition

@dataclass
class PresentationData:
    """プレゼンテーション全体のデータを格納するデータクラス"""
    all_pairs: List[TextPositionPair]


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class PPTXTranslator:
    def __init__(self, source_lang: str = "ja", target_lang: str = "en", model_name: str = "gemini-2.5-flash", logger=logger):
        self.logger = logger
        self.source_lang = source_lang
        self.target_lang = target_lang
        
        # JSON Schemaを定義
        self.translation_schema = {
            "type": "object",
            "properties": {
                "translations": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "id": {"type": "integer"},
                            "translated": {"type": "string"},
                            "run_mapping": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "original_run_index": {"type": "integer"},
                                        "translated_text": {"type": "string"}
                                    },
                                    "required": ["original_run_index", "translated_text"]
                                }
                            }
                        },
                        "required": ["id", "translated", "run_mapping"]
                    }
                }
            },
            "required": ["translations"]
        }
        
        self.genai_client = genai.Client()
        self.model_name = model_name

        self.lang_map = {
            "ja": "Japanese",
            "en": "English",
            "ko": "Korean",
            "zh": "Chinese",
            "es": "Spanish",
            "fr": "French",
            "de": "German"
        }
        assert self.source_lang in self.lang_map, f"Unsupported source language: {self.source_lang}"
        assert self.target_lang in self.lang_map, f"Unsupported target language: {self.target_lang}"

    def extract_all_texts_with_positions(self, pptx_path: str) -> PresentationData:
        """PPTXファイルから全テキストと位置情報を抽出"""
        prs = Presentation(pptx_path)
        all_pairs = []
        
        for slide_idx, slide in enumerate(prs.slides):
            for shape_idx, shape in enumerate(slide.shapes):
                if hasattr(shape, 'text_frame') and shape.text_frame:
                    for para_idx, paragraph in enumerate(shape.text_frame.paragraphs):
                        text = paragraph.text.strip()
                        if not text:
                            continue
                        
                        # runのスタイル情報を収集
                        run_styles = []
                        for run in paragraph.runs:
                            if not run.text:
                                continue
                            
                            font = run.font
                            
                            # フォントサイズの取得
                            font_size_pt = None
                            if hasattr(font, 'size') and font.size is not None:
                                font_size_pt = font.size.pt
                            
                            # 色情報の取得
                            color_rgb = None
                            color_theme = None
                            color_brightness = None
                            if hasattr(font, 'color') and font.color:
                                color_obj = font.color
                                if hasattr(color_obj, 'rgb') and color_obj.rgb:
                                    rgb = color_obj.rgb
                                    color_rgb = (rgb.r, rgb.g, rgb.b)
                                elif hasattr(color_obj, 'theme_color') and color_obj.theme_color is not None:
                                    color_theme = int(color_obj.theme_color)
                                    if hasattr(color_obj, 'brightness') and color_obj.brightness is not None:
                                        color_brightness = float(color_obj.brightness)
                            
                            run_style = RunStyleInfo(
                                text=run.text,
                                font_name=getattr(font, 'name', None),
                                font_size=font_size_pt,
                                bold=getattr(font, 'bold', None),
                                italic=getattr(font, 'italic', None),
                                underline=getattr(font, 'underline', None),
                                color_rgb=color_rgb,
                                color_theme=color_theme,
                                color_brightness=color_brightness
                            )
                            run_styles.append(run_style)
                        
                        position = TextPosition(
                            slide_idx=slide_idx,
                            shape_idx=shape_idx,
                            para_idx=para_idx,
                            original_text=text,
                            run_styles=run_styles
                        )
                        
                        pair = TextPositionPair(text=text, position=position)
                        all_pairs.append(pair)
        
        return PresentationData(all_pairs=all_pairs)
    
    def translate_presentation_with_gemini(self, presentation_data: PresentationData) -> tuple[List[str], List[List[dict]]]:
        """Gemini APIを使用したプレゼンテーション全体の翻訳"""
        if not presentation_data.all_pairs:
            return [], []
        
        texts = [pair.text for pair in presentation_data.all_pairs]
        self.logger.info(f"翻訳対象テキスト数: {len(texts)}")

        source_lang_name = self.lang_map.get(self.source_lang, self.source_lang)
        target_lang_name = self.lang_map.get(self.target_lang, self.target_lang)
        
        # 翻訳対象テキストとrun情報をJSON用に整形
        texts_data = []
        for i, pair in enumerate(presentation_data.all_pairs):
            text_data = {
                "id": i,
                "text": pair.text,
                "runs": [
                    {
                        "index": j,
                        "text": run_style.text,
                        "style_info": f"font:{run_style.font_name}, size:{run_style.font_size}",
                        "is_whitespace": run_style.text.strip() == "",
                        "has_leading_space": run_style.text.startswith(" ") or run_style.text.startswith("\t"),
                        "has_trailing_space": run_style.text.endswith(" ") or run_style.text.endswith("\t")
                    }
                    for j, run_style in enumerate(pair.position.run_styles)
                ]
            }
            texts_data.append(text_data)
        
        texts_json = json.dumps(texts_data, ensure_ascii=False, indent=2)
        
        prompt = f"""
Translate the following texts from {source_lang_name} to {target_lang_name}.
Maintain the same tone and style. Keep technical terms consistent.

IMPORTANT: Each text consists of multiple "runs" with different styles. You must provide a "run_mapping" that maps parts of your translation to specific original run indices.

CRITICAL - SPACE PRESERVATION:
- Pay attention to spaces, whitespace, and punctuation in the original runs
- If a run contains only spaces or whitespace, preserve equivalent spacing in translation
- If a run has leading or trailing spaces (has_leading_space/has_trailing_space flags), preserve those spaces in the translation
- Ensure proper spacing between words and phrases in the translation
- The concatenation of all "translated_text" must equal the full "translated" text exactly

Input texts with run information:
{texts_json}

For each translation, provide:
1. "translated": The complete translated text
2. "run_mapping": An array mapping translation parts to original runs
   - "original_run_index": The index of the original run (0-based)
   - "translated_text": The part of translation that should use this run's style

EXAMPLES:
1. If original runs are: ["Hello", " ", "World"]
   Your run_mapping might be: [
     {{"original_run_index": 0, "translated_text": "こんにちは"}},
     {{"original_run_index": 1, "translated_text": " "}},
     {{"original_run_index": 2, "translated_text": "世界"}}
   ]

2. If original runs are: ["Hello ", "World"] (note trailing space in first run)
   Your run_mapping might be: [
     {{"original_run_index": 0, "translated_text": "こんにちは "}},
     {{"original_run_index": 1, "translated_text": "世界"}}
   ]

Return the translations in the specified JSON schema format.
"""
        
        self.logger.info("翻訳リクエストを送信中...")
        response = self.genai_client.models.generate_content(
            model=self.model_name,
            contents=prompt,
            config={
                "response_mime_type": "application/json",
                "response_schema": self.translation_schema,
            }
        )

        # JSONレスポンスを解析
        translations, run_mappings = self._parse_json_translations_with_mapping(response.text, presentation_data.all_pairs)
        return translations, run_mappings
    
    def _parse_json_translations_with_mapping(self, response_text: str, all_pairs: List[TextPositionPair]) -> tuple[List[str], List[List[dict]]]:
        """JSON形式の翻訳レスポンス（run_mapping付き）をパース"""
        try:
            response_data = json.loads(response_text)
            translations_data = response_data.get("translations", [])
            
            # IDでソートして順序を保証
            translations_data.sort(key=lambda x: x.get("id", 0))
            
            translations = []
            run_mappings = []
            
            for i, pair in enumerate(all_pairs):
                found_translation = None
                found_run_mapping = []
                
                for trans_item in translations_data:
                    if trans_item.get("id") == i:
                        found_translation = trans_item.get("translated", "")
                        found_run_mapping = trans_item.get("run_mapping", [])
                        break
                
                if found_translation is None:
                    # デフォルトの処理
                    found_translation = pair.text
                    found_run_mapping = [
                        {"original_run_index": j, "translated_text": style.text}
                        for j, style in enumerate(pair.position.run_styles)
                    ]
                else:
                    # run mappingの整合性をチェック
                    found_run_mapping = self._validate_run_mapping(found_translation, found_run_mapping, pair.position.run_styles)
                
                translations.append(found_translation)
                run_mappings.append(found_run_mapping)
            
            self.logger.info(f"翻訳完了: {len(translations)} 件")
            return translations, run_mappings
            
        except json.JSONDecodeError as e:
            self.logger.error(f"JSONパースエラー: {e}")
            # フォールバック処理
            translations = [pair.text for pair in all_pairs]
            run_mappings = [
                [{"original_run_index": j, "translated_text": style.text}
                 for j, style in enumerate(pair.position.run_styles)]
                for pair in all_pairs
            ]
            return translations, run_mappings
    
    def _validate_run_mapping(self, translated_text: str, run_mapping: List[dict], original_run_styles: List[RunStyleInfo]) -> List[dict]:
        """run mappingの整合性をチェックし、必要に応じて修正"""
        try:
            # run mappingのテキストを結合して翻訳テキストと一致するかチェック
            concatenated = "".join([item.get("translated_text", "") for item in run_mapping])
            
            if concatenated == translated_text:
                # 整合性OK
                return run_mapping
            
            self.logger.warning(f"run mapping不整合: '{concatenated}' != '{translated_text}'")
            
            # スペースの問題を修正する試行
            fixed_mapping = self._fix_space_issues(translated_text, run_mapping, original_run_styles)
            if fixed_mapping:
                return fixed_mapping
            
            # 修正できない場合はフォールバック
            self.logger.warning("run mapping修正失敗、フォールバックを使用")
            return self._create_fallback_mapping(translated_text, original_run_styles)
            
        except Exception as e:
            self.logger.error(f"run mapping検証エラー: {e}")
            return self._create_fallback_mapping(translated_text, original_run_styles)
    
    def _fix_space_issues(self, translated_text: str, run_mapping: List[dict], original_run_styles: List[RunStyleInfo]) -> List[dict]:
        """スペースの問題を修正"""
        try:
            # run mappingから翻訳テキストを結合
            mapped_text = "".join([item.get("translated_text", "") for item in run_mapping])
            
            # 既に一致している場合は修正不要
            if mapped_text == translated_text:
                return run_mapping
            
            
            # スペースが欠落している可能性をチェック
            mapped_indices = {item.get("original_run_index", -1) for item in run_mapping}
            missing_runs = []
            
            for i, style in enumerate(original_run_styles):
                if i not in mapped_indices:
                    # 欠落しているrunを特定
                    if (style.text.strip() == "" or  # スペースのみのrun
                        style.text.startswith(" ") or style.text.endswith(" ") or  # 先頭・末尾にスペース
                        style.text.startswith("\t") or style.text.endswith("\t")):  # タブ
                        missing_runs.append(i)
            
            if missing_runs:
                # 欠落したrunを追加して修正を試行
                fixed_mapping = list(run_mapping)
                for run_index in missing_runs:
                    run_text = original_run_styles[run_index].text
                    fixed_mapping.append({
                        "original_run_index": run_index,
                        "translated_text": run_text
                    })
                
                # インデックスでソート
                fixed_mapping.sort(key=lambda x: x.get("original_run_index", 0))
                
                # 結合して確認
                concatenated = "".join([item.get("translated_text", "") for item in fixed_mapping])
                if concatenated == translated_text:
                    self.logger.info("先頭・末尾スペース問題を修正しました")
                    return fixed_mapping
            
            # 別のアプローチ: 翻訳テキストの先頭・末尾スペースをチェック
            if translated_text != mapped_text:
                # 先頭・末尾にスペースが必要か判定
                needs_leading_space = translated_text.startswith(" ") and not mapped_text.startswith(" ")
                needs_trailing_space = translated_text.endswith(" ") and not mapped_text.endswith(" ")
                
                if needs_leading_space or needs_trailing_space:
                    # 適切なrunにスペースを追加
                    fixed_mapping = self._add_missing_spaces(run_mapping, needs_leading_space, needs_trailing_space)
                    if fixed_mapping:
                        concatenated = "".join([item.get("translated_text", "") for item in fixed_mapping])
                        if concatenated == translated_text:
                            self.logger.info("先頭・末尾スペースを追加しました")
                            return fixed_mapping
            
            return None
            
        except Exception as e:
            self.logger.error(f"スペース修正エラー: {e}")
            return None
    
    def _add_missing_spaces(self, run_mapping: List[dict], needs_leading: bool, needs_trailing: bool) -> List[dict]:
        """欠落したスペースをrun mappingに追加"""
        try:
            fixed_mapping = list(run_mapping)
            
            if needs_leading:
                # 最初のrunに先頭スペースを追加
                if fixed_mapping:
                    first_item = fixed_mapping[0]
                    first_item["translated_text"] = " " + first_item.get("translated_text", "")
            
            if needs_trailing:
                # 最後のrunに末尾スペースを追加
                if fixed_mapping:
                    last_item = fixed_mapping[-1]
                    last_item["translated_text"] = last_item.get("translated_text", "") + " "
            
            return fixed_mapping
            
        except Exception as e:
            self.logger.error(f"スペース追加エラー: {e}")
            return None
    
    def _create_fallback_mapping(self, translated_text: str, original_run_styles: List[RunStyleInfo]) -> List[dict]:
        """フォールバックのrun mappingを作成"""
        if not original_run_styles:
            return [{"original_run_index": 0, "translated_text": translated_text}]
        
        # 最初のrunに全テキストを割り当て
        return [{"original_run_index": 0, "translated_text": translated_text}]
    
    def apply_run_style(self, run, run_style: RunStyleInfo):
        """runにスタイルを適用"""
        try:
            font = run.font
            
            # フォント名
            if run_style.font_name:
                font.name = run_style.font_name
            
            # フォントサイズ
            if run_style.font_size:
                from pptx.util import Pt
                font.size = Pt(run_style.font_size)
            
            # 太字・斜体・下線
            if run_style.bold is not None:
                font.bold = run_style.bold
            if run_style.italic is not None:
                font.italic = run_style.italic
            if run_style.underline is not None:
                font.underline = run_style.underline
            
            # 色の適用
            if run_style.color_rgb is not None:
                from pptx.dml.color import RGBColor
                r, g, b = run_style.color_rgb
                font.color.rgb = RGBColor(r, g, b)
            elif run_style.color_theme is not None:
                from pptx.enum.dml import MSO_THEME_COLOR
                theme_colors = {
                    0: MSO_THEME_COLOR.BACKGROUND_1,
                    1: MSO_THEME_COLOR.TEXT_1,
                    2: MSO_THEME_COLOR.BACKGROUND_2,
                    3: MSO_THEME_COLOR.TEXT_2,
                    4: MSO_THEME_COLOR.ACCENT_1,
                    5: MSO_THEME_COLOR.ACCENT_2,
                    6: MSO_THEME_COLOR.ACCENT_3,
                    7: MSO_THEME_COLOR.ACCENT_4,
                    8: MSO_THEME_COLOR.ACCENT_5,
                    9: MSO_THEME_COLOR.ACCENT_6,
                    10: MSO_THEME_COLOR.HYPERLINK,
                    11: MSO_THEME_COLOR.FOLLOWED_HYPERLINK,
                    12: MSO_THEME_COLOR.DARK_1,
                    13: MSO_THEME_COLOR.LIGHT_1,
                    14: MSO_THEME_COLOR.DARK_2,
                    15: MSO_THEME_COLOR.LIGHT_2
                }
                if run_style.color_theme in theme_colors:
                    font.color.theme_color = theme_colors[run_style.color_theme]
                    if run_style.color_brightness is not None:
                        font.color.brightness = run_style.color_brightness
                        
        except Exception as e:
            self.logger.debug(f"スタイル適用エラー: {e}")
    
    def replace_text_with_run_mapping(self, paragraph, translated_text: str, run_styles: List[RunStyleInfo], run_mapping: List[dict]):
        """run_mappingを使用してスタイルを保持してテキストを置換"""
        try:
            if not run_styles or not run_mapping:
                # スタイル情報やマッピング情報がない場合は単純置換
                paragraph.text = translated_text
                return
            
            # paragraphをクリア
            paragraph.text = ""
            
            # run_mappingをoriginal_run_indexでソート
            sorted_mapping = sorted(run_mapping, key=lambda x: x.get("original_run_index", 0))
            
            # 新しいrunを作成してスタイルを適用
            for mapping_item in sorted_mapping:
                original_run_index = mapping_item.get("original_run_index", 0)
                translated_part = mapping_item.get("translated_text", "")
                
                if not translated_part:
                    continue
                
                # 対応する元のスタイルを取得
                if 0 <= original_run_index < len(run_styles):
                    original_style = run_styles[original_run_index]
                else:
                    # インデックスが範囲外の場合は最初のスタイルを使用
                    original_style = run_styles[0] if run_styles else None
                
                # 新しいrunを追加
                run = paragraph.add_run()
                run.text = translated_part
                
                # 元のスタイルを適用
                if original_style:
                    self.apply_run_style(run, original_style)
        
        except Exception as e:
            # エラーの場合は単純置換にフォールバック
            self.logger.error(f"run mapping置換エラー: {e}")
            paragraph.text = translated_text
    
    def translate_pptx(self, input_path: str, output_path: str) -> bool:
        """PPTXファイルを翻訳"""
        try:    
            # ファイルをコピー
            self.logger.info(f"ファイルをコピー中: {input_path} -> {output_path}")
            shutil.copy2(input_path, output_path)
            
            # 全テキストと位置情報を抽出
            presentation_data = self.extract_all_texts_with_positions(input_path)
            
            if not presentation_data.all_pairs:
                self.logger.warning("翻訳対象のテキストが見つかりません")
                return False

            self.logger.info(f"翻訳開始: {input_path}")
            self.logger.info(f"翻訳対象テキスト数: {len(presentation_data.all_pairs)}")

            # プレゼンテーション全体を翻訳
            translations, run_mappings = self.translate_presentation_with_gemini(presentation_data)

            # コピーしたファイルを開いて翻訳を適用
            prs = Presentation(output_path)
            
            for i, (translation, run_mapping, pair) in enumerate(zip(translations, run_mappings, presentation_data.all_pairs)):
                try:
                    position = pair.position
                    
                    # 対象のシェイプと段落を取得
                    slide = prs.slides[position.slide_idx]
                    shape = slide.shapes[position.shape_idx]
                    paragraph = shape.text_frame.paragraphs[position.para_idx]
                    
                    # run mappingを使用してテキストを置換
                    self.replace_text_with_run_mapping(paragraph, translation, position.run_styles, run_mapping)
                    
                except Exception as e:
                    self.logger.error(f"テキスト置換エラー (インデックス {i}): {e}")
                    continue

            prs.save(output_path)
            self.logger.info(f"翻訳完了: {output_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"翻訳処理エラー: {e}")
            return False


def main():
    parser = argparse.ArgumentParser(description="PPTX翻訳スクリプト (Gemini API使用)")
    parser.add_argument("input_file", help="入力PPTXファイル")
    parser.add_argument("-o", "--output", help="出力PPTXファイル")
    parser.add_argument("-s", "--source", default="ja", help="翻訳元言語 (デフォルト: ja)")
    parser.add_argument("-t", "--target", default="en", help="翻訳先言語 (デフォルト: en)")
    parser.add_argument("-m", "--model", default="gemini-2.5-flash", help="使用するモデル名 (デフォルト: gemini-2.5-flash)")

    args = parser.parse_args()
    
    # 入力ファイルの確認
    if not Path(args.input_file).exists():
        print(f"エラー: ファイルが見つかりません: {args.input_file}")
        sys.exit(1)
    
    # 出力ファイル名の生成
    if args.output:
        output_file = args.output
    else:
        input_path = Path(args.input_file)
        output_file = str(input_path.parent / f"{input_path.stem}_{args.target}.pptx")

    # 翻訳処理
    translator = PPTXTranslator(args.source, args.target, args.model)
    success = translator.translate_pptx(args.input_file, output_file)
    
    if success:
        print("翻訳が完了しました。")
    else:
        print("翻訳に失敗しました。")
        sys.exit(1)


if __name__ == "__main__":
    main()