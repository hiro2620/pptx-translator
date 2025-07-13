# PDF Translator

PowerPointプレゼンテーション(.pptx)ファイルを翻訳するツールです。Google Gemini APIを使用し、文脈を維持して翻訳を行います。

日本語から英語への翻訳を主にサポートしていますが、他の言語にも対応しています。

## 機能

- レイアウトを保持したままの翻訳
- コンテキストを考慮した一貫性のある翻訳
- 元テキストの長さと文体を保持した翻訳
- 専門用語の統一性を維持
- 文字装飾を意味的に対応するように保持

## インストール

### uvを使用する場合

1. uvをインストール
https://docs.astral.sh/uv/getting-started/installation に従ってインストール

2. リポジトリをクローンして依存関係をインストール
```bash
git clone https://github.com/hiro2620/pptx-translator.git
cd pptx-translator
uv sync
```

3. 仮想環境を有効化
```bash
source .venv/bin/activate
```

### venv を使用する場合

1. 仮想環境を作成
```bash
python -m venv .venv
```

2. 仮想環境を有効化
```bash
# macOS/Linux
source .venv/bin/activate

# Windows
.venv\Scripts\activate
```

3. 依存関係をインストール
```bash
pip install -r requirements.txt
```

## API設定

Google Gemini APIを使用するためのAPIキーが必要です。

1. Google AI Studio (https://aistudio.google.com/apikey) でAPIキーを取得
2. 環境変数に設定

```bash
export GEMINI_API_KEY="your-api-key-here"
```

または `.env` ファイルを作成して設定
```
GEMINI_API_KEY=your-api-key-here
```
`.env`ファイルの自動読み込みには対応していないので、以下のコマンドで手動で読み込む必要があります。
```bash
source ./.env
```

## 使用方法

### 基本的な使用方法
```bash
python main.py input.pptx
```

uvを使用する場合は
```bash
uv run main.py input.pptx
```
のように実行することもできます。

日本語から英語に翻訳されたファイルが `input_en.pptx` として保存されます。

### 詳細オプション

```bash
python main.py input.pptx -o output.pptx -s ja -t en -m gemini-2.5-flash
```

#### コマンドライン引数

- `input_file`: 翻訳するPPTXファイルのパス
- `-o, --output`: 出力ファイル名(デフォルト: `<input_file_name>_<target_lang>.pptx`)
- `-s, --source`: 翻訳元言語(デフォルト: ja)
- `-t, --target`: 翻訳先言語(デフォルト: en)
- `-m, --model`: 使用するGeminiモデル(デフォルト: gemini-2.5-flash)

#### サポートされている言語

- `ja`: 日本語
- `en`: 英語
- `ko`: 韓国語
- `zh`: 中国語
- `es`: スペイン語
- `fr`: フランス語
- `de`: ドイツ語

### 使用例

```bash
# 日本語から英語への翻訳
python main.py presentation.pptx

# 日本語から韓国語へ
python main.py presentation.pptx -t ko

# 英語から日本語へ
python main.py presentation.pptx -s en -t ja

# カスタム出力ファイル名
python main.py presentation.pptx -o translated_presentation.pptx

# 別のモデルを使用
python main.py presentation.pptx -m gemini-2.0-flash-exp
```

## 対応要素

### テキスト要素

- テキストボックス内のテキスト
- 図形内のテキスト
- 表の中のテキスト

### 保持される要素

- レイアウト
- フォントスタイル（太字、斜体、下線など）
- フォントサイズ
- 画像や図形
- テーマを考慮した色やRGB設定など
- 色の塗りつぶし設定

### 制限事項

- PowerPointファイル(.pptx)のみ対応
- 文字の色が引き継がれないことがあります。
- 複雑なレイアウトや図形では文字が検出されないことがあります。