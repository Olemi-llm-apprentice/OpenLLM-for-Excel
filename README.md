# OpenLLM for Excel

<p align="center">
  <img src="assets/icon-128.png" alt="OpenLLM for Excel" width="128" height="128">
</p>

<p align="center">
  <strong>Open source LLM-powered AI assistant for Microsoft Excel</strong>
</p>

<p align="center">
  <a href="#features">Features</a> •
  <a href="#installation">Installation</a> •
  <a href="#usage">Usage</a> •
  <a href="#supported-models">Supported Models</a> •
  <a href="#contributing">Contributing</a> •
  <a href="#license">License</a>
</p>

<p align="center">
  <a href="#-日本語">🇯🇵 日本語</a>
</p>

---

## ✨ Features

- 🤖 **AI Chat Assistant** - Chat with AI to get Excel advice, formula suggestions, and troubleshooting help
- 📊 **Context-Aware** - AI understands your currently selected cells and their values
- ⚡ **Code Execution** - Generate and execute Excel JavaScript API code directly
- 🔄 **Multi-Provider Support** - Works with OpenAI, Anthropic Claude, and Google Gemini
- 🖼️ **Image Generation** - Generate images using OpenAI GPT Image 1 and Gemini 3 Pro Image
- 📎 **File Upload (Vision)** - Upload images and PDFs for AI analysis
- 🌐 **Streaming Responses** - Real-time streaming for faster interaction
- 📝 **Markdown Rendering** - Beautiful formatted responses with code highlighting

## 🚀 Installation

### Prerequisites

- Node.js 18.x or later
- Microsoft Excel (Desktop or Web)
- API key from OpenAI, Anthropic, or Google AI Studio

### Setup

1. Clone this repository:
```bash
git clone https://github.com/Olemi-llm-apprentice/OpenLLM-for-Excel.git
cd OpenLLM-for-Excel
```

2. Install dependencies:
```bash
npm install
```

3. Start the development server:
```bash
npm run dev-server
```

4. Sideload the add-in in Excel:
```bash
npm run start
```

## 🔑 API Key Configuration

This add-in supports **three methods** for API key management, with the following priority:

| Priority | Method | Best For |
|:--------:|--------|----------|
| 1 | Environment Variables | Developers, CI/CD |
| 2 | Office Settings API | End Users (saved & synced) |
| 3 | UI Input Field | Quick testing |

### Method 1: Environment Variables (Recommended for Development)

1. Copy the example file:
```bash
cp .env.local.example .env.local
```

2. Edit `.env.local` with your API keys:
```bash
# OpenAI
OPENAI_API_KEY=sk-xxxxxxxxxxxxxxxxxxxxxxxx

# Anthropic Claude
ANTHROPIC_API_KEY=sk-ant-xxxxxxxxxxxxxxxxxxxxxxxx

# Google Gemini
GEMINI_API_KEY=AIzaSyxxxxxxxxxxxxxxxxxxxxxxxxxx
```

3. Restart the development server:
```bash
npm run dev-server
```

> **Note**: `.env.local` is gitignored and will never be committed.

### Method 2: Office Settings API (Recommended for End Users)

1. Enter your API key in the input field
2. Click **💾 Save** button
3. Your key is securely stored in your Microsoft account
4. The key will automatically sync across devices

> **Benefits**: 
> - Keys are stored in Microsoft's cloud (not browser localStorage)
> - Automatic sync across all your devices
> - Persists even after clearing browser data

### Method 3: UI Input (For Quick Testing)

Simply enter your API key in the input field and use immediately.

> **Note**: Keys entered this way are not saved and must be re-entered each session.

### Testing Your API Key

Click the **🔌 Test** button to verify your API key works correctly before use.

### Security Notes

⚠️ **Important Security Information**:

- API keys are sent directly from your browser to the LLM provider
- In browser DevTools (F12), API keys may be visible in Network requests
- For production/enterprise use, consider implementing a backend proxy
- Never share your API keys or commit them to version control

## 📖 Usage

1. Open Excel and go to the **Home** tab
2. Click the **OpenLLM** button to open the task pane
3. Configure your API key (see above)
4. Select your preferred model
5. Start chatting with the AI assistant!

### Chat Mode
Type your question and click **送信 (Send)** to get AI advice about Excel.

### File Upload (Vision)
1. Click the file input to select images or PDFs
2. Preview your uploaded files
3. Type your question and click **送信 (Send)**
4. AI will analyze the files and respond

### Image Generation
1. Select an image generation model (GPT Image 1 or Gemini 3 Pro Image)
2. Type your prompt describing the image
3. Click **画像生成 (Generate Image)**
4. Download the generated image

### Macro Execution Mode
Describe what you want to do and click **マクロ実行 (Execute Macro)** to generate and run Excel JavaScript API code.

## 🤖 Supported Models

### OpenAI Chat
| Model | Features |
|-------|----------|
| GPT-4o | Recommended, Vision support |
| GPT-4o mini | Vision support |
| GPT-4.1 | Latest non-reasoning model |
| GPT-4 Turbo | Legacy |

### OpenAI Image Generation
| Model | Features |
|-------|----------|
| GPT Image 1 | State-of-the-art image generation |

### Anthropic Claude
| Model | Features |
|-------|----------|
| Claude Sonnet 4.5 | Recommended, Vision + PDF support |
| Claude Haiku 4.5 | Fast, Vision + PDF support |
| Claude Opus 4.5 | Premium, Vision + PDF support |

### Google Gemini Chat
| Model | Features |
|-------|----------|
| Gemini 2.0 Flash | Recommended, Vision + PDF support |
| Gemini 2.5 Flash | Fast, Vision + PDF support |
| Gemini 2.5 Pro | Advanced reasoning, Vision + PDF support |

### Google Gemini Image Generation
| Model | Features |
|-------|----------|
| Gemini 3 Pro Image | Image generation with reasoning |

## 📎 File Support Matrix

| Provider | Image Input | PDF Input | Image Generation |
|----------|:-----------:|:---------:|:----------------:|
| OpenAI GPT-4o/mini | ✅ | ❌ | ❌ |
| OpenAI GPT Image 1 | ❌ | ❌ | ✅ |
| Claude (all models) | ✅ | ✅ | ❌ |
| Gemini (all models) | ✅ | ✅ | ❌ |
| Gemini 3 Pro Image | ❌ | ❌ | ✅ |

## 🛠️ Development

### Available Scripts

| Command | Description |
|---------|-------------|
| `npm run dev-server` | Start webpack dev server |
| `npm run build` | Build for production |
| `npm run build:dev` | Build for development |
| `npm run start` | Start debugging in Excel Desktop |
| `npm run start:web` | Start debugging in Excel Web |
| `npm run lint` | Run ESLint |
| `npm run lint:fix` | Fix ESLint errors |

### Project Structure

```
OpenLLM-for-Excel/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html    # Main UI
│   │   ├── taskpane.css     # Styles
│   │   └── taskpane.js      # Main logic (multi-provider)
│   └── commands/
│       ├── commands.html
│       └── commands.js
├── docs/                     # API documentation
│   ├── openai-models.md
│   ├── anthropic-claude-models.md
│   └── google-gemini-models.md
├── assets/                   # Icons
├── manifest.xml              # Office Add-in manifest
├── webpack.config.js
└── package.json
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with [Office Add-in](https://docs.microsoft.com/office/dev/add-ins/) framework
- Powered by [OpenAI](https://openai.com/), [Anthropic](https://anthropic.com/), and [Google AI](https://ai.google.dev/) APIs

---

# 🇯🇵 日本語

## OpenLLM for Excel

Microsoft Excel用のオープンソースLLM搭載AIアシスタント

## ✨ 機能

- 🤖 **AIチャットアシスタント** - Excelのアドバイス、数式の提案、トラブルシューティングをAIに相談
- 📊 **コンテキスト認識** - 選択中のセルとその値をAIが理解
- ⚡ **コード実行** - Excel JavaScript APIコードを生成して直接実行
- 🔄 **マルチプロバイダー対応** - OpenAI、Anthropic Claude、Google Geminiに対応
- 🖼️ **画像生成** - OpenAI GPT Image 1とGemini 3 Pro Imageで画像を生成
- 📎 **ファイルアップロード（Vision）** - 画像やPDFをアップロードしてAIに分析させる
- 🌐 **ストリーミングレスポンス** - リアルタイムストリーミングで高速な対話
- 📝 **Markdownレンダリング** - コードハイライト付きの美しいフォーマット

## 🚀 インストール

### 前提条件

- Node.js 18.x以降
- Microsoft Excel（デスクトップまたはWeb）
- OpenAI、Anthropic、またはGoogle AI StudioのAPIキー

### セットアップ

1. リポジトリをクローン:
```bash
git clone https://github.com/Olemi-llm-apprentice/OpenLLM-for-Excel.git
cd OpenLLM-for-Excel
```

2. 依存関係をインストール:
```bash
npm install
```

3. 開発サーバーを起動:
```bash
npm run dev-server
```

4. Excelにアドインをサイドロード:
```bash
npm run start
```

## 🔑 APIキーの設定

このアドインは**3つの方法**でAPIキーを管理でき、以下の優先順位で使用されます：

| 優先順位 | 方法 | 推奨用途 |
|:-------:|------|---------|
| 1 | 環境変数 | 開発者、CI/CD |
| 2 | Office Settings API | 一般ユーザー（保存＆同期） |
| 3 | UI入力フィールド | 簡易テスト |

### 方法1: 環境変数（開発者向け推奨）

1. サンプルファイルをコピー:
```bash
cp .env.local.example .env.local
```

2. `.env.local` にAPIキーを設定:
```bash
# OpenAI
OPENAI_API_KEY=sk-xxxxxxxxxxxxxxxxxxxxxxxx

# Anthropic Claude
ANTHROPIC_API_KEY=sk-ant-xxxxxxxxxxxxxxxxxxxxxxxx

# Google Gemini
GEMINI_API_KEY=AIzaSyxxxxxxxxxxxxxxxxxxxxxxxxxx
```

3. 開発サーバーを再起動:
```bash
npm run dev-server
```

> **注意**: `.env.local` は gitignore されているため、コミットされません。

### 方法2: Office Settings API（一般ユーザー向け推奨）

1. 入力フィールドにAPIキーを入力
2. **💾 保存** ボタンをクリック
3. キーはMicrosoftアカウントに安全に保存されます
4. 別のデバイスでも自動的に同期されます

> **メリット**: 
> - ブラウザのlocalStorageではなく、Microsoftクラウドに保存
> - 複数デバイス間で自動同期
> - ブラウザデータを削除しても保持される

### 方法3: UI入力（簡易テスト用）

入力フィールドにAPIキーを入力してすぐに使用できます。

> **注意**: この方法で入力したキーは保存されず、セッションごとに再入力が必要です。

### APIキーのテスト

**🔌 テスト** ボタンをクリックして、APIキーが正しく動作するか確認できます。

### セキュリティに関する注意事項

⚠️ **重要なセキュリティ情報**:

- APIキーはブラウザからLLMプロバイダーに直接送信されます
- ブラウザの開発者ツール（F12）のネットワークタブでAPIキーが見える場合があります
- 本番環境や企業での利用には、バックエンドプロキシの実装を検討してください
- APIキーを他人と共有したり、バージョン管理にコミットしたりしないでください

## 📖 使い方

1. Excelを開き、**ホーム**タブに移動
2. **OpenLLM**ボタンをクリックしてタスクペインを開く
3. APIキーを設定（上記参照）
4. 使用するモデルを選択
5. AIアシスタントとチャット開始！

### チャットモード
質問を入力して**送信**をクリックすると、ExcelについてのAIアドバイスが得られます。

### ファイルアップロード（Vision）
1. ファイル入力をクリックして画像やPDFを選択
2. アップロードしたファイルをプレビュー
3. 質問を入力して**送信**をクリック
4. AIがファイルを分析して回答

### 画像生成
1. 画像生成モデル（GPT Image 1またはGemini 3 Pro Image）を選択
2. 生成したい画像を説明するプロンプトを入力
3. **画像生成**をクリック
4. 生成された画像をダウンロード

### マクロ実行モード
やりたいことを説明して**マクロ実行**をクリックすると、Excel JavaScript APIコードが生成・実行されます。

## 🤖 対応モデル

### OpenAI チャット
| モデル | 特徴 |
|-------|------|
| GPT-4o | 推奨、画像入力対応 |
| GPT-4o mini | 画像入力対応 |
| GPT-4.1 | 最新の非推論モデル |
| GPT-4 Turbo | レガシー |

### OpenAI 画像生成
| モデル | 特徴 |
|-------|------|
| GPT Image 1 | 最先端の画像生成 |

### Anthropic Claude
| モデル | 特徴 |
|-------|------|
| Claude Sonnet 4.5 | 推奨、画像+PDF入力対応 |
| Claude Haiku 4.5 | 高速、画像+PDF入力対応 |
| Claude Opus 4.5 | プレミアム、画像+PDF入力対応 |

### Google Gemini チャット
| モデル | 特徴 |
|-------|------|
| Gemini 2.0 Flash | 推奨、画像+PDF入力対応 |
| Gemini 2.5 Flash | 高速、画像+PDF入力対応 |
| Gemini 2.5 Pro | 高度な推論、画像+PDF入力対応 |

### Google Gemini 画像生成
| モデル | 特徴 |
|-------|------|
| Gemini 3 Pro Image | 推論機能付き画像生成 |

## 📎 ファイルサポート一覧

| プロバイダー | 画像入力 | PDF入力 | 画像生成 |
|------------|:-------:|:------:|:-------:|
| OpenAI GPT-4o/mini | ✅ | ❌ | ❌ |
| OpenAI GPT Image 1 | ❌ | ❌ | ✅ |
| Claude（全モデル） | ✅ | ✅ | ❌ |
| Gemini（全モデル） | ✅ | ✅ | ❌ |
| Gemini 3 Pro Image | ❌ | ❌ | ✅ |

## 🛠️ 開発

### 利用可能なスクリプト

| コマンド | 説明 |
|---------|------|
| `npm run dev-server` | webpack開発サーバーを起動 |
| `npm run build` | 本番用ビルド |
| `npm run build:dev` | 開発用ビルド |
| `npm run start` | Excel デスクトップでデバッグ開始 |
| `npm run start:web` | Excel Webでデバッグ開始 |
| `npm run lint` | ESLintを実行 |
| `npm run lint:fix` | ESLintエラーを修正 |

## 🤝 コントリビューション

コントリビューションを歓迎します！お気軽にPull Requestを送信してください。

1. リポジトリをフォーク
2. フィーチャーブランチを作成 (`git checkout -b feature/AmazingFeature`)
3. 変更をコミット (`git commit -m 'Add some AmazingFeature'`)
4. ブランチにプッシュ (`git push origin feature/AmazingFeature`)
5. Pull Requestを作成

## 📄 ライセンス

このプロジェクトはMITライセンスの下で公開されています。詳細は[LICENSE](LICENSE)ファイルをご覧ください。

---

<p align="center">
  Made with ❤️ for Excel users everywhere
</p>
