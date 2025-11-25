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

---

## ✨ Features

- 🤖 **AI Chat Assistant** - Chat with AI to get Excel advice, formula suggestions, and troubleshooting help
- 📊 **Context-Aware** - AI understands your currently selected cells and their values
- ⚡ **Code Execution** - Generate and execute Excel JavaScript API code directly
- 🔄 **Multiple LLM Support** - Works with OpenAI GPT and Anthropic Claude models
- 🌐 **Streaming Responses** - Real-time streaming for faster interaction
- 📝 **Markdown Rendering** - Beautiful formatted responses with code highlighting

## 🚀 Installation

### Prerequisites

- Node.js 16.x or later
- Microsoft Excel (Desktop or Web)
- API key from OpenAI or Anthropic

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

## 📖 Usage

1. Open Excel and go to the **Home** tab
2. Click the **OpenLLM** button to open the task pane
3. Enter your API key (OpenAI or Anthropic)
4. Select your preferred model
5. Start chatting with the AI assistant!

### Chat Mode
Type your question and click **送信 (Send)** to get AI advice about Excel.

### Macro Execution Mode
Describe what you want to do and click **マクロ実行 (Execute Macro)** to generate and run Excel JavaScript API code.

## 🤖 Supported Models

### OpenAI
- GPT-4o
- GPT-4o mini
- GPT-4 Turbo
- GPT-3.5 Turbo

### Anthropic Claude
- Claude 3.5 Sonnet
- Claude 3.5 Haiku
- Claude 3 Opus

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
│   │   └── taskpane.js      # Main logic
│   └── commands/
│       ├── commands.html
│       └── commands.js
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
- Powered by [OpenAI](https://openai.com/) and [Anthropic](https://anthropic.com/) APIs

---

<p align="center">
  Made with ❤️ for Excel users everywhere
</p>

