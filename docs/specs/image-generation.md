# 画像生成機能 設計仕様書

## 1. 概要

### 1.1 要望

- Excelアドイン上で画像を生成したい
- OpenAIとGeminiの画像生成モデルを使いたい
- 生成した画像をダウンロードできるようにしてほしい

### 1.2 設計方針

**「チャット機能とは独立した画像生成フロー」** を採用

```
[プロンプト入力]
      │
      ▼
[画像生成ボタン]
      │
      ├── openai-image → generateImageOpenAI()
      └── gemini-image → generateImageGemini()
              │
              ▼
        [Base64画像]
              │
              ▼
        [Data URL表示 + ダウンロードリンク]
```

## 2. 対応モデル

### 2.1 OpenAI GPT Image 1

| 項目 | 値 |
|------|-----|
| モデルID | `gpt-image-1` |
| エンドポイント | `POST /v1/images/generations` |
| レスポンス形式 | Base64 or URL |
| 最大サイズ | 1024x1024（デフォルト） |

### 2.2 Gemini 3 Pro Image

| 項目 | 値 |
|------|-----|
| モデルID | `gemini-3-pro-image-preview` |
| エンドポイント | `POST /v1beta/models/{model}:generateContent` |
| レスポンス形式 | Base64（inlineData） |
| 特徴 | 推論機能付き |

## 3. 設計詳細

### 3.1 モデル選択での分離

画像生成モデルは `optgroup` で分離:

```html
<optgroup label="OpenAI 画像生成">
  <option value="openai-image:gpt-image-1">GPT Image 1（画像生成）</option>
</optgroup>
<optgroup label="Google Gemini 画像生成">
  <option value="gemini-image:gemini-3-pro-image-preview">Gemini 3 Pro Image</option>
</optgroup>
```

**プロバイダー値**: `openai-image`, `gemini-image`（チャット用と区別）

### 3.2 画像生成ハンドラ

```javascript
async function handleGenerateImage() {
  const { provider, model } = getProviderAndModel();
  const userInput = document.getElementById("message-input").value;
  
  // 画像生成モデルかチェック
  if (provider !== 'openai-image' && provider !== 'gemini-image') {
    showMessage("画像生成モデルを選択してください");
    return;
  }
  
  const { key: apiKey } = getApiKey(provider);
  
  let imageUrl;
  if (provider === 'openai-image') {
    imageUrl = await generateImageOpenAI(apiKey, model, userInput);
  } else if (provider === 'gemini-image') {
    imageUrl = await generateImageGemini(apiKey, model, userInput);
  }
  
  // 表示
  displayGeneratedImage(imageUrl);
}
```

### 3.3 OpenAI 画像生成

```javascript
async function generateImageOpenAI(apiKey, model, prompt) {
  const url = "https://api.openai.com/v1/images/generations";
  
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: model,
      prompt: prompt,
      n: 1,
      size: "1024x1024",
      response_format: "b64_json"  // Base64で取得
    }),
  });
  
  const data = await response.json();
  const base64 = data.data[0].b64_json;
  return `data:image/png;base64,${base64}`;
}
```

**なぜBase64か**:
- URL形式は一時的（1時間で失効）
- Base64ならローカルで永続的に保持可能
- ダウンロード機能が実装しやすい

### 3.4 Gemini 画像生成

```javascript
async function generateImageGemini(apiKey, model, prompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
  
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{
        parts: [{ text: prompt }]
      }],
      generationConfig: {
        responseModalities: ["image", "text"]
      }
    }),
  });
  
  const data = await response.json();
  const imagePart = data.candidates[0].content.parts.find(p => p.inlineData);
  
  if (imagePart) {
    const { mimeType, data: base64 } = imagePart.inlineData;
    return `data:${mimeType};base64,${base64}`;
  }
  
  throw new Error("画像が生成されませんでした");
}
```

### 3.5 画像表示

```javascript
function displayGeneratedImage(imageUrl) {
  aiMessageElement.innerHTML = `
    AI: 画像を生成しました。<br/>
    <img src="${imageUrl}" class="generated-image" alt="Generated image" /><br/>
    <a href="${imageUrl}" download="generated-image.png" class="image-download-link">
      📥 画像をダウンロード
    </a>
  `;
}
```

## 4. UI設計

### 4.1 ボタン配置

```
[メッセージ入力欄]
[送信] [画像生成] [マクロ実行]
```

### 4.2 スタイル

```css
#generate-image-button {
  background-color: #6f42c1;  /* 紫色で区別 */
  color: white;
  border: none;
  padding: 5px 10px;
  cursor: pointer;
  border-radius: 4px;
}

.generated-image {
  max-width: 100%;
  border-radius: 8px;
  margin-top: 10px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.15);
}
```

## 5. テスト仕様

### 5.1 テストケース

| テストID | シナリオ | 前提条件 | 手順 | 期待結果 |
|----------|---------|---------|------|---------|
| IMG-001 | OpenAI画像生成 | 有効なOpenAIキー | GPT Image 1選択→プロンプト→生成 | 画像が表示される |
| IMG-002 | Gemini画像生成 | 有効なGeminiキー | Gemini 3 Pro Image選択→プロンプト→生成 | 画像が表示される |
| IMG-003 | ダウンロード | IMG-001/002実行後 | ダウンロードリンククリック | 画像がダウンロードされる |
| IMG-004 | 非画像モデルでの生成試行 | - | GPT-4o選択→画像生成ボタン | エラーメッセージ表示 |
| IMG-005 | 空プロンプト | - | プロンプト未入力→生成 | エラーメッセージ表示 |

### 5.2 合格基準

| 基準ID | 基準 | 検証方法 |
|--------|------|---------|
| AC-001 | 両プロバイダーで画像生成動作 | 手動実行 |
| AC-002 | 生成画像がブラウザに表示される | 目視確認 |
| AC-003 | ダウンロードリンクが機能する | 実際にダウンロード |

## 6. 制限事項

| 制限 | 理由 | 回避策 |
|------|------|--------|
| 画像編集（Inpainting）未対応 | 複雑なUI必要 | 将来実装検討 |
| 複数画像生成未対応 | シンプルさ優先 | n=1固定 |
| サイズ選択未対応 | シンプルさ優先 | 1024x1024固定 |

## 7. 今後の拡張

### 7.1 検討中の機能

- 画像編集（Inpainting）
- 生成サイズの選択
- スタイル指定
- 生成履歴の保存

