# マルチプロバイダーLLM統合 設計仕様書

## 1. 概要

### 1.1 要望

- 特定のLLMプロバイダーに縛られたくない
- OpenAI、Claude、Gemini を同じUIで切り替えて使いたい
- 各プロバイダーの最新モデルに対応してほしい
- ストリーミングレスポンスでリアルタイムに表示してほしい

### 1.2 設計方針

**「プロバイダー抽象化層」** を設けず、**直接分岐方式** を採用

```
┌─────────────────┐
│  sendToLLM()    │  ← 統合関数
└─────────────────┘
        │
   switch(provider)
        │
  ┌─────┼─────┐
  ▼     ▼     ▼
OpenAI Claude Gemini
```

**理由**:
- 各プロバイダーのAPIは微妙に異なる（認証、リクエスト形式、レスポンス形式）
- 過度な抽象化は複雑さを増す
- プロバイダーは3つしかないので、直接分岐で十分

## 2. プロバイダー比較

### 2.1 API仕様の違い

| 項目 | OpenAI | Claude | Gemini |
|------|--------|--------|--------|
| 認証方式 | `Authorization: Bearer` | `x-api-key` | URL パラメータ `?key=` |
| エンドポイント | `/v1/chat/completions` | `/v1/messages` | `/v1beta/models/{model}:streamGenerateContent` |
| メッセージ形式 | `messages: [{role, content}]` | `messages: [{role, content}]` | `contents: [{role, parts}]` |
| システムプロンプト | `messages[0]` に含める | `system` パラメータ | `systemInstruction` |
| ストリーミング | `stream: true` | `stream: true` | URL に `?alt=sse` |

### 2.2 ストリーミングレスポンスの違い

#### OpenAI

```
data: {"choices":[{"delta":{"content":"Hello"}}]}
data: {"choices":[{"delta":{"content":" world"}}]}
data: [DONE]
```

#### Claude

```
event: content_block_delta
data: {"type":"content_block_delta","delta":{"type":"text_delta","text":"Hello"}}

event: content_block_delta
data: {"type":"content_block_delta","delta":{"type":"text_delta","text":" world"}}

event: message_stop
```

#### Gemini

```
data: {"candidates":[{"content":{"parts":[{"text":"Hello"}]}}]}
data: {"candidates":[{"content":{"parts":[{"text":" world"}]}}]}
```

## 3. 設計詳細

### 3.1 モデル選択の設計

#### HTML構造

```html
<select id="model-select">
  <optgroup label="OpenAI チャット">
    <option value="openai:gpt-4o">GPT-4o（推奨）</option>
    <option value="openai:gpt-4o-mini">GPT-4o mini</option>
  </optgroup>
  <optgroup label="OpenAI 画像生成">
    <option value="openai-image:gpt-image-1">GPT Image 1</option>
  </optgroup>
  <optgroup label="Anthropic Claude">
    <option value="claude:claude-sonnet-4-5-20250929">Claude Sonnet 4.5</option>
  </optgroup>
  <!-- ... -->
</select>
```

#### value の設計

```
{provider}:{model}

例:
- openai:gpt-4o
- openai-image:gpt-image-1
- claude:claude-sonnet-4-5-20250929
- gemini:gemini-2.0-flash
```

**理由**:
- `optgroup` だけでは JavaScript からプロバイダーを特定できない
- value に明示的にプロバイダーを含めることで、パースが容易

#### パース関数

```javascript
function getProviderAndModel() {
  const selectEl = document.getElementById("model-select");
  const [provider, model] = selectEl.value.split(":");
  return { provider, model };
}
```

### 3.2 統合送信関数

```javascript
async function sendToLLM(userInput, excel_prompt) {
  const { provider, model } = getProviderAndModel();
  const { key: apiKey } = getApiKey(provider);
  
  if (!apiKey) {
    throw new Error("APIキーが設定されていません");
  }
  
  const system_prompt = buildSystemPrompt(excel_prompt);
  
  switch (provider) {
    case 'openai':
      return await sendToOpenAI(apiKey, model, system_prompt, userInput, ...);
    case 'claude':
      return await sendToClaude(apiKey, model, system_prompt, userInput, ...);
    case 'gemini':
      return await sendToGemini(apiKey, model, system_prompt, userInput, ...);
    default:
      throw new Error(`Unknown provider: ${provider}`);
  }
}
```

### 3.3 各プロバイダーの送信関数

#### OpenAI

```javascript
async function sendToOpenAI(apiKey, model, system_prompt, userInput, aiMessageElement, files = []) {
  const url = "https://api.openai.com/v1/chat/completions";
  
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: model,
      messages: [
        { role: "system", content: system_prompt },
        ...conversationHistory,
        { role: "user", content: buildOpenAIContent(userInput, files) }
      ],
      stream: true,
    }),
  });
  
  // ストリーミング処理...
}
```

#### Claude

```javascript
async function sendToClaude(apiKey, model, system_prompt, userInput, aiMessageElement, files = []) {
  const url = "https://api.anthropic.com/v1/messages";
  
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true'  // 重要
    },
    body: JSON.stringify({
      model: model,
      system: system_prompt,
      messages: buildClaudeMessages(userInput, files),
      max_tokens: 4096,
      stream: true,
    }),
  });
  
  // ストリーミング処理...
}
```

**注意**: `anthropic-dangerous-direct-browser-access` ヘッダーは、ブラウザから直接Claudeを呼ぶために必要。Anthropicが「危険」と命名しているが、個人利用では許容。

#### Gemini

```javascript
async function sendToGemini(apiKey, model, system_prompt, userInput, aiMessageElement, files = []) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:streamGenerateContent?alt=sse&key=${apiKey}`;
  
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      systemInstruction: { parts: [{ text: system_prompt }] },
      contents: buildGeminiContents(userInput, files),
    }),
  });
  
  // ストリーミング処理...
}
```

### 3.4 ストリーミング処理

共通パターン:

```javascript
const reader = response.body.getReader();
const decoder = new TextDecoder();
let streamedResponse = "";

while (true) {
  const { done, value } = await reader.read();
  if (done) break;
  
  const chunk = decoder.decode(value, { stream: true });
  // プロバイダー別にパース
  const text = parseChunk(chunk);
  streamedResponse += text;
  
  // UI更新
  aiMessageElement.innerHTML = marked.parse(streamedResponse);
}
```

## 4. 会話履歴の管理

### 4.1 データ構造

```javascript
let conversationHistory = [
  { role: "user", content: "最初の質問" },
  { role: "assistant", content: "最初の回答" },
  { role: "user", content: "フォローアップ" },
  // ...
];
```

### 4.2 プロバイダー間の互換性

| プロバイダー | user role | assistant role |
|-------------|-----------|----------------|
| OpenAI | `user` | `assistant` |
| Claude | `user` | `assistant` |
| Gemini | `user` | `model` |

**対策**: Gemini送信時に `assistant` → `model` に変換

```javascript
const geminiContents = conversationHistory.map(msg => ({
  role: msg.role === 'assistant' ? 'model' : msg.role,
  parts: [{ text: msg.content }]
}));
```

## 5. エラーハンドリング

### 5.1 エラー種別

| エラー | 原因 | 対応 |
|--------|------|------|
| 401 Unauthorized | APIキー無効 | 「APIキーを確認してください」 |
| 429 Rate Limited | レート制限 | 「しばらく待ってから再試行」 |
| 500 Server Error | サーバー障害 | 「サービスが一時的に利用不可」 |
| Network Error | 通信障害 | 「ネットワーク接続を確認」 |

### 5.2 共通エラーハンドリング

```javascript
try {
  const response = await sendToLLM(...);
} catch (error) {
  console.error("LLM Error:", error);
  aiMessageElement.innerHTML = `エラー: ${error.message}`;
}
```

## 6. 対応モデル一覧

### 6.1 OpenAI

| モデルID | 用途 | 特徴 |
|---------|------|------|
| `gpt-4o` | チャット | 最新フラグシップ、Vision対応 |
| `gpt-4o-mini` | チャット | 軽量版、Vision対応 |
| `gpt-4.1` | チャット | 非推論モデル |
| `gpt-4-turbo` | チャット | レガシー |
| `gpt-image-1` | 画像生成 | DALL-E後継 |

### 6.2 Anthropic Claude

| モデルID | 用途 | 特徴 |
|---------|------|------|
| `claude-sonnet-4-5-20250929` | チャット | バランス型、推奨 |
| `claude-haiku-4-5-20251001` | チャット | 高速 |
| `claude-opus-4-5-20251101` | チャット | 最高性能 |

### 6.3 Google Gemini

| モデルID | 用途 | 特徴 |
|---------|------|------|
| `gemini-2.0-flash` | チャット | 推奨、高速 |
| `gemini-2.5-flash` | チャット | 最新Flash |
| `gemini-2.5-pro` | チャット | 高度な推論 |
| `gemini-3-pro-image-preview` | 画像生成 | 推論付き画像生成 |

## 7. テスト仕様

### 7.1 統合テスト

| テストID | シナリオ | 前提条件 | 手順 | 期待結果 |
|----------|---------|---------|------|---------|
| LLM-001 | OpenAIチャット | 有効なOpenAIキー | GPT-4o選択→送信 | ストリーミング返答 |
| LLM-002 | Claudeチャット | 有効なClaudeキー | Sonnet選択→送信 | ストリーミング返答 |
| LLM-003 | Geminiチャット | 有効なGeminiキー | Flash選択→送信 | ストリーミング返答 |
| LLM-004 | プロバイダー切替 | 全キー設定済み | OpenAI→Claude切替→送信 | 正常動作 |
| LLM-005 | 会話履歴保持 | - | 2往復の会話 | 文脈を理解した回答 |
| LLM-006 | 無効キーエラー | 無効なキー | 送信 | エラーメッセージ表示 |

### 7.2 合格基準

| 基準ID | 基準 | 検証方法 |
|--------|------|---------|
| AC-001 | 3プロバイダー全てでチャット動作 | 手動実行 |
| AC-002 | ストリーミングがリアルタイム表示 | 目視確認 |
| AC-003 | Markdownが正しくレンダリング | コードブロック含む返答で確認 |
| AC-004 | 会話履歴が保持される | 2往復以上の会話 |

## 8. 今後の拡張

### 8.1 新プロバイダー追加手順

1. `taskpane.html` に `<optgroup>` 追加
2. `config.js` にテスト関数追加
3. `taskpane.js` に送信関数追加
4. ドキュメント更新

### 8.2 検討中の機能

- モデルの自動選択（タスクに応じて最適モデルを選択）
- コスト計算表示
- レスポンス品質のフィードバック機能

