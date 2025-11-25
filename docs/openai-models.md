# OpenAI API モデル情報

> 最終更新: 2025年11月26日（公式ドキュメントより）

## 概要

OpenAI は、ChatGPT を含む大規模言語モデルの先駆者です。API を通じて様々なモデルを提供しています。

## API エンドポイント

- **ベース URL**: `https://api.openai.com/v1`
- **Chat Completions**: `POST /chat/completions`

## 推奨モデル（2025年11月時点）

### GPT-5.1（推奨・最新）

| 項目 | 内容 |
|------|------|
| モデルID | `gpt-5.1` |
| 説明 | コーディングとエージェントタスク向けの最新フラッグシップモデル。推論深度を調整可能 |
| 特徴 | 最高性能、エージェント・コーディング特化 |

### GPT-5 mini（コスト効率）

| 項目 | 内容 |
|------|------|
| モデルID | `gpt-5-mini` |
| 説明 | GPT-5 の高速・コスト効率版。明確に定義されたタスク向け |
| 特徴 | 高速、コスト効率、日常的なタスクに最適 |

### GPT-5 nano（最速・最低コスト）

| 項目 | 内容 |
|------|------|
| モデルID | `gpt-5-nano` |
| 説明 | GPT-5 シリーズで最も高速かつコスト効率の良いモデル |
| 特徴 | 最速、最低コスト |

### GPT-4.1（非推論モデル）

| 項目 | 内容 |
|------|------|
| モデルID | `gpt-4.1` |
| 説明 | 最もスマートな非推論モデル |
| 特徴 | 推論機能を必要としないタスクに最適 |

### GPT-4o（レガシー）

| 項目 | 内容 |
|------|------|
| モデルID | `gpt-4o` |
| 説明 | 高速・インテリジェント・フレキシブルなGPTモデル |
| 特徴 | 依然として高性能、マルチモーダル対応 |

### GPT-4o mini（レガシー・低コスト）

| 項目 | 内容 |
|------|------|
| モデルID | `gpt-4o-mini` |
| 説明 | 高速・低コストの小型モデル |
| 特徴 | フォーカスされたタスクに最適 |

## API リクエスト例

```javascript
const response = await fetch('https://api.openai.com/v1/chat/completions', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Authorization': `Bearer ${API_KEY}`
  },
  body: JSON.stringify({
    model: 'gpt-4o',
    messages: [
      { role: 'system', content: 'You are a helpful assistant.' },
      { role: 'user', content: 'Hello!' }
    ],
    stream: true  // ストリーミング有効
  })
});
```

## ストリーミングレスポンス形式

```
data: {"id":"chatcmpl-xxx","object":"chat.completion.chunk","choices":[{"delta":{"content":"Hello"},"index":0}]}
data: {"id":"chatcmpl-xxx","object":"chat.completion.chunk","choices":[{"delta":{"content":"!"},"index":0}]}
data: [DONE]
```

## JSON モード

```javascript
{
  model: 'gpt-4o',
  messages: [...],
  response_format: { type: 'json_object' }
}
```

## 公式ドキュメント

- [OpenAI API リファレンス](https://platform.openai.com/docs/api-reference)
- [モデル一覧](https://platform.openai.com/docs/models)
- [料金](https://openai.com/pricing)

