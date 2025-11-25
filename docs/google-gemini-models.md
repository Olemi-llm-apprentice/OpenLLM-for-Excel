# Google Gemini API モデル情報

> 最終更新: 2025年11月26日（公式ドキュメントより）

## 概要

Google の Gemini は、DeepMind によって開発されたマルチモーダル AI モデルファミリーです。テキスト、画像、動画、音声などの複数のモダリティを処理できます。

## API エンドポイント

- **ベース URL**: `https://generativelanguage.googleapis.com/v1beta`
- **Generate Content**: `POST /models/{model}:generateContent`
- **Stream Generate Content**: `POST /models/{model}:streamGenerateContent`

## 推奨モデル（2025年11月時点）

### Gemini 3シリーズ（最新・推奨）

| モデルID | 説明 | 特徴 |
|----------|------|------|
| `gemini-3-pro-preview` | Gemini 3 Pro プレビュー版 | 最新・最高性能、画像/PDF入力対応 |

> **注意**: `gemini-3-pro` （正式版）はまだAPIで利用できません。`gemini-3-pro-preview` を使用してください。

#### Gemini 3 Pro Preview 詳細

| 項目 | 内容 |
|------|------|
| コンテキスト長 | 100万トークン入力 / 64Kトークン出力 |
| 知識カットオフ | 2025年1月 |
| 入力料金 | $2 / 1M tokens（<200K）, $4 / 1M tokens（>200K） |
| 出力料金 | $12 / 1M tokens（<200K）, $18 / 1M tokens（>200K） |
| 特徴 | 最先端の推論、エージェント・コーディング特化 |

### Gemini 3 Pro Image

| 項目 | 内容 |
|------|------|
| モデルID | `gemini-3-pro-image-preview` |
| 説明 | 画像生成・編集機能を持つ Gemini 3 Pro |
| コンテキスト長 | 65Kトークン入力 / 32Kトークン出力 |
| 入力料金 | $2 / 1M tokens |
| 出力料金 | $0.134 / 画像 |
| 特徴 | 4K画像生成、会話型編集 |

### Gemini 2.5 Pro（高度な思考）

| 項目 | 内容 |
|------|------|
| モデルID | `gemini-2.5-pro` |
| 説明 | コード、数学、STEM の複雑な問題を推論できる最先端の思考モデル |
| 特徴 | 長いコンテキストで大規模データセット・コードベース分析に最適 |

### Gemini 2.5 Flash（バランス）

| 項目 | 内容 |
|------|------|
| モデルID | `gemini-2.5-flash` |
| 説明 | 価格とパフォーマンスの面で最適なバランスモデル |
| 特徴 | 大規模処理、低レイテンシ、エージェント向け |

### Gemini 2.5 Flash-Lite（最速）

| 項目 | 内容 |
|------|------|
| モデルID | `gemini-2.5-flash-lite` |
| 説明 | 費用対効果と高スループットを重視した最も高速な Flash モデル |
| 特徴 | 最速、最低コスト |

### Gemini 2.0 Flash（レガシー）

| 項目 | 内容 |
|------|------|
| モデルID | `gemini-2.0-flash` |
| 説明 | 第2世代の主力モデル。100万トークンのコンテキストウィンドウ |
| 特徴 | 安定した性能、広く使用されている |

## API リクエスト例

```javascript
const response = await fetch(
  `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${API_KEY}`,
  {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      contents: [
        {
          parts: [
            { text: 'Hello!' }
          ]
        }
      ]
    })
  }
);
```

## ストリーミングリクエスト例

```javascript
const response = await fetch(
  `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:streamGenerateContent?alt=sse&key=${API_KEY}`,
  {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      contents: [
        {
          parts: [
            { text: 'Hello!' }
          ]
        }
      ]
    })
  }
);
```

## 重要な相違点（OpenAI/Claude との違い）

1. **認証**: URL パラメータで API キーを渡す（`?key=API_KEY`）
2. **メッセージ形式**: `contents` 配列に `parts` を含む形式
3. **ロール**: `user` と `model`（OpenAI の `assistant` に相当）
4. **system メッセージ**: `systemInstruction` パラメータで指定

```javascript
{
  systemInstruction: {
    parts: [{ text: 'You are a helpful assistant.' }]
  },
  contents: [
    {
      role: 'user',
      parts: [{ text: 'Hello!' }]
    }
  ]
}
```

## ストリーミングレスポンス形式

```
data: {"candidates":[{"content":{"parts":[{"text":"Hello"}],"role":"model"}}]}

data: {"candidates":[{"content":{"parts":[{"text":"!"}],"role":"model"}}]}

data: {"candidates":[{"content":{"parts":[{"text":""}],"role":"model"},"finishReason":"STOP"}]}
```

## JSON モード

```javascript
{
  contents: [...],
  generationConfig: {
    responseMimeType: 'application/json'
  }
}
```

## 公式ドキュメント

- [Gemini API ドキュメント](https://ai.google.dev/gemini-api/docs)
- [モデル一覧](https://ai.google.dev/gemini-api/docs/models)
- [料金](https://ai.google.dev/gemini-api/docs/pricing)
- [Gemini 3 ガイド](https://ai.google.dev/gemini-api/docs/gemini-3)

