# Anthropic Claude API モデル情報

> 最終更新: 2025年11月26日（公式ドキュメントより）

## 概要

Anthropic は Claude シリーズの AI モデルを開発しています。安全性と有用性に重点を置いたモデルで、長いコンテキスト長が特徴です。

## API エンドポイント

- **ベース URL**: `https://api.anthropic.com/v1`
- **Messages API**: `POST /messages`

## 推奨モデル（2025年11月時点）

### Claude Sonnet 4.5（推奨・最新）

| 項目 | 内容 |
|------|------|
| モデルID | `claude-sonnet-4-5-20250929` |
| エイリアス | `claude-sonnet-4-5` |
| 説明 | 複雑なエージェントとコーディング向けの最もスマートなモデル |
| コンテキスト長 | 200K トークン（1M トークン beta対応） |
| 最大出力 | 64K トークン |
| 入力料金 | $3 / 1M tokens |
| 出力料金 | $15 / 1M tokens |
| 特徴 | 高品質・高速・エージェント特化 |

### Claude Haiku 4.5（最速）

| 項目 | 内容 |
|------|------|
| モデルID | `claude-haiku-4-5-20251001` |
| エイリアス | `claude-haiku-4-5` |
| 説明 | 近フロンティア知性を持つ最速モデル |
| コンテキスト長 | 200K トークン |
| 最大出力 | 64K トークン |
| 入力料金 | $1 / 1M tokens |
| 出力料金 | $5 / 1M tokens |
| 特徴 | 最速・低コスト |

### Claude Opus 4.5（プレミアム）

| 項目 | 内容 |
|------|------|
| モデルID | `claude-opus-4-5-20251101` |
| エイリアス | `claude-opus-4-5` |
| 説明 | 最大の知性と実用的なパフォーマンスを組み合わせたプレミアムモデル |
| コンテキスト長 | 200K トークン |
| 最大出力 | 64K トークン |
| 入力料金 | $5 / 1M tokens |
| 出力料金 | $25 / 1M tokens |
| 特徴 | 最高品質 |

### Claude Opus 4.1（専門推論）

| 項目 | 内容 |
|------|------|
| モデルID | `claude-opus-4-1-20250805` |
| エイリアス | `claude-opus-4-1` |
| 説明 | 専門的な推論タスク向けの優れたモデル |
| コンテキスト長 | 200K トークン |
| 最大出力 | 32K トークン |
| 入力料金 | $15 / 1M tokens |
| 出力料金 | $75 / 1M tokens |
| 特徴 | 複雑な推論に特化 |

## API リクエスト例

```javascript
const response = await fetch('https://api.anthropic.com/v1/messages', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'x-api-key': API_KEY,
    'anthropic-version': '2023-06-01'
  },
  body: JSON.stringify({
    model: 'claude-sonnet-4-5-20250929',
    max_tokens: 4096,
    messages: [
      { role: 'user', content: 'Hello!' }
    ],
    stream: true  // ストリーミング有効
  })
});
```

## 重要な相違点（OpenAI との違い）

1. **認証ヘッダー**: `x-api-key` を使用（`Authorization: Bearer` ではない）
2. **API バージョン**: `anthropic-version` ヘッダーが必須
3. **max_tokens**: 必須パラメータ
4. **system メッセージ**: messages 配列ではなく、別の `system` パラメータで指定

```javascript
{
  model: 'claude-sonnet-4-5-20250929',
  max_tokens: 4096,
  system: 'You are a helpful assistant.',  // ここが違う
  messages: [
    { role: 'user', content: 'Hello!' }
  ]
}
```

## ストリーミングレスポンス形式

```
event: message_start
data: {"type":"message_start","message":{"id":"msg_xxx","type":"message","role":"assistant"}}

event: content_block_delta
data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"Hello"}}

event: message_stop
data: {"type":"message_stop"}
```

## 公式ドキュメント

- [Anthropic API リファレンス](https://docs.anthropic.com/en/api/getting-started)
- [モデル一覧](https://platform.claude.com/docs/en/about-claude/models/overview)
- [料金](https://www.anthropic.com/pricing)

