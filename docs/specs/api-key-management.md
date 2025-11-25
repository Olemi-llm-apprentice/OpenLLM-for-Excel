# APIキー管理機能 設計仕様書

## 1. 概要

### 1.1 要望

- 開発者は環境変数でAPIキーを設定したい（`.env.local`）
- 一般ユーザーは画面から入力して保存したい
- 毎回入力するのは面倒なので、永続化してほしい
- 複数デバイスで同じキーを使いたい

### 1.2 設計方針

**「3段階優先順位によるハイブリッド方式」** を採用

```
優先順位1: 環境変数（開発者向け）
    ↓ なければ
優先順位2: Office Settings API（ユーザー保存済み）
    ↓ なければ
優先順位3: UI入力フィールド（一時利用）
```

## 2. 要件定義

### 2.1 機能要件

| 要件ID | 要件 | 優先度 |
|--------|------|--------|
| REQ-001 | 環境変数からAPIキーを読み込める | 必須 |
| REQ-002 | UIからAPIキーを入力・保存できる | 必須 |
| REQ-003 | 保存済みキーを自動復元できる | 必須 |
| REQ-004 | プロバイダーごとに別々のキーを管理できる | 必須 |
| REQ-005 | APIキーの有効性をテストできる | 必須 |
| REQ-006 | 保存済みキーを削除できる | 必須 |
| REQ-007 | 複数デバイス間でキーが同期される | 推奨 |

### 2.2 非機能要件

| 要件ID | 要件 | 優先度 |
|--------|------|--------|
| NFR-001 | APIキーが平文でコミットされない | 必須 |
| NFR-002 | Office外でも最低限動作する | 推奨 |
| NFR-003 | 接続テストは5秒以内に完了する | 推奨 |

## 3. 設計詳細

### 3.1 モジュール構成

```
src/taskpane/config.js
├── 定数
│   └── ENV_KEYS          # 環境変数から注入されたキー
├── 内部関数
│   ├── normalizeProvider()   # プロバイダー名正規化
│   ├── getSavedApiKey()      # Office Settings から取得
│   ├── getInputApiKey()      # UI入力から取得
│   ├── testOpenAIKey()       # OpenAI キーテスト
│   ├── testClaudeKey()       # Claude キーテスト
│   └── testGeminiKey()       # Gemini キーテスト
└── エクスポート関数
    ├── getApiKey()           # 優先順位でキー取得
    ├── saveApiKey()          # キー保存
    ├── deleteApiKey()        # キー削除
    ├── testApiKey()          # キーテスト
    ├── hasSavedApiKey()      # 保存済みか確認
    ├── hasEnvApiKey()        # 環境変数あるか確認
    ├── restoreApiKeyToInput() # UIに復元
    └── getAllKeyStatus()     # 全状態取得
```

### 3.2 データ構造

#### ApiKeySource（列挙型）

```javascript
export const ApiKeySource = {
  ENV: 'env',      // 環境変数から取得
  SAVED: 'saved',  // Office Settings から取得
  INPUT: 'input',  // UI入力から取得
  NONE: 'none',    // 取得できず
};
```

#### getApiKey() の戻り値

```javascript
{
  key: string,          // APIキー本体
  source: ApiKeySource  // 取得元
}
```

### 3.3 プロバイダー名の正規化

**問題**: モデル選択の値が `openai-image:gpt-image-1` のように、画像生成用の別プロバイダー扱いになっている。

**解決**: `normalizeProvider()` で正規化

```javascript
function normalizeProvider(provider) {
  if (provider === 'openai-image') return 'openai';
  if (provider === 'gemini-image') return 'gemini';
  return provider;
}
```

**理由**: 画像生成でも同じAPIキーを使うため、内部的には同一プロバイダーとして扱う。

### 3.4 環境変数の注入方法

**実装**: Webpack の `DefinePlugin` を使用

```javascript
// webpack.config.js
new webpack.DefinePlugin({
  __OPENAI_API_KEY__: JSON.stringify(process.env.OPENAI_API_KEY || ""),
  __ANTHROPIC_API_KEY__: JSON.stringify(process.env.ANTHROPIC_API_KEY || ""),
  __GEMINI_API_KEY__: JSON.stringify(process.env.GEMINI_API_KEY || ""),
}),
```

**理由**:
- ビルド時に文字列として埋め込まれる
- 実行時に `process.env` を参照する必要がない（ブラウザで動作しない）
- `dotenv` で `.env.local` から読み込み可能

### 3.5 Office Settings API の選択理由

**比較検討**:

| 保存先 | メリット | デメリット | 判断 |
|--------|---------|-----------|------|
| localStorage | 実装簡単 | ブラウザ限定、同期なし | ❌ |
| IndexedDB | 大容量対応 | 過剰、同期なし | ❌ |
| Office Document Settings | ファイルに保存 | ファイル共有時に漏洩リスク | ❌ |
| Office Roaming Settings | MS アカウント同期 | Office 環境限定 | ✅ |

**結論**: Office Roaming Settings を採用し、Office 外では localStorage にフォールバック

### 3.6 APIキーテストの実装

各プロバイダーの**軽量なエンドポイント**を使用:

| プロバイダー | エンドポイント | 理由 |
|-------------|---------------|------|
| OpenAI | `GET /v1/models` | モデル一覧取得（課金なし） |
| Claude | `POST /v1/messages` (max_tokens=1) | 最小トークンで検証 |
| Gemini | `GET /v1beta/models` | モデル一覧取得（課金なし） |

**注意**: Claude は GET でテストできるエンドポイントがないため、最小限のメッセージを送信。

## 4. UI設計

### 4.1 コンポーネント構成

```html
<div id="api-key-container">
  <label>APIキー:</label>
  <input type="password" id="api-key-input" />
  <div id="api-key-buttons">
    <button id="save-api-key-button">💾 保存</button>
    <button id="test-api-key-button">🔌 テスト</button>
    <button id="delete-api-key-button">🗑️</button>
  </div>
  <div id="api-key-status"></div>
</div>
```

### 4.2 状態表示

| 状態 | 表示 | CSSクラス |
|------|------|----------|
| 情報 | `✓ 環境変数からAPIキーを読み込みました` | `.info` |
| 成功 | `✓ APIキーを保存しました` | `.success` |
| エラー | `✗ 接続失敗: Invalid API key` | `.error` |
| 処理中 | `🔄 接続テスト中...` | `.info` |

### 4.3 環境変数設定時の挙動

環境変数が設定されている場合:
1. 入力フィールドを `disabled` にする
2. プレースホルダーを「環境変数から設定済み」に変更
3. ステータスに「環境変数から読み込みました」を表示

**理由**: 環境変数は開発者が意図的に設定するものなので、UI操作で上書きされるべきではない。

## 5. テスト仕様

### 5.1 単体テスト（将来実装）

| テストID | テスト対象 | 入力 | 期待出力 |
|----------|-----------|------|---------|
| UT-001 | normalizeProvider | `'openai-image'` | `'openai'` |
| UT-002 | normalizeProvider | `'claude'` | `'claude'` |
| UT-003 | getApiKey (ENV設定あり) | provider='openai' | `{ key: 'sk-xxx', source: 'env' }` |
| UT-004 | getApiKey (全なし) | provider='openai' | `{ key: '', source: 'none' }` |

### 5.2 統合テスト（手動）

| テストID | シナリオ | 前提条件 | 手順 | 期待結果 |
|----------|---------|---------|------|---------|
| IT-001 | 環境変数読み込み | `.env.local`設定済み | サーバー再起動 | 「環境変数から読み込み」表示 |
| IT-002 | キー保存 | 有効なOpenAIキー | 入力→保存 | 「保存しました」表示 |
| IT-003 | キー復元 | IT-002実行後 | ページリロード | 入力欄にキー復元 |
| IT-004 | プロバイダー切替 | OpenAI/Claudeキー保存済み | モデル変更 | 対応キーに切替 |
| IT-005 | 接続テスト成功 | 有効なキー | テストボタン | 「接続成功」表示 |
| IT-006 | 接続テスト失敗 | 無効なキー | テストボタン | 「接続失敗」表示 |
| IT-007 | キー削除 | キー保存済み | 削除ボタン | 入力欄クリア |

### 5.3 合格基準

| 基準ID | 基準 | 検証方法 |
|--------|------|---------|
| AC-001 | 全統合テスト(IT-001〜IT-007)がパス | 手動実行 |
| AC-002 | 環境変数未設定時もエラーなく動作 | `.env.local`削除して確認 |
| AC-003 | Office外でもlocalStorageにフォールバック | ブラウザ単体で確認 |

## 6. セキュリティ考慮事項

### 6.1 脅威と対策

| 脅威 | リスク | 対策 |
|------|--------|------|
| 環境変数の誤コミット | 高 | `.env.local` を `.gitignore` に追加済み |
| DevToolsでキー確認 | 中 | 本番では回避困難（注意喚起のみ） |
| Office Settings漏洩 | 低 | MSアカウント保護に依存 |

### 6.2 ユーザーへの注意喚起

READMEに以下を明記:
- ブラウザのDevToolsでAPIキーが見える可能性がある
- 企業利用ではバックエンドプロキシを推奨
- APIキーを他人と共有しない

## 7. 今後の拡張

### 7.1 OAuth対応（将来）

```javascript
// 将来的な実装イメージ
export async function authenticateWithOAuth(provider) {
  // Office Dialog API でポップアップ
  const dialog = await Office.context.ui.displayDialogAsync(
    `https://auth.${provider}.com/oauth/authorize?...`
  );
  // トークン取得後、Office Settings に保存
}
```

### 7.2 キーのローテーション通知（将来）

```javascript
// キーの有効期限チェック（一部プロバイダー対応）
export async function checkKeyExpiration(provider, key) {
  // 実装予定
}
```

