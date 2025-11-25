# OpenLLM for Excel 設計仕様書

このディレクトリには、OpenLLM for Excel の設計思想、機能仕様、および技術的な意思決定の根拠をまとめています。

## 📚 ドキュメント一覧

| ドキュメント | 概要 |
|-------------|------|
| [architecture.md](./architecture.md) | 全体アーキテクチャと設計思想 |
| [api-key-management.md](./api-key-management.md) | APIキー管理機能の設計仕様 |
| [multi-provider-llm.md](./multi-provider-llm.md) | マルチプロバイダーLLM統合の設計仕様 |
| [image-generation.md](./image-generation.md) | 画像生成機能の設計仕様 |
| [file-upload.md](./file-upload.md) | ファイルアップロード（Vision）機能の設計仕様 |
| [excel-integration.md](./excel-integration.md) | Excel統合機能の設計仕様 |

## 🎯 プロジェクトの目標

### ビジョン

**「誰でも無料で使える、オープンソースのExcel AIアシスタント」**

市販のExcel AIツールは多くが有料であり、APIキーを持っていても気軽に試せない状況があります。このプロジェクトは、ユーザーが自身のAPIキーを使って、無料でAI機能をExcelに統合できることを目指しています。

### 設計原則

1. **シンプルさ優先**
   - サーバーレス（クライアント完結）
   - 最小限の依存関係
   - 直感的なUI

2. **オープン性**
   - オープンソース（MIT License）
   - 特定ベンダーにロックインしない
   - 複数のLLMプロバイダーをサポート

3. **セキュリティとプライバシー**
   - APIキーはユーザー管理
   - データは第三者サーバーを経由しない
   - 透明性のあるコード

4. **拡張性**
   - 新しいLLMプロバイダーを追加しやすい設計
   - 機能ごとにモジュール化

## 🏗️ 技術スタック

| カテゴリ | 技術 | 選定理由 |
|---------|------|---------|
| フレームワーク | Office Add-in (JavaScript) | Excel統合の公式方法 |
| ビルドツール | Webpack | Office Add-inテンプレート標準 |
| LLM API | OpenAI, Anthropic, Google | 主要3社をカバー |
| 状態管理 | Office Settings API | クラウド同期、セキュア |
| スタイリング | Pure CSS | 最小限の依存関係 |

## 📊 機能マトリックス

| 機能 | OpenAI | Claude | Gemini |
|------|:------:|:------:|:------:|
| チャット | ✅ | ✅ | ✅ |
| ストリーミング | ✅ | ✅ | ✅ |
| 画像入力（Vision） | ✅ | ✅ | ✅ |
| PDF入力 | ❌ | ✅ | ✅ |
| 画像生成 | ✅ | ❌ | ✅ |
| コード生成・実行 | ✅ | ✅ | ✅ |

## 📝 ドキュメント更新履歴

| 日付 | 内容 |
|------|------|
| 2025-11-26 | 初版作成 |

## 🔗 関連リソース

- [README.md](../../README.md) - プロジェクト概要
- [AGENT_WORK_LOG.md](../AGENT_WORK_LOG.md) - 作業履歴
- [openai-models.md](../openai-models.md) - OpenAI モデル情報
- [anthropic-claude-models.md](../anthropic-claude-models.md) - Claude モデル情報
- [google-gemini-models.md](../google-gemini-models.md) - Gemini モデル情報

