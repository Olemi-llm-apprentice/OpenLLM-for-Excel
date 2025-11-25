# ライセンスレポート

**プロジェクト名**: OpenLLM for Excel  
**レポート作成日**: 2025年11月25日  
**プロジェクトライセンス**: MIT License

---

## 概要

本レポートは、OpenLLM for Excelプロジェクトで使用されているすべてのnpmパッケージのライセンスを調査した結果をまとめたものです。

---

## ライセンス分布サマリー

| ライセンス種別 | パッケージ数 | 比率 |
|--------------|------------|------|
| MIT | 983 | 84.0% |
| ISC | 57 | 4.9% |
| Apache-2.0 | 41 | 3.5% |
| BSD-2-Clause | 35 | 3.0% |
| BSD-3-Clause | 33 | 2.8% |
| BlueOak-1.0.0 | 7 | 0.6% |
| MIT-0 | 1 | 0.1% |
| Python-2.0 | 1 | 0.1% |
| CC-BY-4.0 | 1 | 0.1% |
| (MIT OR WTFPL) | 1 | 0.1% |
| (BSD-3-Clause OR GPL-2.0) | 1 | 0.1% |
| Unlicense | 1 | 0.1% |
| (BSD-2-Clause OR MIT OR Apache-2.0) | 1 | 0.1% |
| 0BSD | 1 | 0.1% |
| (MIT OR CC0-1.0) | 1 | 0.1% |
| **合計** | **約1,165** | **100%** |

---

## 本番依存関係のライセンス（dependencies）

package.jsonで `dependencies` として定義されているパッケージは以下の1つのみです：

| パッケージ名 | バージョン | ライセンス |
|------------|-----------|----------|
| core-js | ^3.47.0 | **MIT** |

✅ 本番依存関係は**MITライセンス**のみで構成されており、商用利用に問題ありません。

---

## 開発依存関係のライセンス（devDependencies）

package.jsonで `devDependencies` として定義されている主要パッケージ：

| パッケージ名 | バージョン | ライセンス | 発行元 |
|------------|-----------|----------|-------|
| @babel/core | ^7.28.5 | MIT | - |
| @babel/preset-env | ^7.28.5 | MIT | - |
| @types/jest | ^30.0.0 | MIT | DefinitelyTyped |
| @types/office-js | ^1.0.560 | MIT | DefinitelyTyped |
| @types/office-runtime | ^1.0.35 | MIT | DefinitelyTyped |
| acorn | ^8.15.0 | MIT | - |
| babel-jest | ^30.2.0 | MIT | Facebook/Meta |
| babel-loader | ^10.0.0 | MIT | - |
| copy-webpack-plugin | ^13.0.1 | MIT | webpack-contrib |
| dotenv | ^16.4.5 | BSD-2-Clause | - |
| eslint | ^9.39.1 | MIT | OpenJS Foundation |
| eslint-plugin-office-addins | ^4.0.6 | MIT | Microsoft |
| html-loader | ^5.1.0 | MIT | webpack-contrib |
| html-webpack-plugin | ^5.6.5 | MIT | webpack-contrib |
| jest | ^30.2.0 | MIT | Facebook/Meta |
| jest-environment-jsdom | ^30.2.0 | MIT | Facebook/Meta |
| office-addin-cli | ^2.0.6 | MIT | Microsoft (Office Dev) |
| office-addin-debugging | ^6.0.6 | MIT | Microsoft (Office Dev) |
| office-addin-dev-certs | ^2.0.6 | MIT | Microsoft (Office Dev) |
| office-addin-lint | ^3.0.6 | MIT | Microsoft (Office Dev) |
| office-addin-manifest | ^2.1.2 | MIT | Microsoft (Office Dev) |
| office-addin-prettier-config | ^2.0.1 | MIT | Microsoft (Office Dev) |
| regenerator-runtime | ^0.14.1 | MIT | Facebook |
| source-map-loader | ^5.0.0 | MIT | webpack-contrib |
| webpack | ^5.103.0 | MIT | Tobias Koppers @sokra |
| webpack-cli | ^6.0.1 | MIT | webpack |
| webpack-dev-server | ^5.2.2 | MIT | webpack |

✅ 開発依存関係もすべて**OSS互換ライセンス**で構成されています。

---

## ライセンス種別の説明

### ✅ 商用利用に問題のないライセンス

| ライセンス | 説明 | 主な条件 |
|-----------|------|---------|
| **MIT** | 最も寛容なライセンス | 著作権表示の保持のみ |
| **ISC** | MITとほぼ同等 | 著作権表示の保持のみ |
| **BSD-2-Clause** | 2条項BSDライセンス | 著作権表示の保持 |
| **BSD-3-Clause** | 3条項BSDライセンス | 著作権表示の保持、発行元名の使用制限 |
| **Apache-2.0** | Apache License 2.0 | 著作権表示、変更の記載、特許ライセンス付与 |
| **BlueOak-1.0.0** | Blue Oak Model License | 非常に寛容、著作権表示のみ |
| **0BSD** | Zero-Clause BSD | 制限なし（パブリックドメイン相当） |
| **MIT-0** | MIT No Attribution | 帰属表示不要 |
| **Unlicense** | パブリックドメイン | 制限なし |
| **CC-BY-4.0** | Creative Commons Attribution | 帰属表示のみ |

### ⚠️ 注意が必要なライセンス

| ライセンス | 説明 | 対象パッケージ数 | 備考 |
|-----------|------|----------------|------|
| **(BSD-3-Clause OR GPL-2.0)** | デュアルライセンス | 1 | BSD-3-Clauseを選択可能 |
| **Python-2.0** | Python Software Foundation License | 1 | GPL互換、商用利用可能 |

---

## コンプライアンス要件

### 必須対応事項

1. **著作権表示の保持**
   - MIT、BSD、ISCライセンスのパッケージについては、配布時にLICENSEファイルまたは著作権表示を含める必要があります。
   - 本プロジェクトでは `node_modules` は配布に含まれないため、ビルド成果物に対する対応は通常不要です。

2. **ライセンステキストの同梱**
   - Apache-2.0ライセンスのパッケージを使用する場合、LICENSEファイルの同梱が推奨されます。

### 推奨対応事項

1. **THIRD-PARTY-LICENSES.txt の作成**
   - 配布物に含まれるサードパーティライセンスの一覧を作成することを推奨します。
   
2. **ライセンス監視の継続**
   - 依存関係の更新時にライセンス変更がないか定期的に確認してください。
   - `npx license-checker --summary` で確認できます。

---

## リスク評価

| 項目 | 評価 | 説明 |
|------|------|------|
| **GPL系ライセンス** | ✅ 低リスク | GPL系ライセンス（強いコピーレフト）は含まれていません |
| **AGPL系ライセンス** | ✅ なし | AGPLライセンスのパッケージはありません |
| **商用利用** | ✅ 問題なし | すべてのライセンスが商用利用を許可しています |
| **ソースコード公開義務** | ✅ なし | コピーレフトライセンスが含まれていないため、ソースコード公開義務はありません |
| **特許関連** | ✅ 低リスク | Apache-2.0ライセンスには明示的な特許ライセンス付与が含まれています |

---

## 結論

✅ **本プロジェクトで使用されているすべてのライブラリは、商用利用に問題のないOSSライセンスで提供されています。**

主なポイント：
- 約84%がMITライセンス（最も寛容）
- GPL、AGPL等の強いコピーレフトライセンスは含まれていない
- 本番依存関係（実行時に必要なパッケージ）はcore-js（MIT）のみ
- Microsoft Office Addin関連パッケージもすべてMITライセンス

---

## 付録: 主要パッケージの詳細ライセンス情報

### ビルドツール系
- webpack (MIT) - Tobias Koppers
- Babel (MIT) - Babel Team
- ESLint (MIT) - OpenJS Foundation

### テストツール系  
- Jest (MIT) - Facebook/Meta

### Office Addin系
- office-addin-* (MIT) - Microsoft Corporation

### ユーティリティ系
- core-js (MIT) - Denis Pushkarev
- dotenv (BSD-2-Clause) - Scott Motte

---

*このレポートは `npx license-checker` を使用して自動生成された情報を元に作成されています。*


