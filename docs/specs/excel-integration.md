# Excel統合機能 設計仕様書

## 1. 概要

### 1.1 要望

- Excelで選択中のセルの情報をAIに伝えたい
- AIにExcelの操作コードを生成させて実行したい
- サンプルデータを簡単に作成したい

### 1.2 設計方針

**「Excel JavaScript API を活用したシームレス統合」**

```
[Excel]
   │
   ├── Office.onReady() で初期化
   │
   ├── セル選択イベント監視
   │       │
   │       ▼
   │   [選択セル情報をグローバル変数に保持]
   │
   └── Excel.run() でコード実行
           │
           ▼
       [AI生成コードを eval() で実行]
```

## 2. Excel JavaScript API

### 2.1 使用するAPI

| API | 用途 |
|-----|------|
| `Office.onReady()` | アドイン初期化 |
| `Excel.run()` | Excel操作のコンテキスト |
| `context.workbook.getSelectedRange()` | 選択範囲取得 |
| `worksheets.onSelectionChanged` | 選択変更イベント |
| `range.values` | セル値の読み書き |
| `range.format` | 書式設定 |

### 2.2 なぜExcel JavaScript APIか

| 選択肢 | メリット | デメリット | 判断 |
|--------|---------|-----------|------|
| VBA | 学習コスト低 | Web版非対応、古い | ❌ |
| VSTO | 高パフォーマンス | Windows限定 | ❌ |
| Excel JavaScript API | クロスプラットフォーム | Web制約 | ✅ |

## 3. 機能詳細

### 3.1 セル情報の取得

#### 初期化時のイベント登録

```javascript
Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    // セル選択変更イベントの登録
    await Excel.run(async (context) => {
      context.workbook.worksheets.onSelectionChanged.add(handleSelectionChange);
      await context.sync();
    });
  }
});
```

#### 選択変更ハンドラ

```javascript
async function handleSelectionChange(event) {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address, values");
    await context.sync();
    
    // グローバル変数に保存
    selectedCellAddress = range.address;
    selectedCellValue = JSON.stringify(range.values);
    
    // UI更新
    updateCellInfoDisplay();
  });
}
```

#### UIへの表示

```javascript
function updateCellInfoDisplay() {
  document.getElementById("selected-cell-address").textContent = 
    `選択中のセル: ${selectedCellAddress}`;
  document.getElementById("selected-cell-value").textContent = 
    `セルの値: ${selectedCellValue}`;
}
```

### 3.2 セル情報のAI提供

#### システムプロンプトへの組み込み

```javascript
function buildSystemPrompt(excel_prompt) {
  let base_prompt = `あなたはExcelに詳しいAIアシスタントです。
ユーザーはMicrosoft Excelを使用しています。`;

  if (excel_prompt) {
    base_prompt += `\n\n現在のExcel状態:\n${excel_prompt}`;
  }
  
  return base_prompt;
}

// 使用時
const excel_prompt = `選択中のセル: ${selectedCellAddress}
セルの値: ${selectedCellValue}`;

const system_prompt = buildSystemPrompt(excel_prompt);
```

### 3.3 コード生成と実行

#### マクロ実行フロー

```
[ユーザー入力] "A1からE10にヘッダー付きの表を作成して"
      │
      ▼
[LLMにコード生成依頼]
      │
      ▼
[JSON形式でコード取得]
{
  "code": "...",
  "explanation": "..."
}
      │
      ▼
[ユーザーに説明表示]
      │
      ▼
[Excel.run()内でコード実行]
```

#### コード生成用システムプロンプト

```javascript
function buildCodeGenSystemPrompt(excel_prompt) {
  return `あなたはExcel JavaScript APIを使用してExcelを操作するコードを生成するAIアシスタントです。

## 重要な制約
- contextは既に存在します。Excel.run()は使用しないでください。
- 必ず最後にawait context.sync()を呼び出してください。
- エラーハンドリングは不要です（呼び出し側で処理します）。

## 現在のExcel状態
${excel_prompt || "なし"}

## 出力形式
必ず以下のJSON形式で出力してください：
{
  "code": "実行するJavaScriptコード",
  "explanation": "コードの説明（日本語）"
}`;
}
```

#### コード実行関数

```javascript
async function executeExcelCode(code) {
  return await Excel.run(async (context) => {
    // AI生成コードを実行
    const asyncFunction = new Function('context', `
      return (async () => {
        ${code}
      })();
    `);
    
    await asyncFunction(context);
    return { success: true };
  });
}
```

**注意**: `eval()` ではなく `new Function()` を使用。
- スコープが明確
- `context` を引数として渡せる

### 3.4 サンプルテーブル作成

```javascript
async function createTable() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("A1:E19");
    
    // データ設定
    range.values = [
      ["報告月", "支店", "売上[円]", "費用[円]", "利益[円]"],
      ["2023/4/30", "東京", 10000000, 7000000, 3000000],
      // ...
    ];
    
    // 罫線設定
    range.format.borders.getItem('EdgeBottom').style = 
      Excel.BorderLineStyle.continuous;
    // ...
    
    // 数値フォーマット
    sheet.getRange("C2:E19").numberFormat = "#,##0";
    
    // 列幅自動調整
    range.format.autofitColumns();
    
    await context.sync();
  });
}
```

## 4. セキュリティ考慮

### 4.1 コード実行のリスク

| リスク | 対策 |
|--------|------|
| 悪意あるコード実行 | AIが生成するため基本的に安全。ユーザーが確認してから実行 |
| 無限ループ | タイムアウト未実装（将来検討） |
| 大量セル操作 | API側で制限あり |

### 4.2 実装上の注意

- AI生成コードを**直接実行**する設計
- ユーザーには**説明を表示**してから実行
- 本番環境では**サンドボックス化**を検討

## 5. テスト仕様

### 5.1 テストケース

| テストID | シナリオ | 前提条件 | 手順 | 期待結果 |
|----------|---------|---------|------|---------|
| EXCEL-001 | セル選択検出 | Excelでアドイン起動 | セル選択変更 | アドレス・値が更新 |
| EXCEL-002 | サンプルテーブル作成 | - | ボタンクリック | A1:E19にデータ |
| EXCEL-003 | コード生成 | - | 「A1に'Hello'と入力して」→マクロ実行 | A1に文字入力 |
| EXCEL-004 | 書式設定生成 | - | 「A1を太字にして」→マクロ実行 | A1が太字に |
| EXCEL-005 | 数式生成 | データあり | 「合計を計算して」→マクロ実行 | SUM関数が入力 |

### 5.2 合格基準

| 基準ID | 基準 | 検証方法 |
|--------|------|---------|
| AC-001 | セル選択がリアルタイム反映 | 複数セル選択して確認 |
| AC-002 | サンプルテーブルが正しく作成 | 罫線・書式含め確認 |
| AC-003 | AI生成コードが実行される | 複数パターンで確認 |

## 6. 制限事項

| 制限 | 理由 | 回避策 |
|------|------|--------|
| VBAマクロは実行不可 | JavaScript API限定 | JavaScript相当の処理に変換 |
| 大量データ操作は遅い | Web API制約 | バッチ処理、チャンク分割 |
| 一部Excel機能は未対応 | API制限 | 対応APIに限定 |

## 7. 今後の拡張

### 7.1 検討中の機能

- コード実行前のプレビュー/確認ダイアログ
- 実行履歴の保存
- よく使う操作のテンプレート化
- コード実行のUndo機能

### 7.2 Excel JavaScript API の発展

Microsoftは継続的にAPIを拡張中。新機能が追加されたら対応を検討:
- Power Query連携
- ピボットテーブル操作の強化
- グラフ操作の強化

