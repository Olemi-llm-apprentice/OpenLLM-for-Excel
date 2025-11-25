/*
 * OpenLLM for Excel - Helper Functions
 * テスト可能なヘルパー関数を分離
 */

// プロバイダーとモデルを取得するヘルパー関数
export function getProviderAndModel() {
  const modelValue = document.getElementById("model-select").value;
  const [provider, model] = modelValue.split(':');
  return { provider, model };
}

// ファイルをbase64に変換
export function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => {
      const base64 = reader.result.split(',')[1];
      resolve({
        name: file.name,
        type: file.type,
        base64: base64
      });
    };
    reader.onerror = error => reject(error);
  });
}

// システムプロンプトを生成
export function buildSystemPrompt(excel_prompt) {
  return `# 役割
あなたは優秀なExcelアドバイザーです。
基本的にはExcelの仕様として回答します。
何かの操作を指示されたら、具体的な指示がない場合は目的を達成するために一般的なExcelの操作方法を教えてください。
日本語で回答してください。
以下の項目は今開いているExcelの選択中のセルの情報です。
セルのアドレスや内容を踏まえて回答します。
3行以上あるデータは2行までのデータのみ記入しています。
${excel_prompt}`;
}

// OpenAI用のメッセージコンテンツを構築（ファイル添付対応）
export function buildOpenAIContent(text, files = []) {
  if (files.length === 0) {
    return text;
  }
  
  const content = [{ type: "text", text: text }];
  
  for (const file of files) {
    if (file.type.startsWith("image/")) {
      content.push({
        type: "image_url",
        image_url: {
          url: `data:${file.type};base64,${file.base64}`
        }
      });
    }
    // PDFは現時点ではOpenAI Vision APIでは直接サポートされていない
  }
  
  return content;
}

// Claude用のメッセージコンテンツを構築（ファイル添付対応）
export function buildClaudeContent(text, files = []) {
  if (files.length === 0) {
    return text;
  }
  
  const content = [];
  
  for (const file of files) {
    if (file.type.startsWith("image/")) {
      content.push({
        type: "image",
        source: {
          type: "base64",
          media_type: file.type,
          data: file.base64
        }
      });
    } else if (file.type === "application/pdf") {
      content.push({
        type: "document",
        source: {
          type: "base64",
          media_type: file.type,
          data: file.base64
        }
      });
    }
  }
  
  content.push({ type: "text", text: text });
  
  return content;
}

// Gemini用のパーツを構築（ファイル添付対応）
export function buildGeminiParts(text, files = []) {
  const parts = [];
  
  for (const file of files) {
    if (file.type.startsWith("image/") || file.type === "application/pdf") {
      parts.push({
        inline_data: {
          mime_type: file.type,
          data: file.base64
        }
      });
    }
  }
  
  parts.push({ text: text });
  
  return parts;
}

// Excel コード生成用のシステムプロンプト
export function buildCodeGenSystemPrompt(excel_prompt) {
  const json_format = `{
    "description": "このマクロは、ExcelのA1セルとB1セルの値を足し合わせ、その結果をC1セルに表示します。さらに、C1セルの罫線を太くします。",
    "excel_code": "(async () => {\\n  await Excel.run(async (context) => {\\n    const sheet = context.workbook.worksheets.getActiveWorksheet();\\n    const rangeA1 = sheet.getRange('A1');\\n    const rangeB1 = sheet.getRange('B1');\\n    rangeA1.load('values');\\n    rangeB1.load('values');\\n    await context.sync();\\n    const sum = rangeA1.values[0][0] + rangeB1.values[0][0];\\n    const rangeC1 = sheet.getRange('C1');\\n    rangeC1.values = [[sum]];\\n    rangeC1.format.borders.getItem('EdgeBottom').style = 'Continuous';\\n    rangeC1.format.borders.getItem('EdgeBottom').weight = 'Thick';\\n    await context.sync();\\n  });\\n})();"
  }`;

  return `# 役割
あなたは優秀なExcelアドバイザーです。

# 重要な制約
- **Excel JavaScript API でサポートされている機能のみ使用すること**
- VBA や COM オブジェクトは使用不可
- 以下のAPIは使用可能: Range, Worksheet, Workbook, Table, Chart, PivotTable, NamedItem, ConditionalFormat
- 以下のAPIは制限あり/未サポート: PivotChart（グラフはChartで作成）、マクロ記録、ActiveX
- プロパティにアクセスする前に必ず load() と context.sync() を呼ぶこと
- 非同期処理は async/await パターンを使用すること

# 条件
- 以下の項目は今開いているExcelの選択中のセルの情報です。
- セルのアドレスや内容を踏まえてExcel JavaScript APIでの処理を出力します。
${excel_prompt}
- 以下のようなjson形式で必ず出力します。
${json_format}

# 命令
`;
}

