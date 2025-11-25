/*
 * OpenLLM for Excel
 * Open source LLM-powered AI assistant for Microsoft Excel
 * Licensed under MIT License
 */

/* global console, document, Excel, Office */

import { 
  getApiKey, 
  saveApiKey, 
  deleteApiKey, 
  testApiKey, 
  ApiKeySource,
  restoreApiKeyToInput,
  getAllKeyStatus 
} from './config.js';

// グローバル変数で選択されたセルのアドレスとテキストを保持
let selectedCellAddress = "";
let selectedCellValue = "";
let conversationHistory = [];
let generatedExcelCode = "";
let uploadedFiles = []; // アップロードされたファイルのbase64データを保持

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("create-table").onclick = () => tryCatch(createTable);
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // セル選択変更イベントの登録
    await Excel.run(async (context) => {
      context.workbook.worksheets.onSelectionChanged.add(handleSelectionChange);
      await context.sync();
    }).catch((error) => {
      console.error("Error:", error);
    });

    // イベントリスナーの設定
    document.getElementById("refresh-cell-info").addEventListener("click", refreshCellInfo);
    document.getElementById("send-button").addEventListener("click", handleSendMessage);
    document.getElementById("execute-excel-code-button").addEventListener("click", handleExecuteExcelCode);
    document.getElementById("generate-image-button").addEventListener("click", handleGenerateImage);
    document.getElementById("file-input").addEventListener("change", handleFileUpload);
    document.getElementById("clear-files-button").addEventListener("click", clearUploadedFiles);
    
    // APIキー管理ボタンのイベントリスナー
    document.getElementById("save-api-key-button").addEventListener("click", handleSaveApiKey);
    document.getElementById("test-api-key-button").addEventListener("click", handleTestApiKey);
    document.getElementById("delete-api-key-button").addEventListener("click", handleDeleteApiKey);
    document.getElementById("model-select").addEventListener("change", handleModelChange);
    
    // Enterキーで送信
    document.getElementById("message-input").addEventListener("keypress", (e) => {
      if (e.key === "Enter") {
        handleSendMessage();
      }
    });

    // 初期化: 保存済みAPIキーを復元
    initializeApiKey();
  }
});

// APIキーの初期化と復元
function initializeApiKey() {
  const { provider } = getProviderAndModel();
  const { key, source } = getApiKey(provider);
  
  const statusEl = document.getElementById("api-key-status");
  
  if (source === ApiKeySource.ENV) {
    // 環境変数から読み込まれた場合
    statusEl.textContent = "✓ 環境変数からAPIキーを読み込みました";
    statusEl.className = "info";
    document.getElementById("api-key-input").placeholder = "環境変数から設定済み";
    document.getElementById("api-key-input").disabled = true;
  } else if (source === ApiKeySource.SAVED) {
    // 保存済みキーを復元
    restoreApiKeyToInput(provider);
    statusEl.textContent = "✓ 保存済みのAPIキーを復元しました";
    statusEl.className = "info";
  } else {
    statusEl.textContent = "";
    statusEl.className = "";
  }
  
  // 3秒後にステータスメッセージを消す
  if (source !== ApiKeySource.NONE) {
    setTimeout(() => {
      if (statusEl.textContent.includes("復元") || statusEl.textContent.includes("読み込み")) {
        statusEl.className = "";
        statusEl.textContent = "";
      }
    }, 3000);
  }
}

// モデル変更時にAPIキーを復元
function handleModelChange() {
  const { provider } = getProviderAndModel();
  const { source } = getApiKey(provider);
  
  // 環境変数がない場合のみ、保存済みキーを復元
  if (source !== ApiKeySource.ENV) {
    document.getElementById("api-key-input").disabled = false;
    restoreApiKeyToInput(provider);
  } else {
    document.getElementById("api-key-input").disabled = true;
    document.getElementById("api-key-input").value = "";
    document.getElementById("api-key-input").placeholder = "環境変数から設定済み";
  }
}

// APIキー保存ハンドラ
async function handleSaveApiKey() {
  const { provider } = getProviderAndModel();
  const key = document.getElementById("api-key-input").value;
  const statusEl = document.getElementById("api-key-status");
  
  if (!key) {
    statusEl.textContent = "⚠ APIキーを入力してください";
    statusEl.className = "error";
    return;
  }
  
  try {
    await saveApiKey(provider, key);
    statusEl.textContent = "✓ APIキーを保存しました";
    statusEl.className = "success";
  } catch (e) {
    statusEl.textContent = "✗ 保存に失敗しました: " + e.message;
    statusEl.className = "error";
  }
}

// APIキーテストハンドラ
async function handleTestApiKey() {
  const { provider } = getProviderAndModel();
  const { key, source } = getApiKey(provider);
  const statusEl = document.getElementById("api-key-status");
  
  if (!key) {
    statusEl.textContent = "⚠ APIキーを入力してください";
    statusEl.className = "error";
    return;
  }
  
  statusEl.textContent = "🔄 接続テスト中...";
  statusEl.className = "info";
  
  const result = await testApiKey(provider, key);
  
  if (result.success) {
    const sourceLabel = source === ApiKeySource.ENV ? "（環境変数）" : 
                       source === ApiKeySource.SAVED ? "（保存済み）" : "";
    statusEl.textContent = `✓ 接続成功 ${sourceLabel}`;
    statusEl.className = "success";
  } else {
    statusEl.textContent = "✗ 接続失敗: " + result.error;
    statusEl.className = "error";
  }
}

// APIキー削除ハンドラ
async function handleDeleteApiKey() {
  const { provider } = getProviderAndModel();
  const statusEl = document.getElementById("api-key-status");
  
  try {
    await deleteApiKey(provider);
    document.getElementById("api-key-input").value = "";
    statusEl.textContent = "✓ 保存済みAPIキーを削除しました";
    statusEl.className = "success";
  } catch (e) {
    statusEl.textContent = "✗ 削除に失敗しました: " + e.message;
    statusEl.className = "error";
  }
}

async function createTable() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const sampleDataRange = currentWorksheet.getRange("A1:E19");

    // サンプルデータの挿入
    sampleDataRange.values = [
      ["報告月", "支店", "売上[円]", "費用[円]", "利益[円]"], 
      ["2023/4/30", "東京", 10000000, 7000000, 3000000],
      ["2023/5/31", "東京", 9500000, 6800000, 2700000],
      ["2023/6/30", "東京", 11000000, 8000000, 3000000],
      ["2023/7/31", "東京", 10500000, 7500000, 3000000],
      ["2023/8/31", "東京", 12000000, 8500000, 3500000],
      ["2023/9/30", "東京", 11500000, 7700000, 3800000],
      ["2023/4/30", "大阪", 8000000, 6500000, 1500000],
      ["2023/5/31", "大阪", 7500000, 5500000, 2000000],
      ["2023/6/30", "大阪", 8200000, 6000000, 2200000],
      ["2023/7/31", "大阪", 7800000, 5800000, 2000000],
      ["2023/8/31", "大阪", 8500000, 6300000, 2200000],
      ["2023/9/30", "大阪", 9000000, 6700000, 2300000],
      ["2023/4/30", "福岡", 3000000, 2800000, 200000],
      ["2023/5/31", "福岡", 2500000, 2200000, 300000],
      ["2023/6/30", "福岡", 2000000, 1800000, 200000],
      ["2023/7/31", "福岡", 3200000, 2900000, 300000],
      ["2023/8/31", "福岡", 2800000, 2500000, 300000],
      ["2023/9/30", "福岡", 3500000, 3100000, 400000]
    ];

    // 罫線の設定
    sampleDataRange.format.borders.getItem('EdgeBottom').style = Excel.BorderLineStyle.continuous;
    sampleDataRange.format.borders.getItem('EdgeTop').style = Excel.BorderLineStyle.continuous;
    sampleDataRange.format.borders.getItem('EdgeLeft').style = Excel.BorderLineStyle.continuous;
    sampleDataRange.format.borders.getItem('EdgeRight').style = Excel.BorderLineStyle.continuous;
    sampleDataRange.format.borders.getItem('InsideVertical').style = Excel.BorderLineStyle.continuous;
    sampleDataRange.format.borders.getItem('InsideHorizontal').style = Excel.BorderLineStyle.continuous;

    // 数値データの書式設定
    const numberFormatRange = currentWorksheet.getRange("C2:E19");
    numberFormatRange.numberFormat = "#,##0";

    const dateFormatRange = currentWorksheet.getRange("A2:A19");
    dateFormatRange.numberFormat = "yyyy/m/d";

    // 列の幅と行の高さを自動調整
    sampleDataRange.format.autofitColumns();
    sampleDataRange.format.autofitRows();

    await context.sync();
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}

// プロバイダーとモデルを取得するヘルパー関数
function getProviderAndModel() {
  const modelValue = document.getElementById("model-select").value;
  const [provider, model] = modelValue.split(':');
  return { provider, model };
}

// ファイルをbase64に変換
function fileToBase64(file) {
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

// ファイルアップロード処理
async function handleFileUpload(event) {
  const files = event.target.files;
  const preview = document.getElementById("file-preview");
  const clearButton = document.getElementById("clear-files-button");
  
  for (const file of files) {
    try {
      const fileData = await fileToBase64(file);
      uploadedFiles.push(fileData);
      
      // プレビュー表示
      const previewItem = document.createElement("div");
      previewItem.className = "preview-item";
      
      if (file.type.startsWith("image/")) {
        const img = document.createElement("img");
        img.src = `data:${file.type};base64,${fileData.base64}`;
        previewItem.appendChild(img);
      } else {
        const icon = document.createElement("div");
        icon.textContent = "📄";
        icon.style.fontSize = "24px";
        previewItem.appendChild(icon);
      }
      
      const fileName = document.createElement("div");
      fileName.className = "file-name";
      fileName.textContent = file.name;
      previewItem.appendChild(fileName);
      
      preview.appendChild(previewItem);
    } catch (error) {
      console.error("File upload error:", error);
    }
  }
  
  if (uploadedFiles.length > 0) {
    clearButton.style.display = "inline-block";
  }
}

// アップロードファイルをクリア
function clearUploadedFiles() {
  uploadedFiles = [];
  document.getElementById("file-preview").innerHTML = "";
  document.getElementById("file-input").value = "";
  document.getElementById("clear-files-button").style.display = "none";
}

// システムプロンプトを生成
function buildSystemPrompt(excel_prompt) {
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
function buildOpenAIContent(text, files = []) {
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

// OpenAI APIにリクエストを送信する関数
async function sendToOpenAI(apiKey, model, system_prompt, userInput, aiMessageElement, files = []) {
  const url = "https://api.openai.com/v1/chat/completions";

  const userContent = buildOpenAIContent(userInput, files);
  conversationHistory.push({ "role": "user", "content": userContent });

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: model,
      messages: [
        { "role": "system", "content": system_prompt },
        ...conversationHistory
      ],
      "stream": true
    })
  });

  const reader = response.body.getReader();
  const decoder = new TextDecoder('utf-8');
  let streamedResponse = "AI: \n";

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;

    const text = decoder.decode(value, { stream: true });
    const lines = text.split(/\n+/);

    for (const line of lines) {
      const json_text = line.replace(/^data:\s*/, '');
      if (json_text === '[DONE]') {
        break;
      } else if (json_text) {
        try {
          const data = JSON.parse(json_text);
          const content = data.choices[0].delta.content;

          if (content) {
            streamedResponse += content;
            aiMessageElement.innerHTML += content;

            const messageArea = document.getElementById("message-area");
            messageArea.scrollTop = messageArea.scrollHeight;
          }
        } catch (e) {
          // JSON parse error - skip this chunk
        }
      }
    }
  }

  return streamedResponse;
}

// Claude用のメッセージコンテンツを構築（ファイル添付対応）
function buildClaudeContent(text, files = []) {
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

// Claude APIにリクエストを送信する関数
async function sendToClaude(apiKey, model, system_prompt, userInput, aiMessageElement, files = []) {
  const url = "https://api.anthropic.com/v1/messages";

  const userContent = buildClaudeContent(userInput, files);

  // Claude用のメッセージ履歴を構築（systemを除く）
  const claudeMessages = conversationHistory.map(msg => ({
    role: msg.role === "assistant" ? "assistant" : "user",
    content: msg.content
  }));
  claudeMessages.push({ "role": "user", "content": userContent });

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true'
    },
    body: JSON.stringify({
      model: model,
      max_tokens: 4096,
      system: system_prompt,
      messages: claudeMessages,
      stream: true
    })
  });

  const reader = response.body.getReader();
  const decoder = new TextDecoder('utf-8');
  let streamedResponse = "AI: \n";

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;

    const text = decoder.decode(value, { stream: true });
    const lines = text.split(/\n+/);

    for (const line of lines) {
      if (line.startsWith('data: ')) {
        const json_text = line.substring(6);
        try {
          const data = JSON.parse(json_text);
          if (data.type === 'content_block_delta' && data.delta?.text) {
            const content = data.delta.text;
            streamedResponse += content;
            aiMessageElement.innerHTML += content;

            const messageArea = document.getElementById("message-area");
            messageArea.scrollTop = messageArea.scrollHeight;
          }
        } catch (e) {
          // JSON parse error - skip this chunk
        }
      }
    }
  }

  conversationHistory.push({ "role": "user", "content": userInput });
  return streamedResponse;
}

// Gemini用のパーツを構築（ファイル添付対応）
function buildGeminiParts(text, files = []) {
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

// Gemini APIにリクエストを送信する関数
async function sendToGemini(apiKey, model, system_prompt, userInput, aiMessageElement, files = []) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:streamGenerateContent?alt=sse&key=${apiKey}`;

  // Gemini用のメッセージ履歴を構築
  const geminiContents = conversationHistory.map(msg => ({
    role: msg.role === "assistant" ? "model" : "user",
    parts: [{ text: typeof msg.content === 'string' ? msg.content : msg.content }]
  }));
  
  const userParts = buildGeminiParts(userInput, files);
  geminiContents.push({
    role: "user",
    parts: userParts
  });

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      systemInstruction: {
        parts: [{ text: system_prompt }]
      },
      contents: geminiContents
    })
  });

  const reader = response.body.getReader();
  const decoder = new TextDecoder('utf-8');
  let streamedResponse = "AI: \n";

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;

    const text = decoder.decode(value, { stream: true });
    const lines = text.split(/\n+/);

    for (const line of lines) {
      if (line.startsWith('data: ')) {
        const json_text = line.substring(6);
        try {
          const data = JSON.parse(json_text);
          if (data.candidates?.[0]?.content?.parts?.[0]?.text) {
            const content = data.candidates[0].content.parts[0].text;
            streamedResponse += content;
            aiMessageElement.innerHTML += content;

            const messageArea = document.getElementById("message-area");
            messageArea.scrollTop = messageArea.scrollHeight;
          }
        } catch (e) {
          // JSON parse error - skip this chunk
        }
      }
    }
  }

  conversationHistory.push({ "role": "user", "content": userInput });
  return streamedResponse;
}

// 統合されたLLM送信関数
async function sendToLLM(userInput, excel_prompt) {
  const { provider, model } = getProviderAndModel();
  const { key: apiKey, source } = getApiKey(provider);
  
  if (!apiKey) {
    throw new Error("APIキーが設定されていません。環境変数、保存済みキー、または入力欄から設定してください。");
  }
  const system_prompt = buildSystemPrompt(excel_prompt);
  const files = [...uploadedFiles]; // コピーを作成

  const aiMessageElement = document.createElement("div");
  aiMessageElement.innerHTML = "AI: ";
  document.getElementById("message-area").appendChild(aiMessageElement);

  // 送信後にファイルをクリア
  clearUploadedFiles();

  try {
    let streamedResponse;

    switch (provider) {
      case 'openai':
        streamedResponse = await sendToOpenAI(apiKey, model, system_prompt, userInput, aiMessageElement, files);
        break;
      case 'claude':
        streamedResponse = await sendToClaude(apiKey, model, system_prompt, userInput, aiMessageElement, files);
        break;
      case 'gemini':
        streamedResponse = await sendToGemini(apiKey, model, system_prompt, userInput, aiMessageElement, files);
        break;
      default:
        throw new Error(`Unknown provider: ${provider}`);
    }

    if (streamedResponse) {
      conversationHistory.push({ "role": "assistant", "content": streamedResponse });
      aiMessageElement.innerHTML = marked.parse(streamedResponse);
      const messageArea = document.getElementById("message-area");
      messageArea.scrollTop = messageArea.scrollHeight;
    }
  } catch (error) {
    console.error('Error:', error);
    aiMessageElement.innerHTML = `AI: エラーが発生しました: ${error.message}`;
  }
}

async function handleSendMessage() {
  const messageInput = document.getElementById("message-input");
  const messageArea = document.getElementById("message-area");
  const userInput = messageInput.value;

  if (userInput) {
    const excel_prompt = `Selected Cell Address: ${selectedCellAddress}, Selected Cell Value: ${selectedCellValue}`;

    const userMessageElement = document.createElement("div");
    userMessageElement.textContent = `You: ${userInput}`;
    messageArea.appendChild(userMessageElement);
    
    await sendToLLM(userInput, excel_prompt);

    messageInput.value = "";
    messageArea.scrollTop = messageArea.scrollHeight;
  }
}

async function handleSelectionChange(eventArgs) {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    range.load("text");
    await context.sync();

    selectedCellAddress = range.address;
    selectedCellValue = range.text;
    updateSelectedCellInfo();
  });
}

function updateSelectedCellInfo() {
  document.getElementById("selected-cell-address").textContent = `選択中のセル: ${selectedCellAddress}`;
  document.getElementById("selected-cell-value").textContent = `セルの値: ${selectedCellValue}`;
}

function showMessage(message) {
  const messageBar = document.getElementById("messageBar");
  messageBar.innerText = message;
  messageBar.style.display = "block";
  messageBar.style.backgroundColor = "#0078D4";
  messageBar.style.color = "white";
  messageBar.style.padding = "10px";
}

async function refreshCellInfo() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    range.load("text");
    await context.sync();

    const eventArgs = {
      address: range.address,
      text: range.text
    };
    await handleSelectionChange(eventArgs);
  }).catch((error) => {
    console.error("Error:", error);
  });
}

// Excel コード生成用のシステムプロンプト
function buildCodeGenSystemPrompt(excel_prompt) {
  const json_format = `{
    "description": "このマクロは、ExcelのA1セルとB1セルの値を足し合わせ、その結果をC1セルに表示します。さらに、C1セルの罫線を太くします。",
    "excel_code": "(async () => {\\n  await Excel.run(async (context) => {\\n    const sheet = context.workbook.worksheets.getActiveWorksheet();\\n    const rangeA1 = sheet.getRange('A1');\\n    const rangeB1 = sheet.getRange('B1');\\n    rangeA1.load('values');\\n    rangeB1.load('values');\\n    await context.sync();\\n    const sum = rangeA1.values[0][0] + rangeB1.values[0][0];\\n    const rangeC1 = sheet.getRange('C1');\\n    rangeC1.values = [[sum]];\\n    rangeC1.format.borders.getItem('EdgeBottom').style = 'Continuous';\\n    rangeC1.format.borders.getItem('EdgeBottom').weight = 'Thick';\\n    await context.sync();\\n  });\\n})();"
  }`;

  return `# 役割
あなたは優秀なExcelアドバイザーです。
# 条件
- 以下の項目は今開いているExcelの選択中のセルの情報です。
- セルのアドレスや内容を踏まえてExcel JavaScript APIでの処理を出力します。
${excel_prompt}
- 以下のようなjson形式で必ず出力します。
${json_format}
# 命令
`;
}

// OpenAI でコード生成
async function sendToOpenAIForCode(apiKey, model, system_prompt, userInput) {
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
        { "role": "system", "content": system_prompt },
        { "role": "user", "content": userInput }
      ],
      "stream": false,
      "response_format": {"type": "json_object"}
    })
  });

  const data = await response.json();
  return data.choices[0].message.content;
}

// Claude でコード生成
async function sendToClaudeForCode(apiKey, model, system_prompt, userInput) {
  const url = "https://api.anthropic.com/v1/messages";

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true'
    },
    body: JSON.stringify({
      model: model,
      max_tokens: 4096,
      system: system_prompt,
      messages: [
        { "role": "user", "content": userInput }
      ]
    })
  });

  const data = await response.json();
  return data.content[0].text;
}

// Gemini でコード生成
async function sendToGeminiForCode(apiKey, model, system_prompt, userInput) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      systemInstruction: {
        parts: [{ text: system_prompt }]
      },
      contents: [{
        role: "user",
        parts: [{ text: userInput }]
      }],
      generationConfig: {
        responseMimeType: 'application/json'
      }
    })
  });

  const data = await response.json();
  return data.candidates[0].content.parts[0].text;
}

// 統合されたコード生成関数
async function sendToLLMForExcelCode(userInput, excel_prompt) {
  const { provider, model } = getProviderAndModel();
  const { key: apiKey } = getApiKey(provider);
  
  if (!apiKey) {
    throw new Error("APIキーが設定されていません。環境変数、保存済みキー、または入力欄から設定してください。");
  }
  const system_prompt = buildCodeGenSystemPrompt(excel_prompt);

  try {
    let jsonResult;

    switch (provider) {
      case 'openai':
        jsonResult = await sendToOpenAIForCode(apiKey, model, system_prompt, userInput);
        break;
      case 'claude':
        jsonResult = await sendToClaudeForCode(apiKey, model, system_prompt, userInput);
        break;
      case 'gemini':
        jsonResult = await sendToGeminiForCode(apiKey, model, system_prompt, userInput);
        break;
      default:
        throw new Error(`Unknown provider: ${provider}`);
    }

    const parsedResult = JSON.parse(jsonResult);

    if (parsedResult.excel_code) {
      return {
        description: parsedResult.description,
        excelCode: parsedResult.excel_code
      };
    } else {
      console.log("Excel JavaScript APIコードが生成されませんでした。");
      return {
        description: "Excel JavaScript APIコードが生成されませんでした。",
        excelCode: ""
      };
    }
  } catch (error) {
    console.error('Error:', error);
    return {
      description: `エラーが発生しました: ${error.message}`,
      excelCode: ""
    };
  }
}

// 画像生成ハンドラー
async function handleGenerateImage() {
  const messageInput = document.getElementById("message-input");
  const messageArea = document.getElementById("message-area");
  const userInput = messageInput.value;
  const { provider, model } = getProviderAndModel();

  if (!userInput) {
    showMessage("画像生成するプロンプトを入力してください");
    return;
  }

  // 画像生成モデルかどうかをチェック
  if (provider !== 'openai-image' && provider !== 'gemini-image') {
    showMessage("画像生成には「OpenAI 画像生成」または「Google Gemini 画像生成」モデルを選択してください");
    return;
  }

  const userMessageElement = document.createElement("div");
  userMessageElement.textContent = `You: [画像生成] ${userInput}`;
  messageArea.appendChild(userMessageElement);

  const aiMessageElement = document.createElement("div");
  aiMessageElement.innerHTML = "AI: 画像を生成中...";
  messageArea.appendChild(aiMessageElement);

  messageInput.value = "";

  try {
    const { key: apiKey } = getApiKey(provider);
    
    if (!apiKey) {
      throw new Error("APIキーが設定されていません。環境変数、保存済みキー、または入力欄から設定してください。");
    }
    
    let imageUrl;

    if (provider === 'openai-image') {
      imageUrl = await generateImageOpenAI(apiKey, model, userInput);
    } else if (provider === 'gemini-image') {
      imageUrl = await generateImageGemini(apiKey, model, userInput);
    }

    if (imageUrl) {
      aiMessageElement.innerHTML = `AI: 画像を生成しました。<br/>
        <img src="${imageUrl}" class="generated-image" alt="Generated image" /><br/>
        <a href="${imageUrl}" download="generated-image.png" class="image-download-link">📥 画像をダウンロード</a>`;
    } else {
      aiMessageElement.innerHTML = "AI: 画像の生成に失敗しました。";
    }
  } catch (error) {
    console.error("Image generation error:", error);
    aiMessageElement.innerHTML = `AI: 画像生成エラー: ${error.message}`;
  }

  messageArea.scrollTop = messageArea.scrollHeight;
}

// OpenAI 画像生成
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
      response_format: "b64_json"
    })
  });

  const data = await response.json();
  
  if (data.error) {
    throw new Error(data.error.message);
  }

  if (data.data && data.data[0] && data.data[0].b64_json) {
    return `data:image/png;base64,${data.data[0].b64_json}`;
  }
  
  return null;
}

// Gemini 画像生成
async function generateImageGemini(apiKey, model, prompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      contents: [{
        role: "user",
        parts: [{ text: prompt }]
      }],
      generationConfig: {
        responseModalities: ["TEXT", "IMAGE"]
      }
    })
  });

  const data = await response.json();
  
  if (data.error) {
    throw new Error(data.error.message);
  }

  // Geminiのレスポンスから画像データを抽出
  if (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts) {
    for (const part of data.candidates[0].content.parts) {
      if (part.inline_data && part.inline_data.data) {
        const mimeType = part.inline_data.mime_type || 'image/png';
        return `data:${mimeType};base64,${part.inline_data.data}`;
      }
    }
  }
  
  return null;
}

async function handleExecuteExcelCode() {
  const messageInput = document.getElementById("message-input");
  const messageArea = document.getElementById("message-area");
  const userInput = messageInput.value;

  if (userInput) {
    const excel_prompt = `Selected Cell Address: ${selectedCellAddress}, Selected Cell Value: ${selectedCellValue}`;

    const userMessageElement = document.createElement("div");
    userMessageElement.textContent = `You: ${userInput}`;
    messageArea.appendChild(userMessageElement);

    const result = await sendToLLMForExcelCode(userInput, excel_prompt);
    generatedExcelCode = result.excelCode;

    const aiMessageElement = document.createElement("div");
    aiMessageElement.innerHTML = `AI: ${result.description}`;
    messageArea.appendChild(aiMessageElement);

    messageInput.value = "";
    messageArea.scrollTop = messageArea.scrollHeight;
  }

  if (generatedExcelCode) {
    try {
      const asyncFunction = new Function("Excel", generatedExcelCode);
      await Excel.run(async (context) => {
        await asyncFunction(Excel);
        await context.sync();
      });
    } catch (error) {
      console.error("Error executing Excel JavaScript API code:", error);
      showMessage(`実行エラー: ${error.message}`);
    }
  }
}
