/*
 * OpenLLM for Excel
 * Open source LLM-powered AI assistant for Microsoft Excel
 * Licensed under MIT License
 */

/* global console, document, Excel, Office */

// グローバル変数で選択されたセルのアドレスとテキストを保持
let selectedCellAddress = "";
let selectedCellValue = "";
let conversationHistory = [];
let generatedExcelCode = "";

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
    
    // Enterキーで送信
    document.getElementById("message-input").addEventListener("keypress", (e) => {
      if (e.key === "Enter") {
        handleSendMessage();
      }
    });
  }
});

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

// OpenAI APIにリクエストを送信する関数
async function sendToOpenAI(userInput, excel_prompt) {
  const apiKey = document.getElementById("api-key-input").value;
  const llm_model = document.getElementById("model-select").value;

  const url = "https://api.openai.com/v1/chat/completions";
  const system_prompt = `# 役割
あなたは優秀なExcelアドバイザーです。
基本的にはExcelの仕様として回答します。
何かの操作を指示されたら、具体的な指示がない場合は目的を達成するために一般的なExcelの操作方法を教えてください。
日本語で回答してください。
以下の項目は今開いているExcelの選択中のセルの情報です。
セルのアドレスや内容を踏まえて回答します。
3行以上あるデータは2行までのデータのみ記入しています。
${excel_prompt}`;

  const aiMessageElement = document.createElement("div");
  aiMessageElement.innerHTML = "AI: ";
  document.getElementById("message-area").appendChild(aiMessageElement);

  try {
    conversationHistory.push({ "role": "user", "content": userInput });

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: llm_model,
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
    
    await sendToOpenAI(userInput, excel_prompt);

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

async function sendToOpenAIForExcelCode(userInput, excel_prompt) {
  const apiKey = document.getElementById("api-key-input").value;
  const llm_model = document.getElementById("model-select").value;
  const json_format = `{
    "description": "このマクロは、ExcelのA1セルとB1セルの値を足し合わせ、その結果をC1セルに表示します。さらに、C1セルの罫線を太くします。",
    "excel_code": "(async () => {\\n  await Excel.run(async (context) => {\\n    const sheet = context.workbook.worksheets.getActiveWorksheet();\\n    const rangeA1 = sheet.getRange('A1');\\n    const rangeB1 = sheet.getRange('B1');\\n    rangeA1.load('values');\\n    rangeB1.load('values');\\n    await context.sync();\\n    const sum = rangeA1.values[0][0] + rangeB1.values[0][0];\\n    const rangeC1 = sheet.getRange('C1');\\n    rangeC1.values = [[sum]];\\n    rangeC1.format.borders.getItem('EdgeBottom').style = 'Continuous';\\n    rangeC1.format.borders.getItem('EdgeBottom').weight = 'Thick';\\n    await context.sync();\\n  });\\n})();"
  }`;

  const url = "https://api.openai.com/v1/chat/completions";
  const system_prompt = `# 役割
あなたは優秀なExcelアドバイザーです。
# 条件
- 以下の項目は今開いているExcelの選択中のセルの情報です。
- セルのアドレスや内容を踏まえてExcel JavaScript APIでの処理を出力します。
${excel_prompt}
- 以下のようなjson形式で必ず出力します。
${json_format}
# 命令
`;
  const json_response_format = {"type": "json_object"};

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: llm_model,
        messages: [
          { "role": "system", "content": system_prompt },
          { "role": "user", "content": userInput }
        ],
        "stream": false,
        "response_format": json_response_format
      })
    });

    const data = await response.json();
    const jsonResult = data.choices[0].message.content;
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

async function handleExecuteExcelCode() {
  const messageInput = document.getElementById("message-input");
  const messageArea = document.getElementById("message-area");
  const userInput = messageInput.value;

  if (userInput) {
    const excel_prompt = `Selected Cell Address: ${selectedCellAddress}, Selected Cell Value: ${selectedCellValue}`;

    const userMessageElement = document.createElement("div");
    userMessageElement.textContent = `You: ${userInput}`;
    messageArea.appendChild(userMessageElement);

    const result = await sendToOpenAIForExcelCode(userInput, excel_prompt);
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
