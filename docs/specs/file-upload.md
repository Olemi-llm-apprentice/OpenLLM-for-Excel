# ファイルアップロード（Vision）機能 設計仕様書

## 1. 概要

### 1.1 要望

- 画像をアップロードしてAIに分析させたい（Vision機能）
- PDFもアップロードして内容を読み取らせたい
- 複数ファイルを同時にアップロードしたい
- アップロード前にプレビューを確認したい

### 1.2 設計方針

**「Base64エンコードによるインライン添付」** を採用

```
[ファイル選択]
      │
      ▼
[File → Base64変換]
      │
      ▼
[uploadedFiles配列に保存]
      │
      ▼
[送信時にAPIリクエストに含める]
```

## 2. 対応状況

### 2.1 プロバイダー別対応

| プロバイダー | 画像入力 | PDF入力 | 備考 |
|-------------|:-------:|:-------:|------|
| OpenAI GPT-4o | ✅ | ❌ | Vision対応 |
| OpenAI GPT-4o mini | ✅ | ❌ | Vision対応 |
| Claude（全モデル） | ✅ | ✅ | PDF対応が強み |
| Gemini（全モデル） | ✅ | ✅ | PDF対応 |

### 2.2 対応ファイル形式

| 形式 | MIMEタイプ | 最大サイズ |
|------|-----------|-----------|
| JPEG | `image/jpeg` | 20MB |
| PNG | `image/png` | 20MB |
| GIF | `image/gif` | 20MB |
| WebP | `image/webp` | 20MB |
| PDF | `application/pdf` | 20MB（Claude/Gemini） |

## 3. 設計詳細

### 3.1 ファイル入力UI

```html
<div id="file-upload-container">
  <label for="file-input">📎 ファイル添付（画像/PDF）:</label>
  <input type="file" id="file-input" accept="image/*,.pdf" multiple />
  <div id="file-preview"></div>
  <button id="clear-files-button" style="display:none;">添付をクリア</button>
</div>
```

**属性の意図**:
- `accept="image/*,.pdf"`: 画像とPDFのみ選択可能
- `multiple`: 複数ファイル選択可能

### 3.2 データ構造

```javascript
// グローバル変数
let uploadedFiles = [];

// 各ファイルの構造
{
  name: "example.png",
  type: "image/png",
  base64: "iVBORw0KGgoAAAANSUhEUgAA...",
}
```

### 3.3 ファイル読み込み処理

```javascript
async function handleFileUpload(event) {
  const files = event.target.files;
  
  for (const file of files) {
    const base64 = await fileToBase64(file);
    uploadedFiles.push({
      name: file.name,
      type: file.type,
      base64: base64,
    });
  }
  
  updateFilePreview();
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      // data:image/png;base64,xxxxx から base64部分だけ抽出
      const base64 = reader.result.split(',')[1];
      resolve(base64);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}
```

### 3.4 プレビュー表示

```javascript
function updateFilePreview() {
  const previewContainer = document.getElementById("file-preview");
  const clearButton = document.getElementById("clear-files-button");
  
  previewContainer.innerHTML = uploadedFiles.map(file => {
    if (file.type.startsWith('image/')) {
      return `
        <div class="preview-item">
          <img src="data:${file.type};base64,${file.base64}" />
          <span class="file-name">${file.name}</span>
        </div>
      `;
    } else {
      return `
        <div class="preview-item">
          <span class="file-icon">📄</span>
          <span class="file-name">${file.name}</span>
        </div>
      `;
    }
  }).join('');
  
  clearButton.style.display = uploadedFiles.length > 0 ? 'block' : 'none';
}
```

### 3.5 プロバイダー別リクエスト形式

#### OpenAI（Vision）

```javascript
function buildOpenAIContent(userInput, files) {
  if (files.length === 0) {
    return userInput;
  }
  
  const content = [
    { type: "text", text: userInput }
  ];
  
  for (const file of files) {
    if (file.type.startsWith('image/')) {
      content.push({
        type: "image_url",
        image_url: {
          url: `data:${file.type};base64,${file.base64}`
        }
      });
    }
    // OpenAI は PDF 未対応のためスキップ
  }
  
  return content;
}
```

#### Claude

```javascript
function buildClaudeContent(userInput, files) {
  const content = [];
  
  for (const file of files) {
    if (file.type.startsWith('image/')) {
      content.push({
        type: "image",
        source: {
          type: "base64",
          media_type: file.type,
          data: file.base64,
        }
      });
    } else if (file.type === 'application/pdf') {
      content.push({
        type: "document",
        source: {
          type: "base64",
          media_type: file.type,
          data: file.base64,
        }
      });
    }
  }
  
  content.push({ type: "text", text: userInput });
  
  return content;
}
```

#### Gemini

```javascript
function buildGeminiParts(userInput, files) {
  const parts = [];
  
  for (const file of files) {
    parts.push({
      inlineData: {
        mimeType: file.type,
        data: file.base64,
      }
    });
  }
  
  parts.push({ text: userInput });
  
  return parts;
}
```

### 3.6 送信後のクリア

```javascript
// メッセージ送信後
function clearUploadedFiles() {
  uploadedFiles = [];
  document.getElementById("file-input").value = "";
  updateFilePreview();
}
```

**理由**: 送信後もファイルが残っていると、次のメッセージにも添付されてしまう。

## 4. UI設計

### 4.1 スタイル

```css
#file-upload-container {
  background-color: #f9f9f9;
  border: 1px dashed #ccc;
  border-radius: 4px;
  padding: 8px;
  margin-top: 10px;
}

#file-preview {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  margin-top: 8px;
}

#file-preview .preview-item {
  position: relative;
  border: 1px solid #ddd;
  border-radius: 4px;
  padding: 4px;
  background: white;
}

#file-preview img {
  max-width: 60px;
  max-height: 60px;
  object-fit: cover;
}
```

## 5. テスト仕様

### 5.1 テストケース

| テストID | シナリオ | 前提条件 | 手順 | 期待結果 |
|----------|---------|---------|------|---------|
| FILE-001 | 画像アップロード（単体） | - | PNG選択 | プレビュー表示 |
| FILE-002 | 画像アップロード（複数） | - | 3枚選択 | 3枚のプレビュー |
| FILE-003 | PDFアップロード | - | PDF選択 | ファイル名表示 |
| FILE-004 | OpenAI+画像 | GPT-4o選択 | 画像+質問送信 | 画像を認識した回答 |
| FILE-005 | Claude+PDF | Sonnet選択 | PDF+質問送信 | PDFを読んだ回答 |
| FILE-006 | Gemini+画像 | Flash選択 | 画像+質問送信 | 画像を認識した回答 |
| FILE-007 | クリアボタン | ファイル添付済み | クリアボタン | プレビュー消去 |
| FILE-008 | 送信後自動クリア | ファイル添付済み | 送信 | 自動的にクリア |

### 5.2 合格基準

| 基準ID | 基準 | 検証方法 |
|--------|------|---------|
| AC-001 | 画像プレビューが正しく表示 | 目視確認 |
| AC-002 | 3プロバイダーで画像認識動作 | 画像の内容を質問 |
| AC-003 | Claude/GeminiでPDF認識動作 | PDFの内容を質問 |
| AC-004 | 送信後にファイルがクリアされる | 送信後確認 |

## 6. 制限事項

| 制限 | 理由 | 回避策 |
|------|------|--------|
| OpenAIはPDF未対応 | API仕様 | Claude/Gemini使用 |
| ファイルサイズ上限 | Base64膨張、API制限 | 大きいファイルは事前圧縮 |
| ドラッグ&ドロップ未対応 | シンプルさ優先 | 将来実装検討 |

## 7. 今後の拡張

### 7.1 検討中の機能

- ドラッグ&ドロップ対応
- 画像のリサイズ/圧縮
- Excel内のセル範囲をスクリーンショット化して添付
- クリップボードからの貼り付け

