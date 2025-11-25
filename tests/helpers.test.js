/**
 * helpers.js ヘルパー関数のユニットテスト
 */

describe('helpers.js', () => {
  let helpers;

  beforeEach(() => {
    jest.resetModules();

    // DOM をセットアップ
    document.body.innerHTML = `
      <select id="model-select">
        <option value="openai:gpt-4o">GPT-4o</option>
        <option value="claude:claude-sonnet-4-5-20250929">Claude Sonnet 4.5</option>
        <option value="gemini:gemini-2.0-flash">Gemini 2.0 Flash</option>
        <option value="openai-image:gpt-image-1">GPT Image 1</option>
      </select>
    `;

    helpers = require('../src/taskpane/helpers.js');
  });

  describe('getProviderAndModel', () => {
    // TC-PM-01: OpenAI モデルの解析
    test('TC-PM-01: OpenAI モデルを正しく解析する', () => {
      // Given: OpenAI モデルが選択されている
      document.getElementById('model-select').value = 'openai:gpt-4o';

      // When: getProviderAndModel を呼び出す
      const { provider, model } = helpers.getProviderAndModel();

      // Then: 正しく解析される
      expect(provider).toBe('openai');
      expect(model).toBe('gpt-4o');
    });

    // TC-PM-02: Claude モデルの解析
    test('TC-PM-02: Claude モデルを正しく解析する', () => {
      // Given: Claude モデルが選択されている
      document.getElementById('model-select').value = 'claude:claude-sonnet-4-5-20250929';

      // When: getProviderAndModel を呼び出す
      const { provider, model } = helpers.getProviderAndModel();

      // Then: 正しく解析される
      expect(provider).toBe('claude');
      expect(model).toBe('claude-sonnet-4-5-20250929');
    });

    // TC-PM-03: Gemini モデルの解析
    test('TC-PM-03: Gemini モデルを正しく解析する', () => {
      // Given: Gemini モデルが選択されている
      document.getElementById('model-select').value = 'gemini:gemini-2.0-flash';

      // When: getProviderAndModel を呼び出す
      const { provider, model } = helpers.getProviderAndModel();

      // Then: 正しく解析される
      expect(provider).toBe('gemini');
      expect(model).toBe('gemini-2.0-flash');
    });

    // TC-PM-04: 画像生成モデルの解析
    test('TC-PM-04: 画像生成モデルを正しく解析する', () => {
      // Given: 画像生成モデルが選択されている
      document.getElementById('model-select').value = 'openai-image:gpt-image-1';

      // When: getProviderAndModel を呼び出す
      const { provider, model } = helpers.getProviderAndModel();

      // Then: 正しく解析される
      expect(provider).toBe('openai-image');
      expect(model).toBe('gpt-image-1');
    });
  });

  describe('buildSystemPrompt', () => {
    // TC-SP-01: システムプロンプトの生成
    test('TC-SP-01: Excel情報を含むシステムプロンプトを生成する', () => {
      // Given: Excel のセル情報
      const excelPrompt = 'Selected Cell Address: A1, Selected Cell Value: 100';

      // When: buildSystemPrompt を呼び出す
      const result = helpers.buildSystemPrompt(excelPrompt);

      // Then: 正しいシステムプロンプトが生成される
      expect(result).toContain('Excelアドバイザー');
      expect(result).toContain('Selected Cell Address: A1');
      expect(result).toContain('Selected Cell Value: 100');
      expect(result).toContain('日本語で回答');
    });

    // TC-SP-02: 空のExcel情報でもプロンプトが生成される
    test('TC-SP-02: 空のExcel情報でもプロンプトが生成される', () => {
      // Given: 空のExcel情報
      const excelPrompt = '';

      // When: buildSystemPrompt を呼び出す
      const result = helpers.buildSystemPrompt(excelPrompt);

      // Then: 基本的なプロンプトが含まれる
      expect(result).toContain('Excelアドバイザー');
      expect(result).toContain('日本語で回答');
    });
  });

  describe('buildOpenAIContent', () => {
    // TC-OC-01: テキストのみの場合
    test('TC-OC-01: テキストのみの場合は文字列を返す', () => {
      // Given: ファイルなし
      const text = 'Hello, AI!';
      const files = [];

      // When: buildOpenAIContent を呼び出す
      const result = helpers.buildOpenAIContent(text, files);

      // Then: 文字列がそのまま返る
      expect(result).toBe('Hello, AI!');
    });

    // TC-OC-02: 画像ファイル付きの場合
    test('TC-OC-02: 画像ファイル付きの場合は配列を返す', () => {
      // Given: 画像ファイルあり
      const text = 'Describe this image';
      const files = [{
        name: 'test.png',
        type: 'image/png',
        base64: 'iVBORw0KGgoAAAANS...',
      }];

      // When: buildOpenAIContent を呼び出す
      const result = helpers.buildOpenAIContent(text, files);

      // Then: 配列形式で返る
      expect(Array.isArray(result)).toBe(true);
      expect(result).toHaveLength(2);
      expect(result[0]).toEqual({ type: 'text', text: 'Describe this image' });
      expect(result[1].type).toBe('image_url');
      expect(result[1].image_url.url).toContain('data:image/png;base64,');
    });

    // TC-OC-03: PDF ファイルは無視される（OpenAI Vision は PDF 非対応）
    test('TC-OC-03: PDF ファイルは無視される', () => {
      // Given: PDFファイルのみ
      const text = 'Read this PDF';
      const files = [{
        name: 'test.pdf',
        type: 'application/pdf',
        base64: 'JVBERi0xLjQ...',
      }];

      // When: buildOpenAIContent を呼び出す
      const result = helpers.buildOpenAIContent(text, files);

      // Then: テキストのみの配列（PDFは無視）
      expect(Array.isArray(result)).toBe(true);
      expect(result).toHaveLength(1);
      expect(result[0]).toEqual({ type: 'text', text: 'Read this PDF' });
    });

    // TC-OC-04: 複数画像ファイルの場合
    test('TC-OC-04: 複数画像ファイルの場合すべて含まれる', () => {
      // Given: 複数画像ファイル
      const text = 'Compare images';
      const files = [
        { name: 'img1.png', type: 'image/png', base64: 'abc' },
        { name: 'img2.jpg', type: 'image/jpeg', base64: 'xyz' },
      ];

      // When: buildOpenAIContent を呼び出す
      const result = helpers.buildOpenAIContent(text, files);

      // Then: すべての画像が含まれる
      expect(result).toHaveLength(3);
      expect(result[0].type).toBe('text');
      expect(result[1].type).toBe('image_url');
      expect(result[2].type).toBe('image_url');
    });
  });

  describe('buildClaudeContent', () => {
    // TC-CC-01: テキストのみの場合
    test('TC-CC-01: テキストのみの場合は文字列を返す', () => {
      // Given: ファイルなし
      const text = 'Hello, Claude!';
      const files = [];

      // When: buildClaudeContent を呼び出す
      const result = helpers.buildClaudeContent(text, files);

      // Then: 文字列がそのまま返る
      expect(result).toBe('Hello, Claude!');
    });

    // TC-CC-02: 画像ファイル付きの場合
    test('TC-CC-02: 画像ファイル付きの場合は配列を返す', () => {
      // Given: 画像ファイルあり
      const text = 'Describe this image';
      const files = [{
        name: 'test.jpg',
        type: 'image/jpeg',
        base64: '/9j/4AAQSkZJRg...',
      }];

      // When: buildClaudeContent を呼び出す
      const result = helpers.buildClaudeContent(text, files);

      // Then: 配列形式で返る（Claude の形式）
      expect(Array.isArray(result)).toBe(true);
      expect(result).toHaveLength(2);
      expect(result[0].type).toBe('image');
      expect(result[0].source.type).toBe('base64');
      expect(result[0].source.media_type).toBe('image/jpeg');
      expect(result[1]).toEqual({ type: 'text', text: 'Describe this image' });
    });

    // TC-CC-03: PDF ファイル付きの場合（Claude は対応）
    test('TC-CC-03: PDF ファイル付きの場合は document として含まれる', () => {
      // Given: PDFファイルあり
      const text = 'Read this PDF';
      const files = [{
        name: 'test.pdf',
        type: 'application/pdf',
        base64: 'JVBERi0xLjQ...',
      }];

      // When: buildClaudeContent を呼び出す
      const result = helpers.buildClaudeContent(text, files);

      // Then: document 形式で含まれる
      expect(Array.isArray(result)).toBe(true);
      expect(result).toHaveLength(2);
      expect(result[0].type).toBe('document');
      expect(result[0].source.type).toBe('base64');
      expect(result[0].source.media_type).toBe('application/pdf');
    });

    // TC-CC-04: 複数ファイルの場合
    test('TC-CC-04: 複数ファイルの場合すべて含まれる', () => {
      // Given: 複数ファイル
      const text = 'Analyze these';
      const files = [
        { name: 'image.png', type: 'image/png', base64: 'abc123' },
        { name: 'doc.pdf', type: 'application/pdf', base64: 'xyz789' },
      ];

      // When: buildClaudeContent を呼び出す
      const result = helpers.buildClaudeContent(text, files);

      // Then: すべてのファイルとテキストが含まれる
      expect(result).toHaveLength(3);
      expect(result[0].type).toBe('image');
      expect(result[1].type).toBe('document');
      expect(result[2].type).toBe('text');
    });
  });

  describe('buildGeminiParts', () => {
    // TC-GP-01: テキストのみの場合
    test('TC-GP-01: テキストのみの場合は text パートのみ', () => {
      // Given: ファイルなし
      const text = 'Hello, Gemini!';
      const files = [];

      // When: buildGeminiParts を呼び出す
      const result = helpers.buildGeminiParts(text, files);

      // Then: text パートのみの配列
      expect(result).toHaveLength(1);
      expect(result[0]).toEqual({ text: 'Hello, Gemini!' });
    });

    // TC-GP-02: 画像ファイル付きの場合
    test('TC-GP-02: 画像ファイル付きの場合は inline_data を含む', () => {
      // Given: 画像ファイルあり
      const text = 'Describe this image';
      const files = [{
        name: 'test.webp',
        type: 'image/webp',
        base64: 'UklGRlYA...',
      }];

      // When: buildGeminiParts を呼び出す
      const result = helpers.buildGeminiParts(text, files);

      // Then: inline_data と text パートを含む配列
      expect(result).toHaveLength(2);
      expect(result[0].inline_data).toBeDefined();
      expect(result[0].inline_data.mime_type).toBe('image/webp');
      expect(result[0].inline_data.data).toBe('UklGRlYA...');
      expect(result[1]).toEqual({ text: 'Describe this image' });
    });

    // TC-GP-03: PDF ファイル付きの場合（Gemini は対応）
    test('TC-GP-03: PDF ファイル付きの場合も inline_data として含まれる', () => {
      // Given: PDFファイルあり
      const text = 'Read this PDF';
      const files = [{
        name: 'test.pdf',
        type: 'application/pdf',
        base64: 'JVBERi0xLjQ...',
      }];

      // When: buildGeminiParts を呼び出す
      const result = helpers.buildGeminiParts(text, files);

      // Then: inline_data として含まれる
      expect(result).toHaveLength(2);
      expect(result[0].inline_data.mime_type).toBe('application/pdf');
    });

    // TC-GP-04: 非対応ファイルは除外される
    test('TC-GP-04: 非対応ファイルは除外される', () => {
      // Given: 非対応ファイル（テキストファイル）
      const text = 'Analyze this';
      const files = [{
        name: 'test.txt',
        type: 'text/plain',
        base64: 'SGVsbG8gV29ybGQ=',
      }];

      // When: buildGeminiParts を呼び出す
      const result = helpers.buildGeminiParts(text, files);

      // Then: text パートのみ（非対応ファイルは除外）
      expect(result).toHaveLength(1);
      expect(result[0]).toEqual({ text: 'Analyze this' });
    });
  });

  describe('buildCodeGenSystemPrompt', () => {
    // TC-CG-01: コード生成用プロンプトの生成
    test('TC-CG-01: Excel情報を含むコード生成用プロンプトを生成する', () => {
      // Given: Excel のセル情報
      const excelPrompt = 'Selected Cell Address: A1, Selected Cell Value: 100';

      // When: buildCodeGenSystemPrompt を呼び出す
      const result = helpers.buildCodeGenSystemPrompt(excelPrompt);

      // Then: 正しいプロンプトが生成される
      expect(result).toContain('Excelアドバイザー');
      expect(result).toContain('Excel JavaScript API');
      expect(result).toContain('json形式');
      expect(result).toContain('excel_code');
      expect(result).toContain('description');
      expect(result).toContain('Selected Cell Address: A1');
    });

    // TC-CG-02: JSON フォーマット例が含まれている
    test('TC-CG-02: JSON フォーマット例が含まれている', () => {
      // Given: 任意のExcel情報
      const excelPrompt = '';

      // When: buildCodeGenSystemPrompt を呼び出す
      const result = helpers.buildCodeGenSystemPrompt(excelPrompt);

      // Then: JSON フォーマット例が含まれる
      expect(result).toContain('"description"');
      expect(result).toContain('"excel_code"');
      expect(result).toContain('Excel.run');
    });
  });

  describe('fileToBase64', () => {
    // TC-FB-01: 画像ファイルを base64 に変換
    test('TC-FB-01: 画像ファイルを base64 に変換する', async () => {
      // Given: モックの File オブジェクト
      const blob = new Blob(['test'], { type: 'image/png' });
      const file = new File([blob], 'test.png', { type: 'image/png' });

      // FileReader のモック
      const mockFileReader = {
        readAsDataURL: jest.fn(),
        onload: null,
        onerror: null,
        result: 'data:image/png;base64,dGVzdA==',
      };
      global.FileReader = jest.fn(() => mockFileReader);

      // When: fileToBase64 を呼び出す
      const promise = helpers.fileToBase64(file);

      // FileReader の onload をトリガー
      mockFileReader.onload();

      const result = await promise;

      // Then: 正しい形式で返る
      expect(result.name).toBe('test.png');
      expect(result.type).toBe('image/png');
      expect(result.base64).toBe('dGVzdA==');
    });

    // TC-FB-02: ファイル読み込みエラー
    test('TC-FB-02: ファイル読み込みエラー時は reject される', async () => {
      // Given: エラーを発生させる FileReader
      const blob = new Blob(['test'], { type: 'image/png' });
      const file = new File([blob], 'test.png', { type: 'image/png' });

      const mockError = new Error('Read error');
      const mockFileReader = {
        readAsDataURL: jest.fn(),
        onload: null,
        onerror: null,
      };
      global.FileReader = jest.fn(() => mockFileReader);

      // When: fileToBase64 を呼び出す
      const promise = helpers.fileToBase64(file);

      // FileReader の onerror をトリガー
      mockFileReader.onerror(mockError);

      // Then: エラーで reject される
      await expect(promise).rejects.toEqual(mockError);
    });

    // TC-FB-03: PDF ファイルを base64 に変換
    test('TC-FB-03: PDF ファイルを base64 に変換する', async () => {
      // Given: モックの PDF File オブジェクト
      const blob = new Blob(['%PDF-1.4'], { type: 'application/pdf' });
      const file = new File([blob], 'document.pdf', { type: 'application/pdf' });

      const mockFileReader = {
        readAsDataURL: jest.fn(),
        onload: null,
        onerror: null,
        result: 'data:application/pdf;base64,JVBERi0xLjQ=',
      };
      global.FileReader = jest.fn(() => mockFileReader);

      // When: fileToBase64 を呼び出す
      const promise = helpers.fileToBase64(file);
      mockFileReader.onload();
      const result = await promise;

      // Then: 正しい形式で返る
      expect(result.name).toBe('document.pdf');
      expect(result.type).toBe('application/pdf');
      expect(result.base64).toBe('JVBERi0xLjQ=');
    });
  });
});

