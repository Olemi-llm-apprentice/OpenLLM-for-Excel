/**
 * config.js ユニットテスト
 * APIキー管理機能のテスト
 */

// テスト対象のモジュールをテスト用にラップしてインポート
// 環境変数モックを設定するため、動的にrequire

// 環境変数のモック用グローバル変数
global.__OPENAI_API_KEY__ = '';
global.__ANTHROPIC_API_KEY__ = '';
global.__GEMINI_API_KEY__ = '';

describe('config.js', () => {
  let config;

  beforeEach(() => {
    // 環境変数をリセット
    global.__OPENAI_API_KEY__ = '';
    global.__ANTHROPIC_API_KEY__ = '';
    global.__GEMINI_API_KEY__ = '';

    // モジュールキャッシュをクリア
    jest.resetModules();

    // DOMをセットアップ
    document.body.innerHTML = `
      <input type="password" id="api-key-input" />
    `;
  });

  describe('normalizeProvider', () => {
    beforeEach(() => {
      config = require('../src/taskpane/config.js');
    });

    // TC-N-01: 通常のプロバイダー名はそのまま返す
    test('TC-N-01: openai はそのまま openai を返す', () => {
      // Given: 通常のプロバイダー名 'openai'
      const provider = 'openai';

      // When: normalizeProvider を呼び出す（内部関数のため getApiKey 経由でテスト）
      // Then: getApiKey が正しく動作することで検証
      const { key, source } = config.getApiKey(provider);
      expect(source).toBe(config.ApiKeySource.NONE); // キー未設定のため
    });

    // TC-N-02: openai-image は openai に正規化
    test('TC-N-02: openai-image は openai に正規化される', () => {
      // Given: 画像生成用プロバイダー名 'openai-image'
      global.__OPENAI_API_KEY__ = 'test-openai-key';
      jest.resetModules();
      config = require('../src/taskpane/config.js');

      // When: openai-image でキーを取得
      const { key, source } = config.getApiKey('openai-image');

      // Then: openai のキーが取得される
      expect(key).toBe('test-openai-key');
      expect(source).toBe(config.ApiKeySource.ENV);
    });

    // TC-N-03: gemini-image は gemini に正規化
    test('TC-N-03: gemini-image は gemini に正規化される', () => {
      // Given: 画像生成用プロバイダー名 'gemini-image'
      global.__GEMINI_API_KEY__ = 'test-gemini-key';
      jest.resetModules();
      config = require('../src/taskpane/config.js');

      // When: gemini-image でキーを取得
      const { key, source } = config.getApiKey('gemini-image');

      // Then: gemini のキーが取得される
      expect(key).toBe('test-gemini-key');
      expect(source).toBe(config.ApiKeySource.ENV);
    });

    // TC-N-04: claude はそのまま claude を返す
    test('TC-N-04: claude はそのまま claude を返す', () => {
      // Given: 通常のプロバイダー名 'claude'
      global.__ANTHROPIC_API_KEY__ = 'test-claude-key';
      jest.resetModules();
      config = require('../src/taskpane/config.js');

      // When: claude でキーを取得
      const { key, source } = config.getApiKey('claude');

      // Then: claude のキーが取得される
      expect(key).toBe('test-claude-key');
      expect(source).toBe(config.ApiKeySource.ENV);
    });
  });

  describe('getApiKey - 優先順位テスト', () => {
    // TC-A-01: 環境変数が最優先
    test('TC-A-01: 環境変数が設定されている場合は ENV が返る', () => {
      // Given: 環境変数にAPIキーが設定されている
      global.__OPENAI_API_KEY__ = 'env-key';
      jest.resetModules();
      config = require('../src/taskpane/config.js');

      // When: getApiKey を呼び出す
      const { key, source } = config.getApiKey('openai');

      // Then: 環境変数のキーが返る
      expect(key).toBe('env-key');
      expect(source).toBe(config.ApiKeySource.ENV);
    });

    // TC-A-02: 環境変数がなければ保存済みキー
    test('TC-A-02: 保存済みキーがある場合は SAVED が返る', () => {
      // Given: 環境変数なし、Office Settings にキーが保存されている
      global.__OPENAI_API_KEY__ = '';
      jest.resetModules();
      config = require('../src/taskpane/config.js');
      global.Office.context.roamingSettings._storage['openai_api_key'] = 'saved-key';

      // When: getApiKey を呼び出す
      const { key, source } = config.getApiKey('openai');

      // Then: 保存済みのキーが返る
      expect(key).toBe('saved-key');
      expect(source).toBe(config.ApiKeySource.SAVED);
    });

    // TC-A-03: 保存済みもなければ入力フィールド
    test('TC-A-03: UI入力フィールドにキーがある場合は INPUT が返る', () => {
      // Given: 環境変数なし、保存済みキーなし、UI入力にキーあり
      global.__OPENAI_API_KEY__ = '';
      jest.resetModules();
      config = require('../src/taskpane/config.js');
      document.getElementById('api-key-input').value = 'input-key';

      // When: getApiKey を呼び出す
      const { key, source } = config.getApiKey('openai');

      // Then: UI入力のキーが返る
      expect(key).toBe('input-key');
      expect(source).toBe(config.ApiKeySource.INPUT);
    });

    // TC-A-04: どこにもなければ NONE
    test('TC-A-04: キーがどこにもない場合は NONE が返る', () => {
      // Given: どこにもキーがない
      global.__OPENAI_API_KEY__ = '';
      jest.resetModules();
      config = require('../src/taskpane/config.js');

      // When: getApiKey を呼び出す
      const { key, source } = config.getApiKey('openai');

      // Then: 空文字と NONE が返る
      expect(key).toBe('');
      expect(source).toBe(config.ApiKeySource.NONE);
    });
  });

  describe('saveApiKey', () => {
    beforeEach(() => {
      config = require('../src/taskpane/config.js');
    });

    // TC-S-01: Office Settings API で保存成功
    test('TC-S-01: Office Settings API で保存が成功する', async () => {
      // Given: Office Settings API が利用可能
      // When: saveApiKey を呼び出す
      const result = await config.saveApiKey('openai', 'new-key');

      // Then: 保存が成功する
      expect(result).toBe(true);
      expect(global.Office.context.roamingSettings.set).toHaveBeenCalledWith('openai_api_key', 'new-key');
      expect(global.Office.context.roamingSettings.saveAsync).toHaveBeenCalled();
    });

    // TC-S-02: Office未対応時は localStorage にフォールバック
    test('TC-S-02: Office未対応時は localStorage に保存', async () => {
      // Given: Office Settings API が利用不可
      const originalOffice = global.Office;
      global.Office = { context: null };
      jest.resetModules();
      config = require('../src/taskpane/config.js');

      // When: saveApiKey を呼び出す
      const result = await config.saveApiKey('openai', 'fallback-key');

      // Then: localStorage に保存される
      expect(result).toBe(true);
      expect(global.localStorage.setItem).toHaveBeenCalledWith('openai_api_key', 'fallback-key');

      // クリーンアップ
      global.Office = originalOffice;
    });
  });

  describe('deleteApiKey', () => {
    beforeEach(() => {
      config = require('../src/taskpane/config.js');
    });

    // TC-D-01: Office Settings API で削除成功
    test('TC-D-01: Office Settings API で削除が成功する', async () => {
      // Given: Office Settings API が利用可能、キーが保存されている
      global.Office.context.roamingSettings._storage['openai_api_key'] = 'to-delete';

      // When: deleteApiKey を呼び出す
      const result = await config.deleteApiKey('openai');

      // Then: 削除が成功する
      expect(result).toBe(true);
      expect(global.Office.context.roamingSettings.remove).toHaveBeenCalledWith('openai_api_key');
    });
  });

  describe('testApiKey', () => {
    beforeEach(() => {
      config = require('../src/taskpane/config.js');
    });

    // TC-T-01: 不明なプロバイダーではエラー
    test('TC-T-01: 不明なプロバイダーではエラーを返す', async () => {
      // Given: 不明なプロバイダー名
      // When: testApiKey を呼び出す
      const result = await config.testApiKey('unknown', 'some-key');

      // Then: エラーが返る
      expect(result.success).toBe(false);
      expect(result.error).toContain('Unknown provider');
    });

    // TC-T-02: OpenAI APIキーテスト成功
    test('TC-T-02: OpenAI APIキーテストが成功する', async () => {
      // Given: fetch がOKレスポンスを返す
      global.fetch.mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ data: [] }),
      });

      // When: testApiKey を呼び出す
      const result = await config.testApiKey('openai', 'valid-key');

      // Then: 成功が返る
      expect(result.success).toBe(true);
      expect(global.fetch).toHaveBeenCalledWith(
        'https://api.openai.com/v1/models',
        expect.objectContaining({
          method: 'GET',
          headers: {
            Authorization: 'Bearer valid-key',
          },
        })
      );
    });

    // TC-T-03: OpenAI APIキーテスト失敗
    test('TC-T-03: OpenAI APIキーテストが失敗する', async () => {
      // Given: fetch がエラーレスポンスを返す
      global.fetch.mockResolvedValueOnce({
        ok: false,
        json: () => Promise.resolve({ error: { message: 'Invalid API key' } }),
      });

      // When: testApiKey を呼び出す
      const result = await config.testApiKey('openai', 'invalid-key');

      // Then: 失敗が返る
      expect(result.success).toBe(false);
      expect(result.error).toBe('Invalid API key');
    });

    // TC-T-04: Claude APIキーテスト成功
    test('TC-T-04: Claude APIキーテストが成功する', async () => {
      // Given: fetch がOKレスポンスを返す
      global.fetch.mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ content: [] }),
      });

      // When: testApiKey を呼び出す
      const result = await config.testApiKey('claude', 'valid-key');

      // Then: 成功が返る
      expect(result.success).toBe(true);
    });

    // TC-T-05: Gemini APIキーテスト成功
    test('TC-T-05: Gemini APIキーテストが成功する', async () => {
      // Given: fetch がOKレスポンスを返す
      global.fetch.mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ models: [] }),
      });

      // When: testApiKey を呼び出す
      const result = await config.testApiKey('gemini', 'valid-key');

      // Then: 成功が返る
      expect(result.success).toBe(true);
    });
  });

  describe('hasSavedApiKey', () => {
    beforeEach(() => {
      config = require('../src/taskpane/config.js');
    });

    // TC-H-01: 保存済みキーがある場合 true
    test('TC-H-01: 保存済みキーがある場合 true を返す', () => {
      // Given: キーが保存されている
      global.Office.context.roamingSettings._storage['openai_api_key'] = 'saved-key';

      // When: hasSavedApiKey を呼び出す
      const result = config.hasSavedApiKey('openai');

      // Then: true が返る
      expect(result).toBe(true);
    });

    // TC-H-02: 保存済みキーがない場合 false
    test('TC-H-02: 保存済みキーがない場合 false を返す', () => {
      // Given: キーが保存されていない
      // When: hasSavedApiKey を呼び出す
      const result = config.hasSavedApiKey('openai');

      // Then: false が返る
      expect(result).toBe(false);
    });
  });

  describe('hasEnvApiKey', () => {
    // TC-E-01: 環境変数キーがある場合 true
    test('TC-E-01: 環境変数キーがある場合 true を返す', () => {
      // Given: 環境変数にキーが設定されている
      global.__OPENAI_API_KEY__ = 'env-key';
      jest.resetModules();
      config = require('../src/taskpane/config.js');

      // When: hasEnvApiKey を呼び出す
      const result = config.hasEnvApiKey('openai');

      // Then: true が返る
      expect(result).toBe(true);
    });

    // TC-E-02: 環境変数キーがない場合 false
    test('TC-E-02: 環境変数キーがない場合 false を返す', () => {
      // Given: 環境変数にキーが設定されていない
      global.__OPENAI_API_KEY__ = '';
      jest.resetModules();
      config = require('../src/taskpane/config.js');

      // When: hasEnvApiKey を呼び出す
      const result = config.hasEnvApiKey('openai');

      // Then: false が返る
      expect(result).toBe(false);
    });
  });

  describe('restoreApiKeyToInput', () => {
    beforeEach(() => {
      config = require('../src/taskpane/config.js');
    });

    // TC-R-01: 保存済みキーを入力フィールドに復元成功
    test('TC-R-01: 保存済みキーを入力フィールドに復元する', () => {
      // Given: キーが保存されている
      global.Office.context.roamingSettings._storage['openai_api_key'] = 'saved-key';

      // When: restoreApiKeyToInput を呼び出す
      const result = config.restoreApiKeyToInput('openai');

      // Then: 入力フィールドにキーが設定される
      expect(result).toBe(true);
      expect(document.getElementById('api-key-input').value).toBe('saved-key');
    });

    // TC-R-02: 保存済みキーがない場合は復元しない
    test('TC-R-02: 保存済みキーがない場合は false を返す', () => {
      // Given: キーが保存されていない
      // When: restoreApiKeyToInput を呼び出す
      const result = config.restoreApiKeyToInput('openai');

      // Then: false が返る
      expect(result).toBe(false);
    });
  });

  describe('getAllKeyStatus', () => {
    beforeEach(() => {
      config = require('../src/taskpane/config.js');
    });

    // TC-G-01: 全プロバイダーの状態を取得
    test('TC-G-01: 全プロバイダーのキー状態を返す', () => {
      // Given: いくつかのキーが保存されている
      global.Office.context.roamingSettings._storage['openai_api_key'] = 'openai-saved';
      global.Office.context.roamingSettings._storage['claude_api_key'] = 'claude-saved';

      // When: getAllKeyStatus を呼び出す
      const status = config.getAllKeyStatus();

      // Then: 各プロバイダーの状態が返る
      expect(status.openai.hasEnv).toBe(false);
      expect(status.openai.hasSaved).toBe(true);
      expect(status.claude.hasEnv).toBe(false);
      expect(status.claude.hasSaved).toBe(true);
      expect(status.gemini.hasEnv).toBe(false);
      expect(status.gemini.hasSaved).toBe(false);
    });
  });

  describe('ApiKeySource 定数', () => {
    beforeEach(() => {
      config = require('../src/taskpane/config.js');
    });

    // TC-C-01: ApiKeySource 定数が正しく定義されている
    test('TC-C-01: ApiKeySource 定数が正しく定義されている', () => {
      // Given/When: ApiKeySource を参照
      // Then: 正しい値が定義されている
      expect(config.ApiKeySource.ENV).toBe('env');
      expect(config.ApiKeySource.SAVED).toBe('saved');
      expect(config.ApiKeySource.INPUT).toBe('input');
      expect(config.ApiKeySource.NONE).toBe('none');
    });
  });
});

