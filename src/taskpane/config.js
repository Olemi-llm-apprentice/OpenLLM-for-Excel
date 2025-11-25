/*
 * OpenLLM for Excel - API Key Configuration Module
 * 
 * APIキーの取得優先順位:
 * 1. 環境変数（開発者向け、webpack DefinePlugin経由）
 * 2. Office Settings API（ユーザーが保存済み）
 * 3. UI入力フィールド（フォールバック）
 */

/* global Office */

// 環境変数から注入されるAPIキー（webpack DefinePlugin経由）
// 未定義の場合はundefinedになる
const ENV_KEYS = {
  openai: typeof __OPENAI_API_KEY__ !== 'undefined' ? __OPENAI_API_KEY__ : '',
  claude: typeof __ANTHROPIC_API_KEY__ !== 'undefined' ? __ANTHROPIC_API_KEY__ : '',
  gemini: typeof __GEMINI_API_KEY__ !== 'undefined' ? __GEMINI_API_KEY__ : '',
};

// プロバイダー名の正規化（openai-image → openai, gemini-image → gemini）
function normalizeProvider(provider) {
  if (provider === 'openai-image') return 'openai';
  if (provider === 'gemini-image') return 'gemini';
  return provider;
}

// Office Settings APIからキーを取得
function getSavedApiKey(provider) {
  const normalizedProvider = normalizeProvider(provider);
  try {
    if (Office.context && Office.context.roamingSettings) {
      return Office.context.roamingSettings.get(`${normalizedProvider}_api_key`) || '';
    }
  } catch (e) {
    console.warn('Office Settings API not available:', e);
  }
  return '';
}

// Office Settings APIにキーを保存
export function saveApiKey(provider, key) {
  const normalizedProvider = normalizeProvider(provider);
  return new Promise((resolve, reject) => {
    try {
      if (Office.context && Office.context.roamingSettings) {
        Office.context.roamingSettings.set(`${normalizedProvider}_api_key`, key);
        Office.context.roamingSettings.saveAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log(`API key saved for ${normalizedProvider}`);
            resolve(true);
          } else {
            console.error('Failed to save API key:', result.error);
            reject(result.error);
          }
        });
      } else {
        // Office外での実行時はlocalStorageにフォールバック
        localStorage.setItem(`${normalizedProvider}_api_key`, key);
        console.log(`API key saved to localStorage for ${normalizedProvider}`);
        resolve(true);
      }
    } catch (e) {
      console.error('Error saving API key:', e);
      reject(e);
    }
  });
}

// 保存済みキーを削除
export function deleteApiKey(provider) {
  const normalizedProvider = normalizeProvider(provider);
  return new Promise((resolve, reject) => {
    try {
      if (Office.context && Office.context.roamingSettings) {
        Office.context.roamingSettings.remove(`${normalizedProvider}_api_key`);
        Office.context.roamingSettings.saveAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log(`API key deleted for ${normalizedProvider}`);
            resolve(true);
          } else {
            reject(result.error);
          }
        });
      } else {
        localStorage.removeItem(`${normalizedProvider}_api_key`);
        resolve(true);
      }
    } catch (e) {
      reject(e);
    }
  });
}

// UI入力フィールドからキーを取得
function getInputApiKey() {
  const input = document.getElementById('api-key-input');
  return input ? input.value : '';
}

// APIキーの取得元を示す列挙型
export const ApiKeySource = {
  ENV: 'env',
  SAVED: 'saved',
  INPUT: 'input',
  NONE: 'none',
};

// APIキーを優先順位に従って取得
export function getApiKey(provider) {
  const normalizedProvider = normalizeProvider(provider);
  
  // 1. 環境変数（最優先）
  const envKey = ENV_KEYS[normalizedProvider];
  if (envKey) {
    return { key: envKey, source: ApiKeySource.ENV };
  }
  
  // 2. Office Settings API（保存済み）
  const savedKey = getSavedApiKey(normalizedProvider);
  if (savedKey) {
    return { key: savedKey, source: ApiKeySource.SAVED };
  }
  
  // 3. UI入力フィールド
  const inputKey = getInputApiKey();
  if (inputKey) {
    return { key: inputKey, source: ApiKeySource.INPUT };
  }
  
  return { key: '', source: ApiKeySource.NONE };
}

// APIキーの有効性をテスト（軽量なAPIコール）
export async function testApiKey(provider, key) {
  const normalizedProvider = normalizeProvider(provider);
  
  try {
    switch (normalizedProvider) {
      case 'openai':
        return await testOpenAIKey(key);
      case 'claude':
        return await testClaudeKey(key);
      case 'gemini':
        return await testGeminiKey(key);
      default:
        return { success: false, error: `Unknown provider: ${provider}` };
    }
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// OpenAI APIキーのテスト（modelsエンドポイントを使用）
async function testOpenAIKey(key) {
  const response = await fetch('https://api.openai.com/v1/models', {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${key}`,
    },
  });
  
  if (response.ok) {
    return { success: true };
  } else {
    const data = await response.json();
    return { success: false, error: data.error?.message || 'Invalid API key' };
  }
}

// Claude APIキーのテスト（簡単なメッセージを送信）
async function testClaudeKey(key) {
  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': key,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true',
    },
    body: JSON.stringify({
      model: 'claude-3-haiku-20240307',
      max_tokens: 1,
      messages: [{ role: 'user', content: 'Hi' }],
    }),
  });
  
  if (response.ok) {
    return { success: true };
  } else {
    const data = await response.json();
    return { success: false, error: data.error?.message || 'Invalid API key' };
  }
}

// Gemini APIキーのテスト（modelsエンドポイントを使用）
async function testGeminiKey(key) {
  const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${key}`, {
    method: 'GET',
  });
  
  if (response.ok) {
    return { success: true };
  } else {
    const data = await response.json();
    return { success: false, error: data.error?.message || 'Invalid API key' };
  }
}

// 保存済みキーがあるかチェック
export function hasSavedApiKey(provider) {
  const normalizedProvider = normalizeProvider(provider);
  const savedKey = getSavedApiKey(normalizedProvider);
  return !!savedKey;
}

// 環境変数キーがあるかチェック
export function hasEnvApiKey(provider) {
  const normalizedProvider = normalizeProvider(provider);
  return !!ENV_KEYS[normalizedProvider];
}

// UIの入力フィールドに保存済みキーを復元
export function restoreApiKeyToInput(provider) {
  const normalizedProvider = normalizeProvider(provider);
  const savedKey = getSavedApiKey(normalizedProvider);
  const input = document.getElementById('api-key-input');
  
  if (input && savedKey) {
    input.value = savedKey;
    return true;
  }
  return false;
}

// すべてのプロバイダーのキー状態を取得
export function getAllKeyStatus() {
  return {
    openai: {
      hasEnv: hasEnvApiKey('openai'),
      hasSaved: hasSavedApiKey('openai'),
    },
    claude: {
      hasEnv: hasEnvApiKey('claude'),
      hasSaved: hasSavedApiKey('claude'),
    },
    gemini: {
      hasEnv: hasEnvApiKey('gemini'),
      hasSaved: hasSavedApiKey('gemini'),
    },
  };
}

