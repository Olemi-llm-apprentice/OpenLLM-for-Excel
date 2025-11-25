/**
 * モデル利用可能性テスト
 * 
 * 各プロバイダーのAPIを呼び出して、モデルが利用可能かどうかを確認します。
 * 
 * 使用方法:
 *   OPENAI_API_KEY=xxx ANTHROPIC_API_KEY=xxx GEMINI_API_KEY=xxx npm run test:models
 * 
 * 注意: 実際のAPIを呼び出すため、APIキーが必要です。
 */

const OPENAI_MODELS = [
  // GPT-5.1シリーズ (Chat Completions API対応)
  'gpt-5.1',
  // GPT-5シリーズ
  'gpt-5',
  'gpt-5-mini',
  'gpt-5-nano',
  // GPT-4シリーズ
  'gpt-4.1',
  'gpt-4o',
  'gpt-4o-mini',
  'gpt-4-turbo',
];

const CLAUDE_MODELS = [
  'claude-sonnet-4-5-20250929',
  'claude-haiku-4-5-20251001',
  'claude-opus-4-5-20251101',
];

const GEMINI_MODELS = [
  // Gemini 3シリーズ
  'gemini-3-pro-preview',
  // Gemini 2.xシリーズ
  'gemini-2.5-pro',
  'gemini-2.5-flash',
  'gemini-2.5-flash-lite',
  'gemini-2.0-flash',
];

const IMAGE_MODELS = {
  openai: ['gpt-image-1'],
  gemini: ['gemini-3-pro-image-preview'],
};

// テスト用の簡単なプロンプト
const TEST_PROMPT = 'Say "Hello" in one word.';

// GPT-5シリーズかどうかを判定
function isGPT5Series(model) {
  return model.startsWith('gpt-5') || model.startsWith('o3') || model.startsWith('o4');
}

/**
 * OpenAI APIをテスト
 */
async function testOpenAIModel(model, apiKey) {
  const url = 'https://api.openai.com/v1/chat/completions';
  
  // GPT-5シリーズは max_completion_tokens を使用
  const tokenParam = isGPT5Series(model) 
    ? { max_completion_tokens: 10 }
    : { max_tokens: 10 };
  
  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model: model,
        messages: [{ role: 'user', content: TEST_PROMPT }],
        ...tokenParam,
      }),
    });

    const data = await response.json();
    
    if (response.ok) {
      return { success: true, model, provider: 'OpenAI' };
    } else {
      return { 
        success: false, 
        model, 
        provider: 'OpenAI',
        error: data.error?.message || JSON.stringify(data),
        code: data.error?.code
      };
    }
  } catch (error) {
    return { success: false, model, provider: 'OpenAI', error: error.message };
  }
}

/**
 * Claude APIをテスト
 */
async function testClaudeModel(model, apiKey) {
  const url = 'https://api.anthropic.com/v1/messages';
  
  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true',
      },
      body: JSON.stringify({
        model: model,
        messages: [{ role: 'user', content: TEST_PROMPT }],
        max_tokens: 10,
      }),
    });

    const data = await response.json();
    
    if (response.ok) {
      return { success: true, model, provider: 'Claude' };
    } else {
      return { 
        success: false, 
        model, 
        provider: 'Claude',
        error: data.error?.message || JSON.stringify(data),
        code: data.error?.type
      };
    }
  } catch (error) {
    return { success: false, model, provider: 'Claude', error: error.message };
  }
}

/**
 * Gemini APIをテスト
 */
async function testGeminiModel(model, apiKey) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
  
  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        contents: [{ parts: [{ text: TEST_PROMPT }] }],
        generationConfig: { maxOutputTokens: 10 },
      }),
    });

    const data = await response.json();
    
    if (response.ok) {
      return { success: true, model, provider: 'Gemini' };
    } else {
      return { 
        success: false, 
        model, 
        provider: 'Gemini',
        error: data.error?.message || JSON.stringify(data),
        code: data.error?.code
      };
    }
  } catch (error) {
    return { success: false, model, provider: 'Gemini', error: error.message };
  }
}

/**
 * 結果を表示
 */
function printResults(results) {
  console.log('\n' + '='.repeat(80));
  console.log('モデル利用可能性テスト結果');
  console.log('='.repeat(80) + '\n');

  const grouped = {
    OpenAI: results.filter(r => r.provider === 'OpenAI'),
    Claude: results.filter(r => r.provider === 'Claude'),
    Gemini: results.filter(r => r.provider === 'Gemini'),
  };

  for (const [provider, providerResults] of Object.entries(grouped)) {
    console.log(`\n【${provider}】`);
    console.log('-'.repeat(60));
    
    for (const result of providerResults) {
      const status = result.success ? '✅ OK' : '❌ NG';
      console.log(`  ${status} ${result.model}`);
      if (!result.success) {
        console.log(`       エラー: ${result.error}`);
        if (result.code) {
          console.log(`       コード: ${result.code}`);
        }
      }
    }
  }

  // サマリー
  const totalTests = results.length;
  const passedTests = results.filter(r => r.success).length;
  const failedTests = totalTests - passedTests;

  console.log('\n' + '='.repeat(80));
  console.log(`サマリー: ${passedTests}/${totalTests} 成功, ${failedTests} 失敗`);
  console.log('='.repeat(80) + '\n');

  // 失敗したモデルのリスト
  if (failedTests > 0) {
    console.log('❌ 利用不可能なモデル:');
    results.filter(r => !r.success).forEach(r => {
      console.log(`  - ${r.provider}: ${r.model}`);
    });
    console.log('\n');
  }

  return failedTests === 0;
}

/**
 * メイン関数
 */
async function main() {
  const openaiKey = process.env.OPENAI_API_KEY;
  const claudeKey = process.env.ANTHROPIC_API_KEY;
  const geminiKey = process.env.GEMINI_API_KEY || process.env.GOOGLE_GENAI_API_KEY;

  console.log('\n🔍 モデル利用可能性テストを開始します...\n');

  const results = [];
  const skipped = [];

  // OpenAI テスト
  if (openaiKey) {
    console.log('📡 OpenAI APIをテスト中...');
    for (const model of OPENAI_MODELS) {
      process.stdout.write(`  Testing ${model}... `);
      const result = await testOpenAIModel(model, openaiKey);
      results.push(result);
      console.log(result.success ? '✅' : '❌');
      // API制限を避けるため少し待機
      await new Promise(r => setTimeout(r, 500));
    }
  } else {
    console.log('⚠️  OPENAI_API_KEY が設定されていません。OpenAIモデルのテストをスキップします。');
    skipped.push('OpenAI');
  }

  // Claude テスト
  if (claudeKey) {
    console.log('\n📡 Claude APIをテスト中...');
    for (const model of CLAUDE_MODELS) {
      process.stdout.write(`  Testing ${model}... `);
      const result = await testClaudeModel(model, claudeKey);
      results.push(result);
      console.log(result.success ? '✅' : '❌');
      await new Promise(r => setTimeout(r, 500));
    }
  } else {
    console.log('⚠️  ANTHROPIC_API_KEY が設定されていません。Claudeモデルのテストをスキップします。');
    skipped.push('Claude');
  }

  // Gemini テスト
  if (geminiKey) {
    console.log('\n📡 Gemini APIをテスト中...');
    for (const model of GEMINI_MODELS) {
      process.stdout.write(`  Testing ${model}... `);
      const result = await testGeminiModel(model, geminiKey);
      results.push(result);
      console.log(result.success ? '✅' : '❌');
      await new Promise(r => setTimeout(r, 500));
    }
  } else {
    console.log('⚠️  GEMINI_API_KEY が設定されていません。Geminiモデルのテストをスキップします。');
    skipped.push('Gemini');
  }

  // 結果表示
  if (results.length > 0) {
    const allPassed = printResults(results);
    process.exit(allPassed ? 0 : 1);
  } else {
    console.log('\n⚠️  テスト可能なAPIキーが設定されていません。');
    console.log('以下の環境変数を設定してください:');
    console.log('  - OPENAI_API_KEY');
    console.log('  - ANTHROPIC_API_KEY');
    console.log('  - GEMINI_API_KEY');
    process.exit(1);
  }
}

// 実行
main().catch(console.error);

