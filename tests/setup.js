/**
 * Jest Setup File
 * テスト実行前のグローバル設定
 */

// Office.jsのモック
global.Office = {
  HostType: {
    Excel: 'Excel',
  },
  AsyncResultStatus: {
    Succeeded: 'succeeded',
    Failed: 'failed',
  },
  context: {
    roamingSettings: {
      _storage: {},
      get: jest.fn(function(key) {
        return this._storage[key] || '';
      }),
      set: jest.fn(function(key, value) {
        this._storage[key] = value;
      }),
      remove: jest.fn(function(key) {
        delete this._storage[key];
      }),
      saveAsync: jest.fn(function(callback) {
        callback({ status: global.Office.AsyncResultStatus.Succeeded });
      }),
    },
  },
  onReady: jest.fn((callback) => {
    callback({ host: global.Office.HostType.Excel });
    return Promise.resolve();
  }),
};

// localStorageのモック
const localStorageMock = {
  _storage: {},
  getItem: jest.fn(function(key) {
    return this._storage[key] || null;
  }),
  setItem: jest.fn(function(key, value) {
    this._storage[key] = value;
  }),
  removeItem: jest.fn(function(key) {
    delete this._storage[key];
  }),
  clear: jest.fn(function() {
    this._storage = {};
  }),
};

Object.defineProperty(global, 'localStorage', {
  value: localStorageMock,
});

// fetchのモック
global.fetch = jest.fn();

// Excel のモック
global.Excel = {
  run: jest.fn((callback) => {
    const mockContext = {
      workbook: {
        worksheets: {
          onSelectionChanged: {
            add: jest.fn(),
          },
          getActiveWorksheet: jest.fn(() => ({
            getRange: jest.fn(() => ({
              values: [[]],
              format: {
                borders: {
                  getItem: jest.fn(() => ({
                    style: '',
                    weight: '',
                  })),
                },
                autofitColumns: jest.fn(),
                autofitRows: jest.fn(),
              },
              numberFormat: '',
            })),
          })),
        },
        getSelectedRange: jest.fn(() => ({
          address: 'A1',
          text: [['']],
          load: jest.fn(),
        })),
      },
      sync: jest.fn(() => Promise.resolve()),
    };
    return Promise.resolve(callback(mockContext));
  }),
  BorderLineStyle: {
    continuous: 'continuous',
  },
};

// marked ライブラリのモック
global.marked = {
  parse: jest.fn((text) => text),
};

// consoleのモック（テスト中のログを制御）
global.console = {
  ...console,
  log: jest.fn(),
  warn: jest.fn(),
  error: jest.fn(),
};

// 各テスト前にモックをリセット
beforeEach(() => {
  jest.clearAllMocks();
  global.Office.context.roamingSettings._storage = {};
  localStorageMock._storage = {};
});

