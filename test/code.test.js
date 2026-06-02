const { loadGas } = require('./gasEnv');

// Build a base64url payload segment for a fake JWT carrying the given claims.
function jwtWithClaims(claims) {
  const payload = Buffer.from(JSON.stringify(claims)).toString('base64url');
  return `header.${payload}.signature`;
}

// Convenience builders for the Apps Script services used by Code.gs.
// Accepts either a single string (legacy: treated as CLOUD_FUNCTION_URL) or a
// map of property name -> value. setProperty mutates the backing map so tests
// can assert on what saveTemplateSelection persisted.
function scriptPropsMock(value) {
  const map = value !== null && typeof value === 'object'
    ? value
    : { CLOUD_FUNCTION_URL: value };
  return {
    getScriptProperties: () => ({
      getProperty: (key) => (key in map && map[key] !== undefined ? map[key] : null),
      setProperty: (key, val) => { map[key] = val; },
    }),
    _map: map,
  };
}

// Minimal DriveApp iterator.
function makeIterator(items) {
  let i = 0;
  return { hasNext: () => i < items.length, next: () => items[i++] };
}

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// A DriveApp mock backed by a list of {name, updated, bytes} file descriptors.
// Each descriptor may also set `mime` (defaults to xlsx) and `trashed`.
function driveAppMock({ files = [{ name: 'template.xlsx', updated: '2026-01-01', bytes: [80, 75, 1, 2], parent: 'folder-1', id: 'file-1' }] } = {}) {
  function toFile(f) {
    return {
      getName: () => f.name,
      getId: () => f.id || 'file-1',
      getMimeType: () => f.mime || XLSX_MIME,
      isTrashed: () => !!f.trashed,
      getLastUpdated: () => new Date(f.updated),
      getBlob: () => ({ getBytes: () => f.bytes }),
      getParents: () => makeIterator(f.parent ? [{ getId: () => f.parent }] : []),
    };
  }
  return {
    getFolderById: jest.fn(() => ({
      getFilesByName: (name) => makeIterator(files.filter((f) => f.name === name).map(toFile)),
    })),
    getFileById: jest.fn((id) => toFile(files.find((f) => (f.id || 'file-1') === id) || files[0])),
  };
}

describe('doGet', () => {
  test('renders the index template with title and viewport meta tag', () => {
    const output = {
      setTitle: jest.fn().mockReturnThis(),
      addMetaTag: jest.fn().mockReturnThis(),
    };
    const HtmlService = {
      createHtmlOutputFromFile: jest.fn().mockReturnValue(output),
    };

    const ctx = loadGas({ HtmlService });
    const result = ctx.doGet();

    expect(HtmlService.createHtmlOutputFromFile).toHaveBeenCalledWith('index');
    expect(output.setTitle).toHaveBeenCalledWith('現況検査レポート作成ツール');
    expect(output.addMetaTag).toHaveBeenCalledWith(
      'viewport',
      'width=device-width, initial-scale=1'
    );
    expect(result).toBe(output);
  });
});

describe('createReport', () => {
  const formData = {
    property_name: 'サンプル物件',
    rooms: [
      {
        floor: '2',
        room_name: 'LDK',
        measurements: { floor_x: { select: '←', diff: '0', distance: '2000' } },
      },
    ],
  };

  function buildEnv({
    url = 'https://cf.example/gen',
    token = jwtWithClaims({ aud: 'aud-x' }),
    fetch,
    drive = driveAppMock(),
    templateFolderId = 'folder-1',
    templateFileName = 'template.xlsx',
  } = {}) {
    return {
      PropertiesService: scriptPropsMock({
        CLOUD_FUNCTION_URL: url,
        TEMPLATE_FOLDER_ID: templateFolderId,
        TEMPLATE_FILE_NAME: templateFileName,
      }),
      ScriptApp: { getIdentityToken: () => token },
      UrlFetchApp: { fetch },
      DriveApp: drive,
      Utilities: { base64Encode: (bytes) => Buffer.from(bytes).toString('base64') },
      Logger: { log: jest.fn() },
    };
  }

  test('throws when CLOUD_FUNCTION_URL script property is not set', () => {
    const ctx = loadGas(buildEnv({ url: null }));

    expect(() => ctx.createReport(formData)).toThrow(
      'スクリプトプロパティ CLOUD_FUNCTION_URL が設定されていません。'
    );
  });

  test('throws when the Excel template has not been configured', () => {
    const ctx = loadGas(buildEnv({ templateFolderId: null, templateFileName: null }));

    expect(() => ctx.createReport(formData)).toThrow(/Excel 雛形が未設定です/);
  });

  test('posts the expected URL, payload and bearer token, returning the parsed result', () => {
    const fetch = jest.fn().mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () =>
        JSON.stringify({ status: 'success', fileName: 'r.xlsx', fileData: 'BASE64' }),
    });
    const ctx = loadGas(buildEnv({ token: 'id-token-abc', fetch }));

    const result = ctx.createReport(formData);

    expect(fetch).toHaveBeenCalledTimes(1);
    const [url, options] = fetch.mock.calls[0];
    expect(url).toBe('https://cf.example/gen');
    expect(options.method).toBe('post');
    expect(options.contentType).toBe('application/json');
    expect(options.muteHttpExceptions).toBe(true);
    expect(options.headers.Authorization).toBe('Bearer id-token-abc');
    // The payload carries the property name and rooms array plus the
    // Base64-encoded template (bytes [80, 75, 1, 2] -> "UEsBAg==") read from Drive.
    expect(JSON.parse(options.payload)).toEqual({
      property_name: formData.property_name,
      rooms: formData.rooms,
      template: 'UEsBAg==',
    });

    expect(result).toEqual({ status: 'success', fileName: 'r.xlsx', fileData: 'BASE64' });
  });

  test.each([401, 403])('throws a permission error when the function returns HTTP %i', (code) => {
    const fetch = jest.fn().mockReturnValue({
      getResponseCode: () => code,
      getContentText: () => '',
    });
    const ctx = loadGas(buildEnv({ fetch }));

    expect(() => ctx.createReport(formData)).toThrow(
      /Cloud Function の呼び出しが拒否されました \(HTTP /
    );
  });

  test('throws when the response body carries an error field', () => {
    const fetch = jest.fn().mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({ error: 'テンプレートが見つかりません' }),
    });
    const ctx = loadGas(buildEnv({ fetch }));

    expect(() => ctx.createReport(formData)).toThrow(/テンプレートが見つかりません/);
  });

  test('wraps lower-level failures with a friendly message and logs them', () => {
    const fetch = jest.fn(() => {
      throw new Error('network down');
    });
    const env = buildEnv({ token: 'tok', fetch });
    const ctx = loadGas(env);

    expect(() => ctx.createReport(formData)).toThrow(/ファイルの生成に失敗しました: .*network down/);
    expect(env.Logger.log).toHaveBeenCalled();
  });
});

describe('getIdentityTokenOrThrow', () => {
  test('returns the token when ScriptApp provides one', () => {
    const ctx = loadGas({ ScriptApp: { getIdentityToken: () => 'the-token' } });
    expect(ctx.getIdentityTokenOrThrow()).toBe('the-token');
  });

  test('throws a descriptive error when no token is available', () => {
    const ctx = loadGas({ ScriptApp: { getIdentityToken: () => null } });
    expect(() => ctx.getIdentityTokenOrThrow()).toThrow(/ID トークンを取得できませんでした/);
  });
});

describe('getTemplateStatus', () => {
  test('reports configured with the saved file name when both props are set', () => {
    const ctx = loadGas({
      PropertiesService: scriptPropsMock({
        TEMPLATE_FOLDER_ID: 'folder-1',
        TEMPLATE_FILE_NAME: '傾斜測定.xlsx',
      }),
    });

    expect(ctx.getTemplateStatus()).toEqual({
      configured: true,
      fileName: '傾斜測定.xlsx',
      folderId: 'folder-1',
    });
  });

  test('reports unconfigured when the properties are missing', () => {
    const ctx = loadGas({ PropertiesService: scriptPropsMock({}) });

    expect(ctx.getTemplateStatus()).toEqual({
      configured: false,
      fileName: '',
      folderId: '',
    });
  });
});

describe('saveTemplateSelection', () => {
  test('persists the parent folder id and file name, returning them', () => {
    const props = scriptPropsMock({});
    const drive = driveAppMock({
      files: [{ name: '雛形.xlsx', updated: '2026-01-01', bytes: [1], parent: 'parent-9', id: 'doc-9' }],
    });
    const ctx = loadGas({ PropertiesService: props, DriveApp: drive });

    const result = ctx.saveTemplateSelection('doc-9');

    expect(result).toEqual({ fileName: '雛形.xlsx', folderId: 'parent-9' });
    expect(props._map.TEMPLATE_FOLDER_ID).toBe('parent-9');
    expect(props._map.TEMPLATE_FILE_NAME).toBe('雛形.xlsx');
  });

  test('throws when no file id is provided', () => {
    const ctx = loadGas({ PropertiesService: scriptPropsMock({}), DriveApp: driveAppMock() });

    expect(() => ctx.saveTemplateSelection('')).toThrow(/ファイルが選択されていません/);
  });

  test('rejects a non-xlsx file such as a native Google Sheet', () => {
    const props = scriptPropsMock({});
    const drive = driveAppMock({
      files: [{ name: 'sheet', updated: '2026-01-01', bytes: [1], parent: 'p', id: 'gs-1', mime: 'application/vnd.google-apps.spreadsheet' }],
    });
    const ctx = loadGas({ PropertiesService: props, DriveApp: drive });

    expect(() => ctx.saveTemplateSelection('gs-1')).toThrow(/Excel 形式 \(\.xlsx\)/);
    // Nothing should have been persisted.
    expect(props._map.TEMPLATE_FOLDER_ID).toBeUndefined();
    expect(props._map.TEMPLATE_FILE_NAME).toBeUndefined();
  });

  test('throws when the selected file has no parent folder', () => {
    const drive = driveAppMock({
      files: [{ name: 'x.xlsx', updated: '2026-01-01', bytes: [1], parent: null, id: 'doc-1' }],
    });
    const ctx = loadGas({ PropertiesService: scriptPropsMock({}), DriveApp: drive });

    expect(() => ctx.saveTemplateSelection('doc-1')).toThrow(/親フォルダを特定できませんでした/);
  });
});

describe('getPickerConfig', () => {
  test('returns the OAuth token plus the configured API key and app id', () => {
    const ctx = loadGas({
      PropertiesService: scriptPropsMock({ PICKER_API_KEY: 'key-1', PICKER_APP_ID: 'app-1' }),
      ScriptApp: { getOAuthToken: () => 'oauth-token' },
    });

    expect(ctx.getPickerConfig()).toEqual({
      token: 'oauth-token',
      apiKey: 'key-1',
      appId: 'app-1',
    });
  });

  test('throws when the Picker properties are not registered', () => {
    const ctx = loadGas({
      PropertiesService: scriptPropsMock({}),
      ScriptApp: { getOAuthToken: () => 'oauth-token' },
    });

    expect(() => ctx.getPickerConfig()).toThrow(/PICKER_API_KEY \/ PICKER_APP_ID/);
  });
});

describe('fetchTemplateBase64_', () => {
  test('returns the base64 of the most recently updated same-named file', () => {
    const drive = driveAppMock({
      files: [
        { name: 'template.xlsx', updated: '2025-01-01', bytes: [1, 1], parent: 'folder-1' },
        { name: 'template.xlsx', updated: '2026-05-01', bytes: [80, 75, 9, 9], parent: 'folder-1' },
      ],
    });
    const ctx = loadGas({
      PropertiesService: scriptPropsMock({
        TEMPLATE_FOLDER_ID: 'folder-1',
        TEMPLATE_FILE_NAME: 'template.xlsx',
      }),
      DriveApp: drive,
      Utilities: { base64Encode: (bytes) => Buffer.from(bytes).toString('base64') },
    });

    // Bytes [80, 75, 9, 9] of the newer file -> "UEsJCQ==".
    expect(ctx.fetchTemplateBase64_()).toBe('UEsJCQ==');
  });

  test('skips trashed files even when they have the newest timestamp', () => {
    const drive = driveAppMock({
      files: [
        { name: 'template.xlsx', updated: '2025-01-01', bytes: [80, 75, 5, 5], parent: 'folder-1' },
        { name: 'template.xlsx', updated: '2026-12-31', bytes: [9, 9], parent: 'folder-1', trashed: true },
      ],
    });
    const ctx = loadGas({
      PropertiesService: scriptPropsMock({
        TEMPLATE_FOLDER_ID: 'folder-1',
        TEMPLATE_FILE_NAME: 'template.xlsx',
      }),
      DriveApp: drive,
      Utilities: { base64Encode: (bytes) => Buffer.from(bytes).toString('base64') },
    });

    // The trashed (newer) file is ignored; the live file's bytes [80,75,5,5] -> "UEsFBQ==".
    expect(ctx.fetchTemplateBase64_()).toBe('UEsFBQ==');
  });

  test('throws when the named file is absent from the folder', () => {
    const drive = driveAppMock({
      files: [{ name: 'other.xlsx', updated: '2026-01-01', bytes: [1], parent: 'folder-1' }],
    });
    const ctx = loadGas({
      PropertiesService: scriptPropsMock({
        TEMPLATE_FOLDER_ID: 'folder-1',
        TEMPLATE_FILE_NAME: 'template.xlsx',
      }),
      DriveApp: drive,
      Utilities: { base64Encode: (bytes) => Buffer.from(bytes).toString('base64') },
    });

    expect(() => ctx.fetchTemplateBase64_()).toThrow(/見つかりませんでした/);
  });
});

describe('logIdentityTokenAudience', () => {
  test('decodes the aud claim from the identity token, logs and returns it', () => {
    const token = jwtWithClaims({ aud: 'my-oauth-client-id' });
    const logger = { log: jest.fn() };

    // Mirror Apps Script's Utilities helpers: decode the base64url segment and
    // expose the bytes as a string.
    const Utilities = {
      base64DecodeWebSafe: (segment) => Buffer.from(segment, 'base64url'),
      newBlob: (bytes) => ({ getDataAsString: () => Buffer.from(bytes).toString('utf8') }),
    };

    const ctx = loadGas({
      ScriptApp: { getIdentityToken: () => token },
      Utilities,
      Logger: logger,
    });

    const aud = ctx.logIdentityTokenAudience();

    expect(aud).toBe('my-oauth-client-id');
    expect(logger.log).toHaveBeenCalledWith(
      expect.stringContaining('my-oauth-client-id')
    );
  });
});
