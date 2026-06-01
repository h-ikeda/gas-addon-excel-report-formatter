const { loadGas } = require('./gasEnv');

// Build a base64url payload segment for a fake JWT carrying the given claims.
function jwtWithClaims(claims) {
  const payload = Buffer.from(JSON.stringify(claims)).toString('base64url');
  return `header.${payload}.signature`;
}

// Convenience builders for the Apps Script services used by Code.gs.
function scriptPropsMock(value) {
  return {
    getScriptProperties: () => ({ getProperty: () => value }),
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
    floor: '2',
    room_name: 'LDK',
    x_tilt: '3',
    y_tilt: '5',
  };

  function buildEnv({ url = 'https://cf.example/gen', token = jwtWithClaims({ aud: 'aud-x' }), fetch } = {}) {
    return {
      PropertiesService: scriptPropsMock(url),
      ScriptApp: { getIdentityToken: () => token },
      UrlFetchApp: { fetch },
      Logger: { log: jest.fn() },
    };
  }

  test('throws when CLOUD_FUNCTION_URL script property is not set', () => {
    const ctx = loadGas(buildEnv({ url: null }));

    expect(() => ctx.createReport(formData)).toThrow(
      'スクリプトプロパティ CLOUD_FUNCTION_URL が設定されていません。'
    );
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
    expect(JSON.parse(options.payload)).toEqual(formData);

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
