function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('現況検査レポート作成ツール')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function createReport(formData) {
  // Cloud Function の URL はスクリプトプロパティ CLOUD_FUNCTION_URL から取得する。
  // 設定方法: Apps Script エディタ → プロジェクトの設定 → スクリプト プロパティ
  //   プロパティ名: CLOUD_FUNCTION_URL
  //   値: gen2 関数の URL（下記コマンドで取得）
  //     gcloud functions describe generate_excel --gen2 \
  //       --region asia-northeast1 --project dev-addons \
  //       --format='value(serviceConfig.uri)'
  const CLOUD_FUNCTION_URL =
    PropertiesService.getScriptProperties().getProperty('CLOUD_FUNCTION_URL');
  if (!CLOUD_FUNCTION_URL) {
    throw new Error('スクリプトプロパティ CLOUD_FUNCTION_URL が設定されていません。');
  }

  const payload = {
    floor: formData.floor,
    room_name: formData.room_name,
    x_tilt: formData.x_tilt,
    y_tilt: formData.y_tilt
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    // 非公開（認証必須）の Cloud Function を呼ぶため、実行ユーザー本人の
    // ID トークンを付与する。Cloud Run 側にはこのトークンの aud（= 本 GAS の
    // OAuth クライアント ID）をカスタム audience として登録しておくこと。
    headers: {
      Authorization: 'Bearer ' + getIdentityTokenOrThrow()
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(CLOUD_FUNCTION_URL, options);
    const code = response.getResponseCode();
    if (code === 401 || code === 403) {
      throw new Error('Cloud Function の呼び出しが拒否されました (HTTP ' + code +
        ')。run.invoker 権限とカスタム audience の設定を確認してください。');
    }
    const result = JSON.parse(response.getContentText());

    if (result.error) {
      throw new Error(result.error);
    }
    return result;
  } catch (e) {
    Logger.log(e.toString());
    throw new Error('ファイルの生成に失敗しました: ' + e.toString());
  }
}

/**
 * セットアップ用ユーティリティ。
 * エディタでこの関数を一度実行し、実行ログ（表示 → ログ）に出力される値を
 * Cloud Run サービスのカスタム audience に登録してください。
 *   gcloud run services update generate-excel \
 *     --region=asia-northeast1 --project=dev-addons \
 *     --add-custom-audiences=<ここに出力された値>
 */
function logIdentityTokenAudience() {
  const token = getIdentityTokenOrThrow();
  const payloadJson = Utilities.newBlob(
    Utilities.base64DecodeWebSafe(token.split('.')[1])
  ).getDataAsString();
  const aud = JSON.parse(payloadJson).aud;
  Logger.log('カスタム audience に登録する値 (aud): ' + aud);
  return aud;
}

/**
 * 実行ユーザーの OpenID Connect ID トークンを返す。
 * 取得できない（null）場合は分かりやすいエラーを投げる。
 * getIdentityToken() は openid スコープが未承認だと null を返すため、
 * "Bearer null" での認証失敗や token.split() の TypeError を防ぐ。
 */
function getIdentityTokenOrThrow() {
  const token = ScriptApp.getIdentityToken();
  if (!token) {
    throw new Error('ID トークンを取得できませんでした。OAuth スコープ（openid）が承認されているか確認してください。');
  }
  return token;
}
