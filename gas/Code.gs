function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('現況検査レポート作成ツール')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function createReport(formData) {
  // TODO: デプロイした gen2 Cloud Functions の URL に置き換えてください。
  // 取得方法:
  //   gcloud functions describe generate_excel --gen2 \
  //     --region asia-northeast1 --project dev-addons \
  //     --format='value(serviceConfig.uri)'
  const CLOUD_FUNCTION_URL = 'https://REGION-PROJECT_ID.cloudfunctions.net/generate_excel';

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
      Authorization: 'Bearer ' + ScriptApp.getIdentityToken()
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
  const token = ScriptApp.getIdentityToken();
  const payloadJson = Utilities.newBlob(
    Utilities.base64DecodeWebSafe(token.split('.')[1])
  ).getDataAsString();
  const aud = JSON.parse(payloadJson).aud;
  Logger.log('カスタム audience に登録する値 (aud): ' + aud);
  return aud;
}
