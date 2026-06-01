function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('現況検査レポート作成ツール')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function createReport(formData) {
  // TODO: デプロイしたCloud FunctionsのURLに置き換えてください
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
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(CLOUD_FUNCTION_URL, options);
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
