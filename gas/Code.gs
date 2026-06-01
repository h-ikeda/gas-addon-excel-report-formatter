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
    y_tilt: formData.y_tilt,
    // Excel 雛形は Google Drive 上の社外秘フォーマットを実行ユーザー権限で
    // 読み取り、Base64 で Cloud Function に渡す。
    template: fetchTemplateBase64_()
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
 * Google Drive 上の Excel 雛形を読み取り、Base64 文字列で返す。
 *
 * 雛形は社外で決められた社外秘フォーマットのため、リポジトリやデプロイ成果物
 * には同梱せず、Drive に置いたファイルを常に参照する。フォーマットが新しい
 * ファイルとして差し替えられても追従できるよう、保存しておいた
 *   - TEMPLATE_FOLDER_ID : 雛形が置かれているフォルダの ID
 *   - TEMPLATE_FILE_NAME : 雛形のファイル名
 * を使ってフォルダ内を検索し、同名ファイルが複数ある場合は最終更新日時が
 * 最も新しいものを採用する。
 *
 * いずれもスクリプトプロパティ（全利用者共通）に保存されており、初回のみ
 * 画面の「雛形を設定」ダイアログ（Google Picker）から選択する。
 */
function fetchTemplateBase64_() {
  const status = getTemplateStatus();
  if (!status.configured) {
    throw new Error('Excel 雛形が未設定です。画面の「雛形を設定」から、Google Drive 上の雛形ファイルを選択してください。');
  }

  const folder = DriveApp.getFolderById(status.folderId);
  const files = folder.getFilesByName(status.fileName);
  let latest = null;
  while (files.hasNext()) {
    const file = files.next();
    // ゴミ箱内の同名ファイルは対象外にする（getFilesByName はゴミ箱の
    // ファイルも返し得るため、更新日時が新しくても採用しない）。
    if (file.isTrashed()) {
      continue;
    }
    if (!latest || file.getLastUpdated() > latest.getLastUpdated()) {
      latest = file;
    }
  }
  if (!latest) {
    throw new Error('指定フォルダ内に雛形ファイル「' + status.fileName +
      '」が見つかりませんでした。フォーマットが差し替えられた可能性があります。「雛形を設定」から選び直してください。');
  }

  return Utilities.base64Encode(latest.getBlob().getBytes());
}

/**
 * 現在保存されている雛形設定の状態を返す。UI の初期表示で使う。
 * @return {{configured: boolean, fileName: string, folderId: string}}
 */
function getTemplateStatus() {
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty('TEMPLATE_FOLDER_ID');
  const fileName = props.getProperty('TEMPLATE_FILE_NAME');
  return {
    configured: !!(folderId && fileName),
    fileName: fileName || '',
    folderId: folderId || ''
  };
}

/**
 * Google Picker で選択された雛形ファイルの ID を受け取り、その親フォルダ ID と
 * ファイル名をスクリプトプロパティ（全利用者共通）に保存する。
 * @param {string} fileId Picker から渡されたファイル ID
 * @return {{fileName: string, folderId: string}}
 */
function saveTemplateSelection(fileId) {
  if (!fileId) {
    throw new Error('ファイルが選択されていません。');
  }
  const file = DriveApp.getFileById(fileId);
  // Google スプレッドシート等を選ぶと getBlob() が PDF を返してしまい、
  // バックエンドの openpyxl が読み込みに失敗する。ネイティブ .xlsx だけを許可する。
  if (file.getMimeType() !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
    throw new Error('Google スプレッドシート等の形式はサポートされていません。Excel 形式 (.xlsx) のファイルを選択してください。');
  }
  const parents = file.getParents();
  if (!parents.hasNext()) {
    throw new Error('選択したファイルの親フォルダを特定できませんでした。マイドライブ直下ではなくフォルダ内に雛形を置いてください。');
  }
  const folderId = parents.next().getId();
  const fileName = file.getName();

  const props = PropertiesService.getScriptProperties();
  props.setProperty('TEMPLATE_FOLDER_ID', folderId);
  props.setProperty('TEMPLATE_FILE_NAME', fileName);

  return { fileName: fileName, folderId: folderId };
}

/**
 * Google Picker をクライアント側で初期化するために必要な情報を返す。
 *   - token  : 実行ユーザーの OAuth トークン（Drive 閲覧用）
 *   - apiKey : Picker 用の API キー（スクリプトプロパティ PICKER_API_KEY）
 *   - appId  : Cloud プロジェクト番号（スクリプトプロパティ PICKER_APP_ID）
 * API キー・App ID は管理者が一度だけスクリプトプロパティに登録する。
 * @return {{token: string, apiKey: string, appId: string}}
 */
function getPickerConfig() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('PICKER_API_KEY');
  const appId = props.getProperty('PICKER_APP_ID');
  if (!apiKey || !appId) {
    throw new Error('Picker の設定（スクリプトプロパティ PICKER_API_KEY / PICKER_APP_ID）が未登録です。README のセットアップ手順を確認してください。');
  }
  return {
    token: ScriptApp.getOAuthToken(),
    apiKey: apiKey,
    appId: appId
  };
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
