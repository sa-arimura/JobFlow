function logSetupError(error) {
  const errorName = error && error.name ? error.name : 'Error';
  const errorMessage = error && error.message ? error.message : String(error);
  const errorStack = error && error.stack ? error.stack : 'スタック情報なし';

  console.error(
    'JobFlowセットアップエラー\n名前: %s\nメッセージ: %s\nスタック: %s',
    errorName,
    errorMessage,
    errorStack
  );
}

function showSetupErrorDialog(error) {
  const errorMessage = error && error.message ? error.message : String(error);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'セットアップに失敗しました。エラー: ' + errorMessage,
    'JobFlow 初期セットアップ',
    10
  );
}

function showSetupCompleteDialog(completedItems, warnings) {
  const warningLines = warnings.length > 0
    ? '\n\n警告 ' + warnings.length + '件:\n- ' + warnings.join('\n- ')
    : '\n\n警告 0件';
  const message =
    'セットアップが完了しました。\n\n完了した処理:\n- ' +
    completedItems.join('\n- ') +
    warningLines;

  SpreadsheetApp.getActiveSpreadsheet().toast(
    message,
    'JobFlow 初期セットアップ',
    10
  );
}
