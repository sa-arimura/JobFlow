/**
 * onOpen()は、スプレッドシートを開いたときにGASが自動で呼び出す特別な関数です。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌸 JobFlow')
    .addItem('応募情報を登録', 'showApplicationRegistrationSidebar')
    .addItem('応募情報を編集', 'showApplicationEditSidebar')
    .addSeparator()
    .addItem('初期セットアップ', 'setupJobFlow')
    .addToUi();
}
