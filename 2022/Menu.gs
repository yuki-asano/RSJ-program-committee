function onOpen() {
  // colorTheDifference()
  SpreadsheetApp
    .getUi()
    .createMenu('Custom Menu')
    .addItem('Color It', 'colorTheDifference')
    .addItem('セッションシート作成', 'generateSessionSheets')
    .addItem('自動生成シート削除', 'removeGeneratedSheets')
    .addToUi()
}
