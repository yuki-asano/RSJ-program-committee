/**
 * スプレッドシートを開いたときにカスタムメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()

    .addToUi();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("セッション自動生成")
    .addItem("セッションシート自動生成", "generateSessionSheets")
    .addItem("自動生成シート全削除", "removeGeneratedSheets")
    .addItem(
      "GSシート削除",
      "removeGeneratedGSSheets"
    )
    .addItem(
      "ISシート削除",
      "removeGeneratedISSheets"
    )
    .addItem(
      "OSシート削除",
      "removeGeneratedOSSheets"
    )
    .addToUi();
}
