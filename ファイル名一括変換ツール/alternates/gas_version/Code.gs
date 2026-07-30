function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('rename_tool')
    .setTitle('ファイル名一括変換ツール')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
