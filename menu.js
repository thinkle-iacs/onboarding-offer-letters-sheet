function onOpen (e) {
  SpreadsheetApp.getUi().createMenu(
    "Hello World"
  ).addItem("authorize","authorize")
  .addItem('setup','setup')
  .addToUi();
}