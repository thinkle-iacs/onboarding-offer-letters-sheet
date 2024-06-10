function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("Onboarding")
    .addItem("authorize", "authorize")
    .addItem("setup", "setup")
    .addToUi();
}
