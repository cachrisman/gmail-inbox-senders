/***********************
 * WEB APP ENTRYPOINT
 ***********************/
function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Gmail Inbox Senders")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}