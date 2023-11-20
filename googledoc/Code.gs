function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Gisica QA')
    .addItem('Open Sidebar', 'showSidebar')
    .addToUi();


}
function showSidebar() {
  var sheet = SpreadsheetApp.openById('');


  var dataSheet = sheet.getSheetByName('Parent');

  // Get the data from the sheet
  var data = dataSheet.getDataRange().getValues();


  var htmlTemplate = HtmlService.createTemplateFromFile('Sidebar');
  htmlTemplate.data = data;

  var htmlOutput = htmlTemplate.evaluate()
    .setTitle('Data from Sheet')
    .setWidth(300)
    .setHeight(400);

  DocumentApp.getUi().showSidebar(htmlOutput);
}
function insertDataToDoc(dataItem) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  body.appendParagraph(dataItem);
}
 