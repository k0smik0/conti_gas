function doGet() {
  var htmlTemplate = HtmlService.createTemplateFromFile('Update');
  var htmlOuputEvaluated = htmlTemplate.evaluate();
  return htmlOuputEvaluated
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setTitle("Update");
}

function updateFromParse() {
  Logger.log("updateFromParse!");
  var results = queryToParse();
  return results;
}
