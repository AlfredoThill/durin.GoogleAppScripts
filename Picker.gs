function showPicker() { //uso este picker para eligir el reporte del moodle
  var html = HtmlService.createHtmlOutputFromFile('Picker.html')
                        .setWidth(600)
                        .setHeight(425)
                        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select File');
}

function getOAuthToken() { return ScriptApp.getOAuthToken() }
function saveVariable(id){ PropertiesService.getUserProperties().setProperty('GLOBAL_AUX', id) }
