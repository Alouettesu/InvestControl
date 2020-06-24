function Run() {

  var token = PropertiesService.getScriptProperties().getProperty('InvestToken');
  var account = PropertiesService.getScriptProperties().getProperty('AccountNumber');
  var t = new Tinkoff(token, account, 'INVEST');
  
  var tinkoffOperationsSheet = SpreadsheetApp.getActive().getSheetByName('Тинькофф Сделки');
  var tinkoffClosedOperationsSheet = SpreadsheetApp.getActive.getSheetByName('Тинькофф Закрытые сделки');
  var tinkoffAccountSheet = SpreadsheetApp.getActive().getSheetByName('Тинькофф Счёт');
  
  spreadsheetFill = new SpreadsheetFill(t, tinkoffOperationsSheet, tinkoffAccountSheet);
  //spreadsheetFill.resetTime();
  //spreadsheetFill.resetPage();
  
  spreadsheetFill.fillOperations();

  closedOperationsSelector = new ClosedOperationsSelector(tinkoffOperationsSheet, tinkoffClosedOperationsSheet);
  closedOperationsSelector.select();
}

function doGet() {
  Run();
}