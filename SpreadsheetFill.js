class SpreadsheetFill {
  constructor(ApiApp, OperationsSheet, AccountSheet) {
    this.ApiApp = ApiApp;
    this.SBER_SHEET = 'Сбербанк';
    this.OPERATION = 'Операция';
    this.OPERATION_SELL = 'Продажа';
    this.OPERATION_BUY = 'Покупка';
    this.OPERATIONS_SHEET = OperationsSheet;
    this.ACCOUNT_SHEET = AccountSheet;
    
    this.Columns = Object.freeze({"operationType":1, 
                                   "date":2, 
                                   "instrumentType":3,
                                   "isin": 4,
                                   "instrumentName": 5,
                                   "quantity": 6,
                                   "price": 7,
                                   "payment": 8,
                                   "currency": 9,
                                   "commission": 10
                                  });    
  }
  
  getSberRowsCount() {
    var sheet = SpreadsheetApp.getActive().getSheetByName(this.SBER_SHEET);
    var rangeData = sheet.getDataRange();
    var lastColumn = rangeData.getLastColumn();
    var lastRow = rangeData.getLastRow();
    return lastRow;
  }
  
  getSheetRowsCount(sheet) {
    var rangeData = sheet.getDataRange();
    var lastRow = rangeData.getLastRow();
    return lastRow;
  }
  
  getSheetColumnsCount(sheet) {
    var rangeData = sheet.getDataRange();
    var lastColumn = rangeData.getLastColumn();
    return lastColumn;
  }
  
  getSberRange() {
    var sheet = SpreadsheetApp.getActive().getSheetByName(this.SBER_SHEET);
    var rangeData = sheet.getDataRange();
    var lastColumn = rangeData.getLastColumn();
    var lastRow = rangeData.getLastRow();
    var searchRange = sheet.getRange(1,1, lastRow-1, lastColumn-1);
    return "'" + SBER_SHEET + "'" + "!" + searchRange.getA1Notation();
  }
  selectClosedPositions() {
    var sheet = SpreadsheetApp.getActive().getSheetByName(this.SBER_SHEET);
    var rangeData = sheet.getDataRange();
    var lastColumn = rangeData.getLastColumn();
    var lastRow = rangeData.getLastRow();
    var searchRange = sheet.getRange(2,1, lastRow-1, lastColumn-1);
    
    Logger.log(lastColumn, lastRow);
  }
  
  resetTime() {
    var defaultTime = PropertiesService.getScriptProperties().getProperty('DefaultLastOperationTime');
    PropertiesService.getScriptProperties().setProperty('lastGetOperationsTime', defaultTime);
  }
  
  resetPage() {
    var sheet = this.OPERATIONS_SHEET;
    var lastRow = this.getSheetRowsCount(sheet);
    var lastColumn = this.getSheetColumnsCount(sheet);
    var range = sheet.getRange(2, 1, lastRow, lastColumn);
    range.clear();
    var sheet = this.ACCOUNT_SHEET;
    var lastRow = this.getSheetRowsCount(sheet);
    var lastColumn = this.getSheetColumnsCount(sheet);
    var range = sheet.getRange(2, 1, lastRow, lastColumn);
    range.clear();
  }
  
  checkCommission(operation) {
    return (operation.commission);
  }
  
   ConvertDateFromString(date) {
    return new Date(Date.parse(date));
  }
  
  appendOperation(sheet, operation) {
    var rowToInsert = this.getSheetRowsCount(sheet) + 1;
    var instrument = {payload: {name: "", isin: ""}};
    if (operation.figi)
    {
      instrument = this.ApiApp.getInstrumentDescription(operation.figi);
      Logger.log(instrument);
    }
    sheet.getRange(rowToInsert, this.Columns.operationType).setValue(operation.operationType);
    sheet.getRange(rowToInsert, this.Columns.date).setValue(this.ConvertDateFromString(operation.date));
    sheet.getRange(rowToInsert, this.Columns.instrumentType).setValue(operation.instrumentType);
    sheet.getRange(rowToInsert, this.Columns.isin).setValue(instrument.payload.ticker);
    sheet.getRange(rowToInsert, this.Columns.instrumentName).setValue(instrument.payload.name);
    sheet.getRange(rowToInsert, this.Columns.quantity).setValue(operation.quantity);
    sheet.getRange(rowToInsert, this.Columns.price).setValue(operation.price);
    sheet.getRange(rowToInsert, this.Columns.payment).setValue(operation.payment);
    sheet.getRange(rowToInsert, this.Columns.currency).setValue(operation.currency);
    if (this.checkCommission(operation)) {
      sheet.getRange(rowToInsert, this.Columns.commission).setValue(operation.commission.value);
    }
  }
  
  fillOperations() {
    var operations = this.ApiApp.getOperations();
    
    var sheet = this.OPERATIONS_SHEET;
    var accountSheet = this.ACCOUNT_SHEET;
    var operationsCount = operations.payload.operations.length;
    for (var i = operationsCount - 1; i >= 0; i--) {
      var operation = operations.payload.operations[i];
      if (operation.status != 'Done')
        continue;
      if (operation.operationType == 'PayIn' || operation.operationType == 'PayOut')
      {
        this.appendOperation(accountSheet, operation);
      }
      else
      {
        this.appendOperation(sheet, operation);
      }      
    }
  }
}