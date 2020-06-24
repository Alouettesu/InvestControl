class Tinkoff {
  
  constructor(authToken, account, mode) {
    this.auth = 'Bearer ' + authToken;
    this.account = account;
    this.SANDBOX_URL = 'https://api-invest.tinkoff.ru/openapi/sandbox/';
    this.INVEST_URL =  'https://api-invest.tinkoff.ru/openapi/';
    this.stocks = 'market/stocks';
    this.operations = 'operations';
    this.by_figi = 'market/search/by-figi';
    if (mode == 'SANDBOX')
      this.url = this.SANDBOX_URL;
    else if (mode == 'INVEST')
      this.url = this.INVEST_URL;
    
  }
  
  ApiGet(url, requestData) {
    var options = {
      'method': 'get',
      'headers': requestData
    }
    var responce = UrlFetchApp.fetch(url, options)
    return responce;
  }
  
  FormatDate(date) {
    return Utilities.formatDate(date, "GMT+5", "yyyy-MM-dd'T'HH:mm:ss.'000000+00:00'");
  }
  
  ResponceToObject(responce) {
    return JSON.parse(responce);
  }
  
  getStocks() {
    var url = this.url + this.stocks;
    var requestData = {
      'accept': 'application/json',
      'Authorization': this.auth
    }    
    var responce = this.ApiGet(url, requestData);
    return this.responceToObject(responce);
  }
  
  getInstrumentDescription(figi) {
    var url = this.url + this.by_figi
    + '?figi=' + encodeURIComponent(figi);
    var requestData = {
      'accept': 'application/json',
      'Authorization': this.auth
    }    
    var responce = this.ApiGet(url, requestData);
    return this.ResponceToObject(responce);
  }
  
  getOperations() {
    var lastTime = PropertiesService.getScriptProperties().getProperty('lastGetOperationsTime');
    if (!lastTime)
    {
      lastTime = PropertiesService.getScriptProperties().getProperty('DefaultLastOperationTime');
    }
    var currentTime = this.FormatDate(new Date());
    var url = this.url + this.operations
    + '?from=' + encodeURIComponent(lastTime)
    + '&to=' + encodeURIComponent(currentTime)
    + '&brokerAccountId=' + this.account;
    var requestData = {
      'accept': 'application/json',
      'Authorization': this.auth
    }
    var responce = this.ApiGet(url, requestData);
    PropertiesService.getScriptProperties().setProperty('lastGetOperationsTime', currentTime);
    return this.ResponceToObject(responce);
  }
}