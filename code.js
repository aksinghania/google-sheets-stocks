function getIndustryPEs() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var codesRange = sheet.getRange('G2:G').getValues();
    var baseUrl = 'https://api.stockedge.com/Api/SecurityDashboardApi/GetCompanyEquityInfo/';
    var queryParams = '/2?lang=en';
    var options = {
      'method': 'GET',
      'headers': {
        'sec-ch-ua': '"Google Chrome";v="125", "Chromium";v="125", "Not.A/Brand";v="24"',
        'Accept': 'application/json, text/plain, */*',
        'Referer': 'https://web.stockedge.com/',
        'DNT': '1',
        'sec-ch-ua-mobile': '?0',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36',
        'sec-ch-ua-platform': '"macOS"',
        'Cookie': 'ARRAffinity=d32b1112ed7cffec331d954886d899961b090c47b7cb3c22cef7e05970977c29'
      }
    };
    
    for (var i = 0; i < codesRange.length; i++) {
      var code = codesRange[i][0];
      if (code) {
        var url = baseUrl + code + queryParams;
        var response = UrlFetchApp.fetch(url, options);
        var data = JSON.parse(response.getContentText());
        var industryPE = data.IndustryPE;
  
        // Set the Industry PE in the corresponding cell in column F
        sheet.getRange(i + 2, 6).setValue(industryPE); // F column is 6th column
      }
    }
  }
  