function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle("Private Web App")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSheetData() {
  try {
    var ss = SpreadsheetApp.openById("1bL_YWhS8NHDYbKBOfpz7yOIvS2Vuz4vKEbDNhXMfrsc"); 
    var sheet = ss.getSheetByName("Submissions");
    
    if (!sheet) {
      return JSON.stringify({ error: "Could not find a tab named exactly 'Submissions'" });
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      return JSON.stringify({ error: "The 'Submissions' tab is completely empty!" });
    }
    
    var data = sheet.getDataRange().getValues();
    return JSON.stringify(data); 
    
  } catch(error) {
    return JSON.stringify({ error: error.message });
  }
}
