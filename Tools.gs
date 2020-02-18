function tabList() {
  var out = new Array()
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 6 ; i < sheets.length+1 ; i++ ) {
    var n = [sheets[i-1].getName()];
    out[i-6] = n + "!A:A";
    
  }
  return out 
}
