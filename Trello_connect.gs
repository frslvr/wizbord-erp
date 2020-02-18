// trello variables

var trelloUrl = "https://api.trello.com/1/";
var authentication = "key=6d1cccf9ea1ab75e12bfd1b6986847ed&token=4f0def23314e656bd4597b41357d9b38977436f1882ce0fa68c28903968df7df";
var inboxList = "55465f3437fea46bddfdeb91";
var updatedLabel = "5aee2704dd787975486cc4e8";

var trelloTables = {
  projectCards: {board: "XXXXXXXXXXXXXXXXXX", sheet: "Project Cards", type: "cards", props: ["name", "id", "idList", "idMembers", "due", "url"]},
  projectLists: {board: "XXXXXXXXXXXXXXXXXXXXX", sheet: "Project Lists", type: "lists", props: ["id", "name"]},
  projectMembers: {board: "XXXXXXXXXXXXXXXXXXXX", sheet: "Project Members", type: "members", props: ["id", "fullName", "username"]},
  vendorCards: {board: "YYYYYYYYYYYYYYY", sheet: "Vendor Cards", type: "cards", props: ["name", "id", "idList", "idMembers", "due", "url"]},
  vendorLists: {board: "YYYYYYYYYYYYYYYYY", sheet: "Vendor Lists", type: "lists", props: ["id", "name"]}
};

//called by google docs apps
function refreshTrello() {
  for (var x in trelloTables) {
    refreshTrelloX(x);
  }
}

// x = members or lists or cards
function refreshTrelloX(x) {
  var params = trelloTables[x].props;
  var ss = Array();
  var response = UrlFetchApp.fetch(trelloUrl + "boards/" + trelloTables[x].board + "/" + trelloTables[x].type + "?" + authentication);

  var details = JSON.parse(response.getContentText());
  if(!details) return; // alert?

  for (var k=0; k < details.length; k++) {
    var line = Array();
    for (var p in params) {
      line.push(details[k][params[p]]);
    }
    ss.push(line);
  }
  // output
  writeTrelloX(x, ss);
}

function newCard(response) {
  var payload = {
    name: response.name,
    idList: inboxList
  };

  var url = trelloUrl + 'cards?' + authentication + '&idList=' + inboxList + '&name=' + response.name;

  var resp = UrlFetchApp.fetch(url, {"method" : "POST"});
  var responseJson = JSON.parse(resp.getContentText());
  return responseJson.id;
}

/**
* ###########################################################################
* # ----------------------------------------------------------------------- #
* # -------------------------- WRITE TO SPREADSHEET ----------------------- #
* # ----------------------------------------------------------------------- #
* ###########################################################################
*/

function writeTrelloX(x, param) {
  var header = trelloTables[x].props;
   
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getActiveSheet();
  var sheet = ss.getSheetByName(trelloTables[x].sheet);
 
  var matrix = Array(header);
  matrix = matrix.concat(param);
  
  // Writing the table to the spreadsheet
  var range = sheet.getRange(1,1,sheet.getMaxRows(),matrix[0].length).clear();
  range = sheet.getRange(1,1,matrix.length,matrix[0].length);
  range.setValues(matrix);
}
