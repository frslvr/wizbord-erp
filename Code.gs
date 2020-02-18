var projectDescCol = 5; //  //TODO: relat
var splitter = " | ";
var sheetNameProjects = "Projects";
var sheetNameClients = "Clients";
var sheetNameDeals = "Deals";
var sheetNameVendors = "Vendors";
var sheetNameForms = "Form Data";

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
      .addItem('Refresh All', 'menuRefreshAll')
      .addSeparator()
      .addItem('Refresh Hubspot', 'menuRefreshHubspot')
      .addItem('Refresh Trello', 'refreshTrello')
      .addItem('Refresh Folder Structure', 'updateFolderStructure')
      .addItem('Refresh Vendors', 'refreshVendors')  
      .addSeparator()
      .addItem('Process Gmail', 'processGmail')
      .addItem('Copy Folder struct to Gmail', 'processGmail') // TODO  
      .addItem('Init Hubspot', 'menuInitHubspot')
      .addToUi();
  refresh();
  refreshTrello();
}

function menuRefreshAll()
{
  menuRefreshHubspot();
  refreshTrello();
  processGmail();
  
  Utilities.sleep(3000);
  updateFolderStructure();
  refreshVendors();
}

function myonEdit(e) {
  var aSheet = SpreadsheetApp.getActiveSheet();
  var aCell = aSheet.getActiveCell();
  var aColumn = aCell.getColumn();
  var aRow = aCell.getRow();
  var aValue = aCell.getValue().toString();
  var aSubs = aValue.split(splitter);
  var aSS = SpreadsheetApp.getActiveSpreadsheet();

  if (
//    aSheet.getName() == sheetNameProjects && 
    aCell.getDataValidation() != null) 
  {
    // New Card
    if(aValue.indexOf("[Create new project Trello card]") >= 0) 
    {
      var ret = newCard({name: aSheet.getRange(aRow, projectDescCol, 1, 1).getValue()});
      refreshTrello();
      aCell.setValue("=HYPERLINK(VLOOKUP(\""+ret+"\",'"+trelloTables.projectCards.sheet+"'!B2:F, 5, FALSE), \""+ret+"\")");
    }

    if(aValue.indexOf("[Create new vendor Trello card]") >= 0) 
    {
      // TODO
    }

    // Connect Trello Card
    if(aValue.indexOf("[Connect project Trello card]") >= 0) 
    {
      aCell.setValue("=HYPERLINK(VLOOKUP(\""+aSubs[1]+"\",'"+trelloTables.projectCards.sheet+"'!B2:F, 5, FALSE), \""+aSubs[1]+"\")");
    }

    if(aValue.indexOf("[Connect vendor Trello card]") >= 0) 
    {
      aCell.setValue("=HYPERLINK(VLOOKUP(\""+aSubs[1]+"\",'"+trelloTables.vendorCards.sheet+"'!B2:F, 5, FALSE), \""+aSubs[1]+"\")");
    }
    
    if(aValue.indexOf("[Update from CRM]") >= 0) 
    {
      aSubs = Array();
      var dealId = aSheet.getRange(aRow, aColumn-1).getValue();
      var aMatrix = aSheet.getRange(sheetNameDeals+"!D2:D").getValues();
      // add proj descr / deal name
      var i = scriptVLOOKUP(dealId, aMatrix) + 1;
      if (i<2) return; //something is wrong
      aSubs.push(aSS.getSheetByName(sheetNameDeals).getRange(i, 5).getValue());
      aSubs.push(dealId);
      
      // continue with the next if
      aValue = "[Hubspot connect]";
      aColumn--;
      aCell = aSheet.getRange(aRow, aColumn);
    }
    
    // Connect Hubspot Deal
    if(aValue.indexOf("[Hubspot connect]") >= 0) 
    {
      var dealId = aSubs[1];
      aCell.setValue("=HYPERLINK(\"https://app.hubspot.com/sales/2687042/deal/"+dealId+"/?interaction=note\","+dealId+")");
      // set client id
      aSheet.getRange(aRow, projectDescCol, 1, 1).setValue(aSubs[0]); // set project description

      // get company details
      // vlookup in clients sheet
      var aRange;
      var aMatrix = aSheet.getRange(sheetNameDeals+"!D2:D").getValues();
      var i = scriptMATCH(dealId, aMatrix) + 1; // +1 since we start from D2

      var companyId = aSS.getSheetByName(sheetNameDeals).getRange(i, 6).getValue();
      aSheet.getRange(aRow, aColumn + 1).setValue("=HYPERLINK(\"https://app.hubspot.com/contacts/2687042/company/"+companyId+"/?interaction=note\","+companyId+")");

      var aMatrix = aSheet.getRange(sheetNameClients+"!A2:A").getValues();
      i = scriptMATCH(companyId, aMatrix) + 1; //+1 since A2

      // update clients sheet
      if (i>1) {
        // request api
        var aCompanyInfo = getHubspotCompany(companyId);
        // updating 2 columns
        aRange = aSS.getSheetByName(sheetNameClients).getRange(i, 1,1,2);
        aRange.setValues(aCompanyInfo);
        // decorating companyId
        aRange = aSS.getSheetByName(sheetNameClients).getRange(i, 1,1,1);
        aRange.setValue("=HYPERLINK(\"https://app.hubspot.com/contacts/2687042/company/"+companyId+"/?interaction=note\","+companyId+")");
      }
      
      // TODO: when someone selected to add a CompanyId number into Clients sheet

    }

    if(aValue.indexOf("[Fetch Client ID]") >= 0) 
    {
      // get deal id from 1 cell on the left
      var dealId = aSheet.getRange(aRow, aColumn-1, 1, 1).getValue();
       // set client id

      var aMatrix = aSheet.getRange(sheetNameDeals+"!D2:D").getValues();
      var i = scriptMATCH(dealId, aMatrix) + 1; // +1 since we start from D2
      var companyId = aSS.getSheetByName(sheetNameDeals).getRange(i, 6).getValue();

      aSheet.getRange(aRow, aColumn, 1, 1).setValue("=HYPERLINK(\"https://app.hubspot.com/contacts/2687042/company/"+companyId+"/?interaction=note\","+companyId+")");
    }

    if(aValue.indexOf("[Generate Code]") >= 0) 
    {
      var c1 = aSheet.getRange(aRow, aColumn-1, 1, 1).getValue();
      var r1 = aSheet.getRange(1, aColumn, aSheet.getDataRange().getLastRow(), 1).getA1Notation();
      var ret = getAbbreviature(c1, 0, r1);
      aCell.setValue(ret);
    }

    if(aValue. indexOf("[Connect Form Data]") >= 0) 
    {
      aCell.setValue("=HYPERLINK(VLOOKUP(\""+aSubs[0]+"\",'"+sheetNameForms+"'!A2:B, 2, FALSE), \""+aSubs[0]+"\")");
    }

    if(aValue.indexOf("[Connect Folder]") >= 0) 
    {
      var parrentFolderName = aSheet.getRange("'Folders'!C1").getValue()+"/";
      var folder = getFolderByPath_(parrentFolderName + aSubs[0]);
      Logger.log(parrentFolderName + aSubs[0]);
      aCell.setValue( "=HYPERLINK(\""+folder.getUrl()+"\",\""+aSubs[0]+"\")");

    }
    
    if(aValue.indexOf("[Folder vs CRM]") >= 0) 
    {
      aSheet.getRange(aRow, aColumn).setValue("Loading...");
      var folder = aSheet.getRange(aRow, aColumn-1).getValue();
      var companies = getHubspotCompanies(
        {
          companyId: null,
          properties: {
            name: { value: [folder] },
            vendor_status: {value: null}
          } 
        }
      );
      Logger.log(companies);

      try {
        aSheet.getRange(aRow, 1, 1, 3).setValues(Array.isArray(companies[0]) ? [companies[0]] : [["💩 not found", null, null]]);

        if (Array.isArray(companies[0])) {
          var companyId = companies[0][0];
          aSheet.getRange(aRow, 1).setValue("=HYPERLINK(\"https://app.hubspot.com/contacts/2687042/company/"+companyId+"/?interaction=note\","+companyId+")");
        }
      } catch(e)
      {
        aSheet.getRange(aRow, aColumn).setValue("Error");
        Logger.log(companies[0][0]);
      }
      aSheet.getRange(aRow, aColumn).setValue(companies.length > 1 ? "💔Duplicate CRM record" : null);
    }    
  } 
  else {
    //do nothing
  }

  return;
}

function refreshVendors()
{
  var aSheet = SpreadsheetApp.getActive().getSheetByName(sheetNameVendors);
  var aRow = 2, aColumn = 1;
  
  // 3 values from CRM
  var aColumns = 6;
  var matrix = aSheet.getRange(aRow, aColumn, aSheet.getDataRange().getLastRow()-(aRow-1), aColumns).getValues();
  var tMatrix = transpose(matrix);
  var companies = getHubspotCompanies(
    {
      companyId: tMatrix[0], 
      properties: { 
        name: { value: null }, 
        vendor_status: { 
          value: ["NEW", "OPEN_DEAL", "UNQUALIFIED", "OPEN"] 
        } 
      } 
    });
  var contacts = getHubspotCompanyContacts(
    {
      companyId: transpose(companies)[0],
      properties: {
        email: { value: null },
        lastname: { value: null },
        firstname: { value: null },
        contact_type: { value: ["PRIMARY"] }
      }
    }
  );

  
  for (var r=0; r<matrix.length; r++)   {
    for (var c=0; c<companies.length; c++)   {
      if (companies[c] && companies[c][0] && (matrix[r][0] == companies[c][0] || (matrix[r][0] == "" && r == matrix.length - 1) )) {
        
        matrix[r] = companies[c].concat(contacts[c]);
        
        // decorate companyId
        var companyId = companies[c][0];
        matrix[r][0] ="=HYPERLINK(\"https://app.hubspot.com/contacts/2687042/company/"+companyId+"/?interaction=note\","+companyId+")";
        companies[c] = null;
        break;
      } 
    }    
    if (r == matrix.length - 1 && companies.join("")!="") {
      var x = companies.join("");
      matrix.push(["","",""]);
    }        
  }
  aSheet.getRange(aRow, aColumn, matrix.length, aColumns).setValues(matrix);
}    

// transposing matrix
function transpose(a) {

  // Calculate the width and height of the Array
  var w = a.length || 0;
  var h = a[0] instanceof Array ? a[0].length : 0;

  // In case it is a zero matrix, no transpose routine needed.
  if(h === 0 || w === 0) { return []; }

  /**
   * @var {Number} i Counter
   * @var {Number} j Counter
   * @var {Array} t Transposed data is stored in this array.
   */
  var i, j, t = [];

  // Loop through every item in the outer array (height)
  for(i=0; i<h; i++) {

    // Insert a new row (array)
    t[i] = [];

    // Loop through every item per item in outer array (width)
    for(j=0; j<w; j++) {

      // Save transposed data.
      t[i][j] = a[j][i];
    }
  }

  return t;
}

function scriptVLOOKUP(str,matrix) {
  var flag = false;
  for(var i=0;i<matrix.length;i++) {
    if(matrix[i][0]==str) {
      flag = true;
      break;
    }
  }
  return flag ? i + 1 : null;
};

function scriptMATCH(str,v) {
  var flag = false;
  for(var i=0;i<v.length;i++) {
    if(v[i][0]==str) {
      flag = true;
      break;
    }
  }
  return flag ? i + 1 : null;
};

function getAbbreviature(str, abrIndex, range)
{
  if (str == "") return "Empty name";

  var abrTemplates = [
    /\b([a-zA-Z]){1,1}/g,
    /\b([a-zA-Z]){1,1}([a-zA-Z]){0,1}/g,
    /\b([a-zA-Z]){1,1}([a-zA-Z]){0,1}([a-zA-Z]){0,1}/g,
    /\b([a-zA-Z]){1,1}([a-zA-Z]){0,1}([a-zA-Z]){0,1}([a-zA-Z]){0,1}/g    
    ];

  var aSheet = SpreadsheetApp.getActiveSheet();
  var aRange = aSheet.getRange(range);
  var matrix = aRange.getValues();

  var ret = str.match(abrTemplates[abrIndex]).join('').toUpperCase();

  var retest;
  do
  {
    retest = false;
    for(var i=0; i<matrix.length; i++)
    {

      if (matrix[i][0] == ret)
      {
        abrIndex++;
        if (abrIndex < abrTemplates.length) 
        {
          // try to keep it 4 letter
          ret = str.match(abrTemplates[abrIndex]).join('').toUpperCase().substr(0, 4 + abrIndex); // check out 4?
        }
        else 
        {
          ret += abrIndex - abrTemplates.length + 1;
        }
        retest = true;
      }
    }
  } while(retest); 
  
  return ret;
}

