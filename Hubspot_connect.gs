/**
 * ###########################################################################
 * # Name: Hubspot Automation                                                #
 * # Description: This script let's you connect to Hubspot CRM and retrieve  #
 * #              its data to populate a Google Spreadsheet.                 #
 * # Date: March 11th, 2018                                                  #
 * # Author: Alexis Bedoret                                                  #
 * # Detail of the turorial: https://goo.gl/64hQZb                           #
 * ###########################################################################
 */
// We’ll need to add a library to authenticate using OAuth2:

//Click on the menu item Resources > Libraries…
//In the Find a Library text box, enter the script ID 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF and click the Select button. This is the easiest way to add the library to our project, but check out this article if you want to add the code manually into the project.
//Choose a version in the dropdown box (usually best to pick the latest version).
//Click the Save button.

/**
 * ###########################################################################
 * # ----------------------------------------------------------------------- #
 * # ------------------------------- CONFIG -------------------------------- #
 * # ----------------------------------------------------------------------- #
 * ###########################################################################
 */

/**
 * Fill in the following variables
 */
var CLIENT_ID = '';
var CLIENT_SECRET = '';
//TODO: move to gitignore
var SCOPE = 'contacts';
var AUTH_URL = "https://app.hubspot.com/oauth/authorize";
var TOKEN_URL = "https://api.hubapi.com/oauth/v1/token";
var API_URL = "https://api.hubapi.com";

/**
 * Create the following sheets in your spreadsheet
 * "Stages"
 * "Deals"
 */
var sheetNameStages = "Stages";
var sheetNameDeals = "Deals";
var sheetNameLogSources = "Log: Sources";
var sheetNameLogStages = "Log: Stages";

var hubspotTables = {
  stages: {sheet: "Stages", props: ["stageId","label"]},
  deals: {sheet: "Deals", props: ["stageId","source", "amount", "dealId", "dealname", "associatedCompanyIds"]},
  contacts: {sheet: "Companies", props: ["id", "fullName", "username"]}
};


/**
 * ###########################################################################
 * # ----------------------------------------------------------------------- #
 * # --------------------------- AUTHENTICATION ---------------------------- #
 * # ----------------------------------------------------------------------- #
 * ###########################################################################
 */

/**
 * Authorizes and makes a request to get the deals from Hubspot.
 */
function  getOAuth2Access() {
  var service = getService();
  if (service.hasAccess()) {
    // ... do whatever ...
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s',
        authorizationUrl);
    menuInitHubspot();
  }
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  getService().reset();
}

/**
 * Configures the service.
 */
function getService() {
  return OAuth2.createService('hubspot')
      // Set the endpoint URLs.
      .setTokenUrl(TOKEN_URL)
      .setAuthorizationBaseUrl(AUTH_URL)

      // Set the client ID and secret.
      .setClientId(CLIENT_ID)
      .setClientSecret(CLIENT_SECRET)

      // Set the name of the callback function in the script referenced
      // above that should be invoked to complete the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())
      .setScope(SCOPE);
}

/**
 * Handles the OAuth2 callback.
 */
function authCallback(request) {
  var service = getService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success!');
  } else {
    return HtmlService.createHtmlOutput('Denied.');
  }
}

/**
 * Logs the redict URI to register.
 */
function logRedirectUri() {
  Logger.log(getService().getRedirectUri());
}



/**
 * ###########################################################################
 * # ----------------------------------------------------------------------- #
 * # ------------------------------- GET DATA ------------------------------ #
 * # ----------------------------------------------------------------------- #
 * ###########################################################################
 */

/**
 * Get the different stages in your Hubspot pipeline
 * API & Documentation URL: https://developers.hubspot.com/docs/methods/deal-pipelines/get-deal-pipeline
 */
function getStages() {
  // Prepare authentication to Hubspot
  var service = getService();
  var headers = {headers: {'Authorization': 'Bearer ' + service.getAccessToken()}};  //, muteHttpExceptions: true};
  
  // API request
  var url = API_URL + "/deals/v1/pipelines/default"; // :pipelineId: = default 
  //var url = API_URL + "/crm-pipelines/v1/pipelines/deals";
  var response = UrlFetchApp.fetch(url, headers);
  var result = JSON.parse(response.getContentText());
  
  // Let's sort the stages by displayOrder
  result.stages.sort(function(a,b) {
    return a.displayOrder-b.displayOrder;
  });
  
  // Let's put all the used stages (id & label) in an array
  var stages = Array();
  result.stages.forEach(function(stage) {
    stages.push([stage.stageId,stage.label]);  
  });
  
  return stages;
}

/**
 * Get the deals from your Hubspot pipeline
 * API & Documentation URL: https://developers.hubspot.com/docs/methods/deals/get-all-deals
 */
function getDeals() {
  // Prepare authentication to Hubspot
  var service = getService();
  var headers = {headers: {'Authorization': 'Bearer ' + service.getAccessToken()}};
  
  // Prepare pagination
  // Hubspot lets you take max 250 deals per request. 
  // We need to make multiple request until we get all the deals.
  var keep_going = true;
  var offset = 0;
  var deals = Array();

  while(keep_going)
  {
    // We'll take three properties from the deals: the source, the stage & the amount of the deal
    var url = API_URL + "/deals/v1/deal/paged?properties=dealstage&properties=source&properties=amount&properties=dealname&includeAssociations=true&limit=250&offset="+offset;
    var response = UrlFetchApp.fetch(url, headers);
    var result = JSON.parse(response.getContentText());
    
    // Are there any more results, should we stop the pagination ?
    keep_going = result.hasMore;
    offset = result.offset;
    
    // For each deal, we take the stageId, source & amount
    result.deals.forEach(function(deal) {
      var stageId = getProp(getProp(deal.properties, "dealstage"), "value") || "unknown";
      var source = getProp(getProp(deal.properties, "source"), "value") || "unknown";
      var amount = getProp(getProp(deal.properties, "amount"), "value") || "unknown";
      var dealId = getProp(deal, "dealId");
      var dealname = getProp(getProp(deal.properties, "dealname"), "value") || "unknown";
      var companyId = getProp(getProp(deal, "associations"), "associatedCompanyIds").join(",") || "unknown";
      
      for (var i in hubspotTables) {
        
      }
      
      deals.push([stageId,source,amount,dealId,dealname,companyId]);
      
    });
  }
  
  return deals;
}

function getHubspotCompany(companyId)
{
  // Prepare authentication to Hubspot
  var service = getService();
  var headers = {headers: {'Authorization': 'Bearer ' + service.getAccessToken()}};

  var line = Array();  
  var items = Array();

  var url = API_URL + "/companies/v2/companies/"+companyId+"?";
  var response = UrlFetchApp.fetch(url, headers);
  var result = JSON.parse(response.getContentText()); 
  line.push(result.companyId);
  line.push(getProp(getProp(result.properties, "name"), "value") || "unknown");
  
  items.push(line);    
  
  return items;
}

function test2_getHubspotCompanies()
{
  var companies = getHubspotCompanies({properties: {name: { value: ["Masterdata"] } } });
}

function test_getHubspotCompanies()
{
  var x = getHubspotCompanies( {
    companyId: null,
    properties: {
      name: { 
        value: ["Sannacode", "Rozdoum"]
      },
      vendor_status: { 
        value: "NEW"
      }
    } 
  });

  /* var x = getHubspotCompanies( {
    companyId: 220837981,
    properties: {
      name: { 
        value: ["Sannacode", "Rozdoum"]
      },
      vendor_status: { 
        value: "NEW"
      }
    } 
  });*/
  
  return x;
}

function myIsArray(arr)
{
  return Object.prototype.toString.call(arr) == '[object Array]' && arr.constructor === Array && Array.isArray(arr) ;
}


function getHubspotCompanies(filter)
{
  var service = getService();
  var headers = {headers: {'Authorization': 'Bearer ' + service.getAccessToken()}};
  var keep_going = true;
  var offset = 0;
  var items = [];

  while(keep_going)
  {
    var url = API_URL + "/companies/v2/companies/paged?limit=250&offset="+offset;
    for (var f in filter.properties) {
      url += "&properties="+f;
    }
    var response = UrlFetchApp.fetch(url, headers);
    var result = JSON.parse(response.getContentText()); 
    
    // Are there any more results, should we stop the pagination ?
    keep_going = result["has-more"];
    offset = result.offset;

    for (var c in result.companies) {
      var comp = result.companies[c];
      var line = [];

      // check filter      
      iter(filter, comp, comp, line);
      
      if (getProp(comp, "select")) {
        items.push(line);
      }
    }
  }
  
  return items;
}

function iter(f, v, c, r) {
  Object.keys(f).forEach(function (k) {
    if (f[k] && !myIsArray(f[k]) && f[k].constructor === Object) {
      iter(f[k], getProp(v, k), c, r);
      return;
    }
    r.push(getProp(v, k));
    if (!getProp(v, k) || !f[k]) return;
    if (!Array.isArray(f[k])) f[k] = [f[k]];
    for (var i=0; i<f[k].length; i++) {
      if (f[k][i]) {
        if ((v[k]+"").toLowerCase() == (f[k][i]+"").toLowerCase()) {
          c.select = true;
        }
        else if ((v[k]+"").toLowerCase().indexOf((f[k][i]+"").toLowerCase()) >= 0) {
          c.contains = true;
        }
      }
    }
  });
}

function test_getHubspotCompanyContacts()
{
  var x = getHubspotCompanyContacts( 
    {
      companyId: [477525766, 
                  064906008,
                  793504662,
                  794195888
                 ],
      properties: {
        email: { value: null },
        lastname: { value: null },
        firstname: { value: null },
        contact_type: { value: ["PRIMARY"] }
      }
    }
  );
  
  return x;
}


function getHubspotCompanyContacts(filter)
{
  var query = "compannyContacts";
  var service = getService();
  var headers = {headers: {'Authorization': 'Bearer ' + service.getAccessToken()}};
  var items = [];
  const table = {
    compannyContacts: { root: "/companies/v2/companies/", suffix: "/contacts?", offsetQuery: "count=20&vidOffset=", 
                       hasMore: "hasMore", offset: "vidOffset", result: "contacts" },
    contact: { root: "/contacts/v1/contact/vid/", suffix: "/profile?propertyMode=value_only&showListMemberships=false"}
  };
  var sets = table[query];

  var comIds = filter.companyId;
  
  for (var i=0; i<comIds.length; i++) {
    var line = ["","",""];
    var offset = 0;
    var keep_going = true;
    while(keep_going)
    {
      
      var url = API_URL + sets.root + comIds[i] + sets.suffix + sets.offsetQuery + offset;
      if (getProp(filter, "properties"))
        for (var f in filter.properties) {
          url += "&properties="+f;
        }
      var response = UrlFetchApp.fetch(url, headers);
      var result = JSON.parse(response.getContentText()); 
      
      // Are there any more results, should we stop the pagination ?
      keep_going = result[sets.hasMore];
      offset = result[sets.offset];
      
      for (var c in result[sets.result]) {
        var r = result[sets.result][c];
        var sets2 = table["contact"];

        var url2 = API_URL + sets2.root + getProp(r, "vid") + sets2.suffix;
        if (getProp(filter, "properties"))
          for (var f in filter.properties) {
            url2 += "&properties="+f;
          }
        
        Utilities.sleep(100); // secondly limit at hubspot
        var response2 = UrlFetchApp.fetch(url2, headers);
        var result2 = JSON.parse(response2.getContentText()); 

        var temp = [];
        line = ["","",""];
        iter({ properties: filter.properties}, result2, result2, temp);
        if (getProp(result2, "select") || getProp(result2, "contains")) {
          line = temp.slice(0, 3);
          break;
        }
      }
    }
    items.push(line);
  }
  
  return items;
}


function getProp(prop, name)
{
  return prop ? prop.hasOwnProperty(name) ? prop[name] : null : null;
}

function getPropWhere(prop, name, key, value)
{ 
  if (Array.isArray(prop))
    for (var i in prop) {
      if (getProp(getProp(prop, i), key) == value) 
        return getProp(getProp(prop, i), name);
    }
  
}


/**
* ###########################################################################
* # ----------------------------------------------------------------------- #
* # -------------------------- WRITE TO SPREADSHEET ----------------------- #
* # ----------------------------------------------------------------------- #
* ###########################################################################
*/

/**
 * Print the different stages in your pipeline to the spreadsheet
 */
function writeStages(stages) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetNameStages);
  
  // Let's put some headers and add the stages to our table
  var matrix = Array(["StageID","StageName"]);
  matrix = matrix.concat(stages);
  
  // Writing the table to the spreadsheet
  var range = sheet.getRange(1,1,sheet.getMaxRows(),matrix[0].length).clear();
  range = sheet.getRange(1,1,matrix.length,matrix[0].length);
  range.setValues(matrix);
}

/**
 * Print the different deals that are in your pipeline to the spreadsheet
 */
function writeDeals(deals) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetNameDeals);
  
  // Let's put some headers and add the deals to our table
  var matrix = Array(["StageID","Source", "Amount", "DealId", "DealName", "CompanyId"]);
  matrix = matrix.concat(deals);
  
  // Writing the table to the spreadsheet
  var range = sheet.getRange(1,1,sheet.getMaxRows(),matrix[0].length).clear();
  range = sheet.getRange(1,1,matrix.length,matrix[0].length);
  range.setValues(matrix);
}

function writeHubspotX(x, param) {
   
  var header = hubspotTables[x].props;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getActiveSheet();
  var sheet = ss.getSheetByName(hubspotTables[x].sheet);
 
  var matrix = Array(header);
  matrix = matrix.concat(param);
  
  // Writing the table to the spreadsheet
  var range = sheet.getRange(1,1,sheet.getMaxRows(),matrix[0].length).clear();
  range = sheet.getRange(1,1,matrix.length,matrix[0].length);
  range.setValues(matrix);
}



/**
* ###########################################################################
* # ----------------------------------------------------------------------- #
* # -------------------------------- ROUTINE ------------------------------ #
* # ----------------------------------------------------------------------- #
* ###########################################################################
*/

/**
 * This function will update the spreadsheet. This function should be called
 * every hour or so with the Project Triggers.
 */
function menuRefreshHubspot() {
  var service = getService();
  
  if (service.hasAccess()) {
    var stages = getStages();
    writeStages(stages);
  
    var deals = getDeals();
    writeDeals(deals);
    
//    var companies = getCompanies(deals);
//    writeHubspotX("companies", companies);
    
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s',
        authorizationUrl);
  }
}

function menuInitHubspot() {
  var service = getService();
  
  var authorizationUrl = service.getAuthorizationUrl();
  showAlert('1. Open the following URL \n \
  2. Select User account (not development) \n 3. Re-run the script \n\n' + authorizationUrl);
}

function showAlert(str) {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     str,
      ui.ButtonSet.OK);
}

/**
 * This function will log the amount of leads per stage over time
 * and print it into the sheet "Log: Stages"
 * It should be called once a day with a Project Trigger
 */
function logStages() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Stages: Count");
  var getRange = sheet.getRange("B2:B12");
  var row = getRange.getValues();
  row.unshift(new Date);
  var matrix = [row];
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Log: Stages");
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
    
    // Writing at the end of the spreadsheet
  var setRange = sheet.getRange(lastRow+1,1,1,row.length);
  setRange.setValues(matrix);
}

/**
 * This function will log the amount of leads per source over time
 * and print it into the sheet "Log: Sources"
 * It should be called once a day with a Project Trigger
 */
function logSources() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sources: Count & Conversion Rates");
  var getRange = sheet.getRange("M3:M13");
  var row = getRange.getValues();
  row.unshift(new Date);
  var matrix = [row];
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Log: Sources");
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
    
    // Writing at the end of the spreadsheet
  var setRange = sheet.getRange(lastRow+1,1,1,row.length);
  setRange.setValues(matrix);
}
