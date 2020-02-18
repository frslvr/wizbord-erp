/*
 * Note:
 * If you need any paid assistant, please write to waqar@accemy.com
 * We provide Apps Script Development services at very reasonable price.
 */

//ScriptApp.newTrigger("parseEmailMessages").timeBased().everyMinutes(5).create();

// GLOBALS
//Array of file extension which you would like to extract to Drive
var applyTypeFilter = true;
var fileTypesToExtract = ['tif', 'bmp', 'svg', 'docx', 'doc', 'pdf', 'xlsx', 'xls', 'ppt', 'pptx', 'zip', 'txt'];
//Name of the folder in google drive i which files will be put
//var folderName = 'GmailToDrive';
//Name of the label which will be applied after processing the mail message
var labelName = 'GmailToDrive';


var targetFolder = "Management";
var sheetNameFolders = "Folders";

// TODO: !!!!!!!!!!!!!!! save thread number or file names to know avoid marking the whole thread
// TODO: speed optimization

function processGmail()
{
  var labelNames = getSublabels_();

  for(var i in labelNames) {
    GmailToDrive_(labelNames[i]);
    extractURLs(labelNames[i]);
  }
}

function GmailToDrive_(wantedLabel){
  //build query to search emails
  var query = '';
  //filename:jpg OR filename:tif OR filename:gif OR fileName:png OR filename:bmp OR filename:svg'; //'after:'+formattedDate+
  for(var i in fileTypesToExtract){
	query += (query === '' ?('filename:'+fileTypesToExtract[i]) : (' OR filename:'+fileTypesToExtract[i]));
  }
  
  //query = 'in:inbox has:nouserlabels ' + query;
  //query = 'label:GDRIVE ' + query;
  query = '-label:' + labelName + ' label:' + wantedLabel + ' ' + query;
  
  var threads = GmailApp.search(query);
  var label = getGmailLabel_(labelName);
  var parentFolder;
  if(threads.length > 0){
    //parentFolder = getFolder_(wantedLabel);
    parentFolder = getFolderByPath_(wantedLabel);
  }
  var root = DriveApp.getRootFolder();
  for(var i in threads){
    var mesgs = threads[i].getMessages();
	for(var j in mesgs){
      //get attachments
      var attachments = mesgs[j].getAttachments();
      for(var k in attachments){
        var attachment = attachments[k];
        if(applyTypeFilter) {
          var isDefinedType = checkIfDefinedType_(attachment);
          if(!isDefinedType) continue;
        }
    	var attachmentBlob = attachment.copyBlob();
        var file = DriveApp.createFile(attachmentBlob);
        parentFolder.addFile(file);
        root.removeFile(file);
      }
	}
	threads[i].addLabel(label);
  }
}


function getFolderByPath_(path) {
      Logger.log(path);

  var parts = path.split("/");

  if (parts[0] == '') parts.shift(); // Did path start at root, '/'?

  var folder = DriveApp.getRootFolder();
  for (var i = 0; i < parts.length; i++) {
    var result = folder.getFoldersByName(parts[i]);
    if (result.hasNext()) {
      folder = result.next();
    } else {
      folder = folder.createFolder(parts[i]);

      break;
    }
  }
  return folder;
}

function getSublabels_()
{
  var wantedLabel = getGmailLabel_(targetFolder);
  var labels = GmailApp.getUserLabels();
  var labelNames = [];
  for (var i = 0; i < labels.length; i++)
  {
    var name = labels[i].getName();
    if(name.indexOf(targetFolder) == 0) {
      labelNames.push(name);
    }
  }

  return labelNames;
}

//This function will get the parent folder in Google drive
function getFolder_(folderName){
  var folder;
  var fi = DriveApp.getFoldersByName(folderName);
  if(fi.hasNext()){
    folder = fi.next();
  }
  else{
    folder = DriveApp.createFolder(folderName);
  }
  return folder;
}

//getDate n days back
// n must be integer
function getDateNDaysBack_(n){
  n = parseInt(n);
  var today = new Date();
  var dateNDaysBack = new Date(today.valueOf() - n*24*60*60*1000);
  return dateNDaysBack;
}

function getGmailLabel_(name){
  var label = GmailApp.getUserLabelByName(name);
  if(!label){
	label = GmailApp.createLabel(name);
  }
  return label;
}

//this function will check for filextension type.
// and return boolean
function checkIfDefinedType_(attachment){
  var fileName = attachment.getName();
  var temp = fileName.split('.');
  var fileExtension = temp[temp.length-1].toLowerCase();
  if(fileTypesToExtract.indexOf(fileExtension) !== -1) return true;
  else return false;
}

function saveUrls(urls, folderPath)
{
  var file = null;
  var gDocIds = [];
  var root = DriveApp.getRootFolder();
  var parentFolder = getFolderByPath_(folderPath);

  for (var i in urls) {
    var url = urls[i];
    var urlParams = parseURL(url);
    var gDocId = (urlParams && urlParams[1] == "docs.google.com") ? getIdFromUrl(url) : null; //TODO: dropbox
    if (gDocId) 
    {
      if (gDocIds.indexOf(gDocId) < 0) 
      {
        gDocIds.push(gDocId);
        file = copyDocs(gDocId); //TODO: try/catch?
      }
    }
    else // download nongoogle docs files
    {  // TODO: schotcut files with links
      continue;
//      var response = UrlFetchApp.fetch(url);
//      file = DriveApp.createFile(response.getBlob()); //.setName(filename);
    }
    parentFolder.addFile(file);
    root.removeFile(file);
  }
}

function extractURLs(wantedLabel) {
  var label="DownloadLinks";

  var query = '-label:' + labelName + ' label:' + wantedLabel + ' ';
  
  var threads = GmailApp.search(query);
  var label = getGmailLabel_(labelName);
  
  for (var i in threads) { // TODO: finish for messages not only threads
    var messages = threads[i].getMessages();
    
    for (var m in messages) {
      var text = messages[m].getBody();
      if (text) {  
        var urls = text.match(/(http(s)?:\/\/.)(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)/g); 
        if (urls) {
          saveUrls(uniq_fast(urls), wantedLabel);
        }
      }
    }
    threads[i].addLabel(label);
  }
}

function parseURL(url) {
  //                  (host-ish)/(path-ish/)(filename)
  var re = /^https?:\/\/([^\/]+)\/([^?]*\/)?([^\/?]+)/;
  return re.exec(url);
}

function copyDocs(gDocId) {
  var file = DriveApp.getFileById(gDocId);
  return file.makeCopy();
}

function getIdFromUrl(url) 
{ 
  return url.match(/[-\w]{25,}/)[0]; 
}

function uniq_fast(a) {
    var seen = {};
    var out = [];
    var len = a.length;
    var j = 0;
    for(var i = 0; i < len; i++) {
         var item = a[i];
         if(seen[item] !== 1) {
               seen[item] = 1;
               out[j++] = item;
         }
    }
    return out;
}

function getFilenameFromURL(url) {
  //                  (host-ish)/(path-ish/)(filename)
  var re = /^https?:\/\/([^\/]+)\/([^?]*\/)?([^\/?]+)/;
  var match = re.exec(url);
  if (match) {
    return unescape(match[3]);
  }
  return null;
}


function updateFolderStructure()
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameFolders);
  // header is at Folders!B1:C1
  var header = sheet.getRange(1, 2, 1, 2).getValues()[0]; 

  for (var c=0; c<header.length && header[c] != ""; c++) {
    var parrentFolder = getFolderByPath_(header[c]);
    var folders = parrentFolder.getFolders();
    var out = [];

    for (var r=0; folders.hasNext(); r++) {
      var folder = folders.next();
      out.push([folder.getName()]);      
    }
    
    out = out.sort(function (a, b) {
      return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0;
    });

    sheet.getRange(2, c+2, sheet.getMaxRows()-1, out[0].length).clearContent(); 
    sheet.getRange(2, c+2, out.length, out[0].length).setValues(out); 
  }
  

}

