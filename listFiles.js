function listFilesInFolder() {
SpreadsheetApp.getActive().getSheetByName('List').getRange('A2:D').clear();

var MAX_FILES = 100; //use a safe value, don't be greedy
var id = '1rCuhjbx1-FlNjxpnrNyL8BMwA7fCvnbk'; 
var scriptProperties = PropertiesService.getScriptProperties();
var lastExecution = scriptProperties.getProperty('LAST_EXECUTION');
var sheet = SpreadsheetApp.getActiveSheet();
var data;
var url = 'https://docs.google.com/spreadsheets/d/';
if( lastExecution === null )
lastExecution = '';

var continuationToken = scriptProperties.getProperty('IMPORT_ALL_FILES_CONTINUATION_TOKEN');
  var iterator = continuationToken == null ?
  DriveApp.getFolderById(id).getFiles() : DriveApp.continueFileIterator(continuationToken);


 try { 
   for( var i = 0; i < MAX_FILES && iterator.hasNext(); ++i ) {
   var file = iterator.next();
   var dateCreated = formatDate(file.getDateCreated());
    if(dateCreated > lastExecution)
     processFile(file);
    data = [
       i+1,
       file.getName(),
       file.getId(),
       file.getUrl(),
      ];
  

    sheet.appendRow(data);

}
} catch(err) {
 Logger.log(err);
}

if( iterator.hasNext() ) {
 scriptProperties.setProperty('IMPORT_ALL_FILES_CONTINUATION_TOKEN', iterator.getContinuationToken());
 } else { // Finished processing files so delete continuation token
 scriptProperties.deleteProperty('IMPORT_ALL_FILES_CONTINUATION_TOKEN');
 scriptProperties.setProperty('LAST_EXECUTION', formatDate(new Date()));
 }
 }

function formatDate(date) { return Utilities.formatDate(date, "GMT", "yyyy-MM-dd HH:mm:ss"); }

function processFile(file) {
var id = file.getId();
var name = file.getName();
//your processing...
Logger.log(name);

}
