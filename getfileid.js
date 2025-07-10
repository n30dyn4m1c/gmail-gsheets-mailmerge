function listFilesToSheet() {
  var folderId = '1exkNZ2OdRYJuOpCCWkP75GrrA3SvQjmZ'; // Replace with the ID of the folder containing your files
  var files = DriveApp.getFolderById(folderId).getFiles();
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear(); // Clear existing data
  
  sheet.appendRow(["File Name", "File ID"]); // Header row
  
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var fileId = file.getId();
    
    sheet.appendRow([fileName, fileId]);
  }
}
