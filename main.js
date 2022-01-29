function convertExcel2Sheets(excelFile, filename, arrParents) {
  
    Logger.log(excelFile);
    Logger.log(filename);
    Logger.log(arrParents);
    var parents  = arrParents || []; // check if optional arrParents argument was provided, default to empty array if not
    if ( !parents.isArray ) parents = []; // make sure parents is an array, reset to empty array if not
    
    // Parameters for Drive API Simple Upload request (see https://developers.google.com/drive/web/manage-uploads#simple)
    var uploadParams = {
      method:'post',
      contentType: 'application/vnd.ms-excel', // works for both .xls and .xlsx files
      contentLength: excelFile.getBytes().length,
      headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
      payload: excelFile.getBytes()
    };
    
    // Upload file to Drive root folder and convert to Sheets
    var uploadResponse = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v2/files/?uploadType=media&convert=true', uploadParams);
    
    // Parse upload&convert response data (need this to be able to get id of converted sheet)
    var fileDataResponse = JSON.parse(uploadResponse.getContentText());
    
    // Create payload (body) data for updating converted file's name and parent folder(s)
    var payloadData = {
      title: filename, 
      parents: []
    };
    if ( parents.length ) { // Add provided parent folder(s) id(s) to payloadData, if any
      for ( var i=0; i<parents.length; i++ ) {
        try {
          var folder = DriveApp.getFolderById(parents[i]); // check that this folder id exists in drive and user can write to it
          payloadData.parents.push({id: parents[i]});
        }
        catch(e){} // fail silently if no such folder id exists in Drive
      }
    }
    // Parameters for Drive API File Update request (see https://developers.google.com/drive/v2/reference/files/update)
    var updateParams = {
      method:'put',
      headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
      contentType: 'application/json',
      payload: JSON.stringify(payloadData)
    };
    
    // Update metadata (filename and parent folder(s)) of converted sheet
    UrlFetchApp.fetch('https://www.googleapis.com/drive/v2/files/'+fileDataResponse.id, updateParams);
  
    sheet = SpreadsheetApp.openById(fileDataResponse.id);
    cancelFile(fileDataResponse); //clean Google Drive from trash
    cancelFile(filename); //clean Google Drive from trash
    cancelFile(excelFile); //clean Google Drive from trash
    return sheet;
  
}

function returnSpreedSheetFromEmail(emailsLabel){
    var label = GmailApp.getUserLabelByName(emailsLabel); 
    var thread = label.getThreads(0,1)[0]; // Get first thread in inbox
    var lenght = thread.getMessages().length;
    var message = thread.getMessages()[lenght-1]; // Get first message
    var attachments = message.getAttachments(); // Get attachment of first message 
    var folderTemp = DriveApp.getRootFolder();
    Logger.log(thread.getMessages().length);
  
    
    var fileTemp = folderTemp.createFile(attachments[0]);
    
    var xlsBlob = fileTemp.getBlob(); // Blob source of Excel file for conversion
    var xlsFilename = fileTemp.getName(); // File name to give to converted file; defaults to same as source file
    var txtXlsPos = xlsFilename.search(".xlsx");
    xlsFilename = xlsFilename.slice(0,txtXlsPos); 
    var destFolders = []; // array of IDs of Drive folders to put converted file in; empty array = root folder
    var ss = convertExcel2Sheets(xlsBlob, xlsFilename, destFolders);
    cancelFile(fileTemp.getName());//clean Google Drive from trash
    cancelFile(xlsFilename);//clean Google Drive from trash
    cancelFile(xlsBlob.getName());  //clean Google Drive from trash
    return ss;
  }