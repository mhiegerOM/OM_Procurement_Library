function saveWeekendAttachmentsToDrive() {
  // Get today's date
  var today = new Date();
  today.setHours(0, 0, 0, 0); // Set hours, minutes, seconds, and milliseconds to zero
  
  var twoDaysAgo = new Date();
  twoDaysAgo.setDate(twoDaysAgo.getDate() - 2)
  twoDaysAgo.setHours(0, 0, 0, 0); // Set hours, minutes, seconds, and milliseconds to zero

  // Define an array of objects containing subject lines, folder IDs, and file types
  var filesInfo = [
    { subject: "Coupa Report: Invoice Datamart - Daily", folderId: "1prXPusa1KCHSIseQ6h3w0TSuUd14z3GC"},
    { subject: "Coupa Report: Invoice LI Datamart - Daily", folderId: "1prXPusa1KCHSIseQ6h3w0TSuUd14z3GC"},
    { subject: "Coupa Report: Invoice Datamart - Abbrev", folderId: "1prXPusa1KCHSIseQ6h3w0TSuUd14z3GC"},
    { subject: "Coupa Report: Invoice Pmt Details - Daily", folderId: "1prXPusa1KCHSIseQ6h3w0TSuUd14z3GC"},
    { subject: "Coupa Report: IPP Datamart Daily", folderId: "1prXPusa1KCHSIseQ6h3w0TSuUd14z3GC"},
    { subject: "Coupa Report: New Supplier Req - Daily", folderId: "1prXPusa1KCHSIseQ6h3w0TSuUd14z3GC"},
    { subject: "Coupa Report: Payments - Daily", folderId: "1prXPusa1KCHSIseQ6h3w0TSuUd14z3GC"},
    { subject: "Coupa Report: Pmt Acct Datamart - Daily", folderId: "1prXPusa1KCHSIseQ6h3w0TSuUd14z3GC"},
    { subject: "Coupa Report: PO Datamart - Daily", folderId: "1prXPusa1KCHSIseQ6h3w0TSuUd14z3GC"},
    { subject: "Coupa Report: PO LI Datamart - Daily", folderId: "1prXPusa1KCHSIseQ6h3w0TSuUd14z3GC"},
    { subject: "Coupa Report: Suppliers Datamart - Daily", folderId: "1prXPusa1KCHSIseQ6h3w0TSuUd14z3GC"},
    { subject: "Coupa Report: Supplier Info Datamart - Daily", folderId: "1prXPusa1KCHSIseQ6h3w0TSuUd14z3GC"},
    // Add more objects for additional files and folders
    // Example: { subject: "Subject3", folderId: "FolderID3" }
  ];

  // Get all threads in inbox
  var threads = GmailApp.getInboxThreads();

  // Iterate through each thread
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    // Iterate through each message in the thread
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var date = message.getDate(); // Get the date the message was received
      // Check if the email was received today
      if (date >= twoDaysAgo) {
        var attachments = message.getAttachments();
        // Iterate through each attachment
        for (var k = 0; k < attachments.length; k++) {
          var attachment = attachments[k];
          // Iterate through the filesInfo array to find a match for subject and file type
          for (var l = 0; l < filesInfo.length; l++) {
            var fileInfo = filesInfo[l];
            // Check if the subject matches and attachment type matches
            //if (message.getSubject() === fileInfo.subject && attachment.getContentType() === fileInfo.fileType) {
            // Check if the subject matches  
            if (message.getSubject() === fileInfo.subject) {
              // Save attachment to the corresponding Google Drive folder
              var folder = DriveApp.getFolderById(fileInfo.folderId);
              var blob = attachment.copyBlob();
              // If it's a zip file, unzip and save individual files
              if (attachment.getContentType() === "application/zip") {
                var unzippedFiles = Utilities.unzip(blob);
                for (var m = 0; m < unzippedFiles.length; m++) {
                  var unzippedFile = unzippedFiles[m];
                  folder.createFile(unzippedFile);
                }
              } else {
                folder.createFile(blob);
              }
            }
          }
        }
      }
    }
  }
}
