function sendMostRecentReconToStakeholderEmail(folderId, recipientEmails, ccrecipientEmails, subject) {
  //var subject = "Most Recent File from Google Drive Folder";
  var body = "Please see the attached recon";

  // Get the folder by ID
  var folder = DriveApp.getFolderById(folderId);
  // Get all files in the folder
  var files = folder.getFiles();
  var mostRecentFile = null;
  var lastModifiedTime = new Date(0); // Initialize with the earliest possible date

  // Loop through each file to find the most recent one
  while (files.hasNext()) {
    var file = files.next();
    var modifiedTime = file.getLastUpdated();
    if (modifiedTime > lastModifiedTime) {
      lastModifiedTime = modifiedTime;
      mostRecentFile = file;
    }
  }

  // Check if a file was found
  if (mostRecentFile) {
    // Get the blob of the most recent file
    var blob = mostRecentFile.getBlob();

    // Convert recipient email arrays to comma-separated strings
    var recipientEmailsString = recipientEmails.join(',');
    var ccrecipientEmailsString = ccrecipientEmails.join(',');

    // Send email with attachment
    GmailApp.sendEmail(recipientEmailsString, subject, body, {attachments: [blob], cc: ccrecipientEmailsString});
  } else {
    Logger.log("No files found in the folder");
  }
}

// Call the function for each recon folder and recipient list
function sendAllReconEmailsToStakeholders() {
  
  function getFormattedMondayDate() {
  var today = new Date();
  var dayOfWeek = today.getDay(); // 0 for Sunday, 1 for Monday, ..., 6 for Saturday
  var mondayDate = new Date(today); // Clone today's date
  mondayDate.setDate(mondayDate.getDate() - (dayOfWeek - 1)); // Go back to Monday of the current week
  var formattedDate = mondayDate.toLocaleDateString('en-US', { year: 'numeric', month: '2-digit', day: '2-digit' });
  return formattedDate.replace(/\//g, '-'); // Format the date as MM-DD-YYYY
  }
  
  var thisMonday = getFormattedMondayDate();
  sendMostRecentReconToStakeholderEmail("1hkMXGCDx3cyEJM0EJf4y6borqhsN23s2", 
    ["rbang@onemedical.com"], 
    ["mhieger@onemedical.com"], 
    "Amazon Weekly Invoice Recon " + thisMonday);
  sendMostRecentReconToStakeholderEmail("17nEqELAr7Ye7NqBnF7jznYsCQabW3uUP", 
    ["pritha.sethi@cdw.com", "usparman@onemedical.com", "lspinelli@onemedical.com"], 
    ["rbang@onemedical.com", "mhieger@onemedical.com"], 
    "CDW Weekly Invoice Recon " + thisMonday);
  sendMostRecentReconToStakeholderEmail("1Z88sZnnGiB1MRPSDDE72qXbuSXC7NFza", 
    ["dchevrette@cmecorp.com", "awood@onemedical.com"], 
    ["rbang@onemedical.com", "mhieger@onemedical.com"], 
    "CME Weekly Invoice Recon " + thisMonday);
  sendMostRecentReconToStakeholderEmail("19ETmWITy3QP4qM-Xu-gQpMjoOihfPpqj", 
    ["awood@onemedical.com", "ftsao@onemedical.com"], 
    ["rbang@onemedical.com", "mhieger@onemedical.com"],
    "CuraScript Weekly Invoice Recon " + thisMonday);
  sendMostRecentReconToStakeholderEmail("1QI9vmFAgjBdmI-MLA2mOnP0FYf5E4N2H", 
    ["BPeterson@imagefirst.com", "mvandegrift@imagefirst.com", "mwilliams@imagefirst.com", "mimercurio@onemedical.com"],
    ["rbang@onemedical.com", "mhieger@onemedical.com"],
    "ImageFirst Weekly Invoice Recon " + thisMonday);
  sendMostRecentReconToStakeholderEmail("1EwQPafmkZFmJjlrQxxWn3SSgpXiQ4nbW", 
    ["Jeffrey.Tiedens@mckesson.com", "awood@onemedical.com", "ftsao@onemedical.com"], 
    ["rbang@onemedical.com", "mhieger@onemedical.com"],
    "McKesson Weekly Invoice Recon " + thisMonday);
  sendMostRecentReconToStakeholderEmail("1kBO665Cnd6uuoMS6UsieojAxcn8_oprZ", 
    ["JThomas@medline.com", "KKass@medline.com", "MMount@medline.com", "ftsao@onemedical.com"],
    ["rbang@onemedical.com", "mhieger@onemedical.com"],
    "Medline Weekly Invoice Recon " + thisMonday);
  sendMostRecentReconToStakeholderEmail("1up93F_dlWE_zELJlnBNrgu1c_-Cx568Y", 
    ["Marc@medprodisposal.com", "dmutka@medprodisposal.com", "nzych@medprodisposal.com", "mimercurio@onemedical.com", "knarayan@onemedical.com"],
    ["rbang@onemedical.com", "mhieger@onemedical.com"],
    "MedPro Weekly Invoice Recon " + thisMonday);
  sendMostRecentReconToStakeholderEmail("13pwmejrAK5aFQQNBI_Vo741-KP2V025T", 
    ["ekelleher@onemedical.com", "ftsao@onemedical.com"],
    ["rbang@onemedical.com", "mhieger@onemedical.com"], 
    "Staples Weekly Invoice Recon " + thisMonday);
  sendMostRecentReconToStakeholderEmail("1MgszjrU4EItWcTITrwGvF_ddClrzReMF", 
    ["awood@onemedical.com", "ftsao@onemedical.com"],
    ["rbang@onemedical.com", "mhieger@onemedical.com"], 
    "VaxServe Weekly Invoice Recon " + thisMonday);
  // Add more function calls for other folders and recipients as needed
}

//save