// Function to send interview emails and update the spreadsheet
function sendAndTrackInterviewEmails() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var applicantsSheet = ss.getSheetByName("Applicant");
    var interviewersSheet = ss.getSheetByName("Interviewer");
    var applicantsData = applicantsSheet.getDataRange().getValues();
    var interviewersData = interviewersSheet.getDataRange().getValues();
    
    var keywords = ["JavaScript", "Python", "Machine Learning", "Project Management"]; // Example criteria
    
    // Clear previous highlights and statuses
    applicantsSheet.getRange(2, 1, applicantsSheet.getMaxRows() - 1, applicantsSheet.getMaxColumns()).setBackground(null);
    interviewersSheet.getRange(2, 1, interviewersSheet.getMaxRows() - 1, interviewersSheet.getMaxColumns()).setBackground(null);
    applicantsSheet.getRange(2, 4, applicantsSheet.getMaxRows() - 1, 2).clearContent(); // Assuming columns D and E are used for responses
    interviewersSheet.getRange(2, 3, interviewersSheet.getMaxRows() - 1, 1).clearContent(); // Assuming column C is used for responses
  
    for (var i = 1; i < applicantsData.length; i++) {
      var applicantName = applicantsData[i][0];
      var applicantEmail = applicantsData[i][1];
      var resumeLink = applicantsData[i][2];
      var matchCount = 0;
      var relevantDetails = [];
      
      // Extract text from PDF
      var fileId = getFileIdFromUrl(resumeLink);
      if (fileId) {
        try {
          var text = convertPdfToDocAndExtractText(fileId);
          
          // Log the scraped text to the logger
          Logger.log('Scraped text for ' + applicantName + ': ' + text);
          
          // Scan the text for keywords and extract relevant details
          var sentences = text.split(/[.!?]/); // Split text into sentences
          for (var j = 0; j < sentences.length; j++) {
            for (var k = 0; k < keywords.length; k++) {
              if (sentences[j].includes(keywords[k])) {
                matchCount++;
                relevantDetails.push(sentences[j].trim());
                break; // Stop searching for other keywords in this sentence
              }
            }
          }
          
          // Log the relevant details to the logger
          Logger.log('Relevant details for ' + applicantName + ': ' + relevantDetails.join('. '));
          
        } catch (e) {
          Logger.log('Error processing file with ID ' + fileId + ': ' + e.message);
        }
      } else {
        Logger.log('Invalid file URL: ' + resumeLink);
      }
  
      if (matchCount >= 2) { // Example criteria: at least 2 keyword matches
        var interviewerName = interviewersData[i][0];
        var interviewerEmail = interviewersData[i][1];
        var interviewTime = new Date(); // Set your preferred interview time
        
        // Create Calendar event
        var calendar = CalendarApp.getDefaultCalendar();
        var event = calendar.createEvent('Interview with ' + applicantName,
                                         interviewTime,
                                         new Date(interviewTime.getTime() + 60 * 60 * 1000), // 1-hour interview
                                         {
                                           guests: applicantEmail + ',' + interviewerEmail,
                                           sendInvites: true
                                         });
        
        // Send email to applicant
        var applicantSubject = "Interview Scheduled";
        var applicantBody = "Dear " + applicantName + ",\n\nYou have been scheduled for an interview with " + interviewerName + " on " + interviewTime + ".\n\nPlease confirm your availability by accepting or declining the calendar invite.\n\nBest regards,\nHR Team";
        MailApp.sendEmail(applicantEmail, applicantSubject, applicantBody);
        
        // Send email to interviewer
        var interviewerSubject = "Interview Scheduled";
        var interviewerBody = "Dear " + interviewerName + ",\n\nYou have been scheduled for an interview with " + applicantName + " on " + interviewTime + ".\n\nPlease confirm your availability by accepting or declining the calendar invite.\n\nBest regards,\nHR Team";
        MailApp.sendEmail(interviewerEmail, interviewerSubject, interviewerBody);
        
        // Highlight the row and update the status
        applicantsSheet.getRange(i + 1, 1, 1, applicantsSheet.getMaxColumns()).setBackground("yellow");
        interviewersSheet.getRange(i + 1, 1, 1, interviewersSheet.getMaxColumns()).setBackground("yellow");
        applicantsSheet.getRange(i + 1, 4).setValue('Scheduled');
        interviewersSheet.getRange(i + 1, 3).setValue('Scheduled');
      }
    }
  }
  
  // Function to update the interview status
  function updateInterviewStatus() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var applicantsSheet = ss.getSheetByName("Applicant");
    var interviewersSheet = ss.getSheetByName("Interviewer");
    var applicantsData = applicantsSheet.getDataRange().getValues();
    var interviewersData = interviewersSheet.getDataRange().getValues();
  
    for (var i = 1; i < applicantsData.length; i++) {
      var applicantResponse = applicantsData[i][3];
      var interviewerResponse = interviewersData[i][2];
      
      if (applicantResponse == "Yes" && interviewerResponse == "Yes") {
        applicantsSheet.getRange(i + 1, 5).setValue('Confirmed');
        interviewersSheet.getRange(i + 1, 4).setValue('Confirmed');
      } else if (applicantResponse == "No" || interviewerResponse == "No") {
        applicantsSheet.getRange(i + 1, 5).setValue('Cancelled');
        interviewersSheet.getRange(i + 1, 4).setValue('Cancelled');
      }
    }
  }
  
  // Function to send and track interviews
  function sendAndTrackInterviews() {
    sendAndTrackInterviewEmails();
    updateInterviewStatus();
  }
  
  // Function to extract the file ID from the URL
  function getFileIdFromUrl(url) {
    var fileId = '';
    
    // Check for the format 'https://drive.google.com/open?id=FILE_ID'
    var regex = /(?:https:\/\/drive.google.com\/open\?id=|\/d\/|\/file\/d\/|\/u\/0\/d\/|id=|file\/d\/)([-\w]+)/;
    var matches = url.match(regex);
    if (matches && matches[1]) {
      fileId = matches[1];
    }
    
    return fileId;
  }
  
  // Function to convert PDF to Google Doc and extract plain text
  function convertPdfToDocAndExtractText(fileId) {
    var pdfFile = DriveApp.getFileById(fileId);
    var blob = pdfFile.getBlob();
    
    // Create a new Google Doc
    var doc = DocumentApp.create(pdfFile.getName() + ' - Converted');
    var body = doc.getBody();
    
    // Append PDF content to the Google Doc
    var paragraphs = [];
    try {
      body.appendParagraph(blob.getDataAsString());
    } catch (e) {
      Logger.log('Error converting PDF to Google Doc: ' + e.message);
    }
    
    doc.saveAndClose();
    
    // Extract text from the Google Doc
    var text = DocumentApp.openById(doc.getId()).getBody().getText();
    
    // Optionally, delete the temporary document
    DriveApp.getFileById(doc.getId()).setTrashed(true); // Move the created doc to trash
    
    return text;
  }
  