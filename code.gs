//BASED FROM https://gist.github.com/automagictv/48bc3dd1bc785601422e80b2de98359e
//A simple tweak was added which is the conversion of the google docs to pdf, saved in gdrive
//You can also send the pdf as attachment to mail

//Make sure you provide the google docs id and destination files
const TEMPLATE_FILE_ID = '';
const DESTINATION_FOLDER_ID = '';
const PDF_FOLDER_ID = ''

// Parse and extract the data submitted through the form.
function parseFormData(values, header) {
    // Set temporary variables to hold prices and data.
    var response_data = {};

    // Iterate through all of our response data and add the keys (headers)
    // and values (data) to the response dictionary object.
    for (var i = 0; i < values.length; i++) {
      // Extract the key and value
      var key = header[i];
      var value = values[i];

      // Add the key/value data pair to the response dictionary.
      response_data[key] = value;
    }
    return response_data;
}

// Helper function to inject data into the template
function populateTemplate(document, response_data) {

    // Get the document header and body (which contains the text we'll be replacing).
    var document_header = document.getHeader();
    var document_body = document.getBody();

    // Replace variables in the header
    for (var key in response_data) {
      var match_text = `{{${key}}}`;
      var value = response_data[key];

      // Replace our template with the final values
      document_header.replaceText(match_text, value);
      document_body.replaceText(match_text, value);
    }

}


// Function to populate the template form
function createDocFromForm() {

  // Get active sheet and tab of our response data spreadsheet.
  var sheet = SpreadsheetApp.getActiveSheet();
  var last_row = sheet.getLastRow() - 1;

  // Get the data from the spreadsheet.
  var range = sheet.getDataRange();
 
  // Identify the most recent entry and save the data in a variable.
  var data = range.getValues()[last_row];
  
  // Extract the headers of the response data to automate string replacement in our template.
  var headers = range.getValues()[0];

  // Parse the form data.
  var response_data = parseFormData(data, headers);

  // Retreive the template file and destination folder.
  var template_file = DriveApp.getFileById(TEMPLATE_FILE_ID);
  var target_folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
  var target_pdf_folder = DriveApp.getFolderById(PDF_FOLDER_ID);
  
  // Copy the template file so we can populate it with our data.
  // The name of the file will be the company name and the invoice number in the format: DATE_COMPANY_NUMBER
  var filename = `${response_data["Your Name"]}`;
  var document_copy = template_file.makeCopy(filename, target_folder);

  // Open the copy.
  var document = DocumentApp.openById(document_copy.getId());

  // Populate the template with our form responses and save the file.
  populateTemplate(document, response_data);
  document.saveAndClose();
  
  //save as pdf
   var pdfContent = document.getAs('application/pdf');
   var pdfFile = target_pdf_folder.createFile(pdfContent.copyBlob());
   
  /* Remove the comment if you want to send the pdf to your mail
  //Send Email PDF Attachment
  MailApp.sendEmail({
      to: `${response_data["Your Email"]}`,
      name: "Google Docs PDF",
      subject: "Letter PDF attachment",
      htmlBody: `${response_data["Content"]}`,
      attachments: pdfFile.getAs(MimeType.PDF)
    });
   */
}
