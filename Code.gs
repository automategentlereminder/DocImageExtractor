function onOpen() {
    DocumentApp.getUi()
        .createMenu('Image link Extractor 🔗')
        .addItem('Extract links', 'appendImageLinks')
        .addItem('Help', 'helpBox')
        .addToUi();
}

function appendImageLinks() {
  var doc = DocumentApp.getActiveDocument();
  var docId = doc.getId(); // Gets the ID of the current document
  
  // Set the sharing settings to view-only for everyone
  DriveApp.getFileById(docId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // URL to fetch the current document as HTML
  var url = "https://docs.google.com/feeds/download/documents/export/Export?id=" + docId + "&exportFormat=html";
  var param = {
    method: "get",
    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions: true,
  };

  var html = UrlFetchApp.fetch(url, param).getContentText();

  // Extract 'src' attributes from 'img' tags
  var imgSrcRegex = /<img[^>]+src="([^">]+)"/g;
  var matches;
  var imageLinks = [];

  // Collect all image links
  while ((matches = imgSrcRegex.exec(html)) !== null) {
    imageLinks.push(matches[1]);
  }

  // Get the body of the document
  var body = doc.getBody();

  // Optionally, add a header to indicate the start of image links
  body.appendParagraph("Image links extracted from the document:");

  // Append a two-column table to the end of the document
  var table = body.appendTable();
  var headerRow = table.appendTableRow();
  headerRow.appendTableCell().setText("#").setWidth(25); // Adjust width for the first column
  headerRow.appendTableCell().setText("Link");

  // Append each image link with serial number to the table
  for (var i = 0; i < imageLinks.length; i++) {
    var row = table.appendTableRow();
    row.appendTableCell().setText(String(i + 1)); // Serial number starts from 1
    row.appendTableCell().setText(imageLinks[i]);
  }
}

function helpBox() {
    try {
        DocumentApp.getUi().alert(
            'Hi!',
            'Contact us at https://gentlereminder.in/ or just DM us on any social media with a screenshot and the issue you are facing, and we will get back to you soon!',
            DocumentApp.getUi().ButtonSet.OK
        );
    } catch (error) {}
}
