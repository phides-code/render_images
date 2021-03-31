/*
This script runs on a Google Sheet formatted as follows, containing a list of 
images stored on Google Drive, with URLS and descriptions:

Column A: Category
Column B: Keywords
Column C: Google Drive Link
Column D: Date Added
Column E: User's e-mail address

The images are rendered in a Google Doc with keywords.
*/

function onOpen() {
    createCustomMenus();
  }
  
// when the URL column is modified, automatically fill in the date/time stamp and user's e-mail 
function onEdit() { 
    
    var s = SpreadsheetApp.getActiveSheet();
    
    if( s.getName() == "URL List" ) { // checks that we're on the correct sheet
      
      var r = s.getActiveCell();
      
      if( r.getColumn() == 3 ) { // if we're on column 3 
      
        var dateColumn = r.offset(0, 1); // set the date column 1 to the right 
        
        dateColumn.setValue(new Date()); // write the date
        
        var userColumn = r.offset(0, 2); // set the user column 2 to the right 
        
        userColumn.setValue(Session.getActiveUser().getEmail()); // write the user
      }
    }
  }
    
function createCustomMenus() { // menu item to trigger the reGenerateDocument function
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('My Scripts')
      .addItem('reGenerateDocument', 'reGenerateDocument')
      .addToUi();
  }
  
function reGenerateDocument() {
  
    var keywordUrlPairs = Sheets.Spreadsheets.Values.get(/*'hash/ID of the google sheet'*/, 'B2:C999'); // Google Sheet ID and relevant range
    var keywords;
    var myDoc = DocumentApp.openById(/*'hash/ID of the google doc'*/).getBody(); 
    var myText = myDoc.editAsText();
    var imageUrl;
    var imageId;
    var imageBlob;
    var appendedImage;
    var appendedImageScaled;
    var width;
    var height;
    var ratio;
    var previewHeight; 
    var previewWidth = 480; // desired Width for the image preview
  
    myDoc.clear(); // clear the document
    
    for(var i = 0; i < keywordUrlPairs.values.length; i++){ // for each row in the sheet, do the following
  
      keywords = keywordUrlPairs.values[i][0];  // get the keywords from column A of this row
      imageUrl = keywordUrlPairs.values[i][1];  // get the URL from column B of this row
      
      imageId = imageUrl.match(/[-\w]{25,}/); // get the image ID number from the URL
          
      myDoc.appendParagraph(keywords); // create a new paragraph and insert the keywords 
      myText.appendText('\n'); // insert a line break
  
      imageBlob = DriveApp.getFileById(imageId).getBlob();  // get the image blob from the URL
      appendedImage = myDoc.appendImage(imageBlob); // append the image
  
      width = appendedImage.getWidth(); // get the image width
      height = appendedImage.getHeight(); // get the image height
      ratio = width/height; // determine the w/h ratio of the image
      previewHeight = previewWidth/ratio; // set the new width based on the desired preview height/ratio
        
      appendedImage.setWidth(previewWidth).setHeight(previewHeight).setLinkUrl(imageUrl); // insert the image at the desired scale, and hyperlinked
          
      myText.appendText('\n'); // insert a line break
      myDoc.appendHorizontalRule(); // insert a horizontal line
    }
  
    Browser.msgBox(i);  // Display the number of rows counted
  }
  