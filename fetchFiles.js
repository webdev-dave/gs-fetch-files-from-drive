//NOTES for users
// 1. fetchFiles finds the file(s) we want to email and stores a copy in Folder "000 Email Signed PDFs"
//    The copied file(s) has a "_" appended to it.
//    The variable name for this folder is :  const emailFolderId = "1AKzlniDXyDQlV2RnhuaSdhzUc3HbireO"
// 2. Search currently runs on all files in the (parent) folder named "LEAD Remedial". 
//    The variable name for the above folder in the code below is: const emailFolderId = "1AKzlniDXyDQlV2RnhuaSdhzUc3HbireO"

// ------------------------ code for users --------------
//column range
const fileNamesCol = 14;
const searchResultsCol = 13;
//row range
const firstRow = 6;
const maxRows = firstRow + 50;



// -------------------------------------- complex code (for developers) -----------


const activeSheet = SpreadsheetApp.getActive();
const activeTab = SpreadsheetApp.getActiveSheet();
const testEmailFolderId = "181qKnkhVne7LOWyugeTT8oqu1PuGy1zV";
//const emailFolderId = "1AKzlniDXyDQlV2RnhuaSdhzUc3HbireO";
const currentClientInvoicesFolder = DriveApp.getFolderById(testEmailFolderId);
const testCustomTrashId = "1NUA3UXIU9b4msXK5_7GMd7Xg0_WRLHqS"
//const customTrashId = "10IE1J-NV7q9eyTakYHCFTnb2LHSftWsP";
const customTrashFolder = DriveApp.getFolderById(testCustomTrashId);


function fetchFiles() {

  //this deletes all previousClientInvoices 
  const previousClientInvoices = currentClientInvoicesFolder.getFiles();
  while (previousClientInvoices.hasNext()) {
    const currentFile = previousClientInvoices.next();
    currentFile.moveTo(customTrashFolder);
  }
  //the loop below is where the actual spreadsheet search iterations begin

  for (let i = firstRow; i <= maxRows; i++) {
    const currentRow = i;
    const currentFileName = activeTab.getRange(currentRow, fileNamesCol).getValue();
    const previousSearchResult = activeTab.getRange(currentRow, searchResultsCol).getValue();

    if (currentFileName || previousSearchResult) {
      search(currentFileName, currentRow);
    }
  }

//
  // ----------------------------------------------------- search helper function ----------------------

  function search(fileName, row) {
    if (!fileName.length) {
      //this clears all messages if there used to be a name in this location but there is no name currently
      activeTab.getRange(row, searchResultsCol).setValue("");
      //then return to break out of search();
      return;
    }

    let matchingFiles = DriveApp.getFilesByName(fileName);
    
    if (!matchingFiles.hasNext()) {
      // if initial fileName search returned nothing, try alternative fileName
      const alterntiveFileName = fileName.slice(0, -4) + "_DIRECT.pdf";
      matchingFiles = DriveApp.getFilesByName(alterntiveFileName);
    }


    const searchResultsArray = [];
    while (matchingFiles.hasNext()) {
      const file = matchingFiles.next();
      searchResultsArray.push(file);
    };

    if (searchResultsArray.length === 1) {
      const searchResult = searchResultsArray[0];
      const richValue = SpreadsheetApp.newRichTextValue()
        .setText(searchResult)
        .setLinkUrl(searchResult.getUrl())
        .build();
      //set search result
      activeTab.getRange(row, searchResultsCol).setRichTextValue(richValue);
      //
      //make a copy of the pdf and add to currentClientInvoicesFolder
      //
      searchResult.makeCopy(searchResult.getName() + "_", currentClientInvoicesFolder);
    } else {
      let errorText = "No Matching Files!";;
      if (searchResultsArray.length > 1) {
        errorText = `Here is a link to the first result. (There are ${searchResultsArray.length} duplicate files with the name ${fileName})`;
      }
      //set searchResult to errorText
      activeTab.getRange(row, searchResultsCol).setValue(errorText);
      activeTab.getRange(row, searchResultsCol).setFontColor("red");
    }
  }
}



function mendelEmptyCustomTrash() {
  const trashFiles = customTrashFolder.getFiles();
  while(trashFiles.hasNext()){
    const file = trashFiles.next();
    const ownerName = file.getOwner().getName();
    if(ownerName === "Mendel Lichtenstein"){
      file.setTrashed(true);
    }
  }
}

function simonEmptyCustomTrash() {
  const trashFiles = customTrashFolder.getFiles();
  while(trashFiles.hasNext()){
    const file = trashFiles.next();
    const ownerName = file.getOwner().getName();
    if(ownerName === "Simon Licht"){
      file.setTrashed(true);
    }
  }
}





// --------------------------------------------- for testing purposes only!! --------------------------------------------------------------

function davidEmptyCustomTrash() {
  const trashFiles = customTrashFolder.getFiles();
  while(trashFiles.hasNext()){
    const file = trashFiles.next();
    const ownerName = file.getOwner().getName();
    if(ownerName === "David Richard"){
      file.setTrashed(true);
    }
  }
}

