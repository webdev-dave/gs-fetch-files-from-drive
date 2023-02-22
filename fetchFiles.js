//ToDo
//1. 
//2. 


function fetchFiles() {
  // ------------------------ easy access --------------
  //column range
  const fileNamesCol = 13;
  const searchResultsCol = 12;
  //row range
  const firstRow = 6;
  const maxRows = firstRow + 50;

  // ----------=--------------------------- complex code -----------
  const activeSheet = SpreadsheetApp.getActive();
  const activeTab = SpreadsheetApp.getActiveSheet();
  const currentClientInvoicesFolder = DriveApp.getFolderById("181qKnkhVne7LOWyugeTT8oqu1PuGy1zV");
  //
  //
  //this deletes all previousClientInvoices 
  const previousClientInvoices = currentClientInvoicesFolder.getFiles();
  while(previousClientInvoices.hasNext()){
    previousClientInvoices.next().setTrashed(true);
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


  // ----------------------------------------------------- search helper function ----------------------

  function search(fileName, row) {
    if (!fileName.length) {
      //this clears all messages if there used to be a name in this location but there is no name currently
      activeTab.getRange(row, searchResultsCol).setValue("");
      //then return to break out of search();
      return;
    }

    let matchingFiles = DriveApp.getFilesByName(fileName);
    if(!matchingFiles.hasNext()){
      // if initial fileName search returned nothing, try alternative fileName
      const alterntiveFileName = fileName.slice(0,-4) + "_DIRECT.pdf";
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
      searchResult.makeCopy("copy of " + searchResult.getName() , currentClientInvoicesFolder);
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






