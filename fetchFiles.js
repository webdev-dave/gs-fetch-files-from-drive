//ToDo
//1. make a button that will cause program to run
//2. make search run only within a certain directory 


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




  // //Begin Implimentation of Strict Search within a chosen parent folder
  
  // //issue: can only search within DIRECT children of parent folder (and not children of children);
  
  // // simple, quick and cheap solutions
  //    // a. create a parent folder which has a copy of all invoices as DIRECT CHILDREN of the PARENT directory 
  //    // b. create a new google drive account that is specifically used for invoice pdf's (this will speed up searching in ENTIRE DRIVE)

  // // alterntive expensive and not necessarily quicker i can try to build and implement a custom solution but i would advice against it


  // const searchDirectoryIterator = DriveApp.getFoldersByName("Test Files Directory");
  // const searchDirectory = searchDirectoryIterator.next();
  // const filesIterator = searchDirectory.getFilesByName();
  // if(filesIterator.hasNext()){
  //   Logger.log(filesIterator.next())
  // } else {
  //   Logger.log("nothing found")
  // }


