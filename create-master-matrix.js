// created for MUNSA XXVI by Eric Love
// eric.love02@yahoo.com

function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
  .addItem("Send to Matrix", "sendtoMatrix")
  .addToUi();

};

function onInstall(e) {
  onOpen(e);
};

function sendtoMatrix(e) {
  var ui = DocumentApp.getUi();
  var response1 = ui.prompt("Committee", "Enter the name of the committee to fill: ", ui.ButtonSet.OK).getResponseText(); 
  var response2 = ui.prompt("Number of Countries", "Enter the number of countries in the room: ", ui.ButtonSet.OK).getResponseText(); 
  var response3 = ui.prompt("Committee listed after", "Enter the name of the commitee listed next on the document, if this is last type 'None': ", ui.ButtonSet.OK).getResponseText(); 
  var response4 = ui.prompt("Sheet Column", "Enter the column for " + response1 + " from the matrix: ", ui.ButtonSet.OK).getResponseText(); 

  var doc = DocumentApp.openById('1wiKROs2yX79PDdCNiLY6fFBCP4AfW5NDqaQOhGUsY_U');
  var body = doc.getBody();
  var text = body.editAsText();
  var textString = text.getText();

  if(response3 != "None"){
    var committeList = textString.substring(textString.search(response1) + response1.length + 3 + response2.toString.length + 2, textString.search(response3)).split("\n").filter(item => item).filter(doesntInclude);
  } 
  else {
    var committeList = textString.substring(textString.search(response1) + response1.length + 3 + response2.toString.length + 2).split("\n").filter(item => item).filter(doesntInclude);
  }

  //console.log(response1, response2, committeList);

  var ss = SpreadsheetApp.openById("1kq0yiuOyr__2vvEGX9oNxELSPvs2mH5HPlRSBgwHavQ");
  var sheet = ss.getSheetByName("Master Matrix"); 
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  var committeesNotFound = [];
  for(m = 0; m < committeList.length; m++){
    
    if(!searchString(committeList[m].trim()) || committeList[m].trim() === "China" || committeList[m].trim() === "United States" || committeList[m].trim() === "United Kingdom" || committeList[m].trim() === "United Kingdom of Great Britain"){
      committeesNotFound.push(committeList[m]);
    }
  }
console.log(committeesNotFound)
  for(var l = 0; l < committeList.length; l++){
    searchForCommittee = committeList[l].trim();

    for (var i = 0; i < values.length; i++) {
      var row = "";
      for (var j = 0; j < values[i].length; j++) {     
        if (values[i][j] == searchForCommittee) {
          row = values[i][j+1];
          valueRow = i + 1;

        }
      }    
    } 

    try{
      sheet.getRange(response4 + valueRow).setValue("1");
    } catch(e){
      console.log("dont care L + ratio + " + e);
      committeesNotFound.push(searchForCommittee);
    }
    
    
  }
  if(sheet.getRange(response4 + "181").getValue() != response2){
    ui.alert("It appears some countries may have been skipped. Check through the name on the document and the sheet and make sure everything matches up. \n\n These countries may have been missed: \n" + prettyArray(committeesNotFound));

  }
  else{
    ui.alert("Done!");
  }
  sheet.getRange(response4 + "181").setValue("=SUM("+response4+"2:"+response4+"180)")
}

function doesntInclude(value) {
  if(value.indexOf(':') != -1 || value.indexOf(' -') != -1 || value.indexOf('- ') != -1 || value.indexOf('\\') != -1){
    return false;
  }
  else{
    return true;
  }
}

function searchString(value){
  var ss = SpreadsheetApp.openById("1kq0yiuOyr__2vvEGX9oNxELSPvs2mH5HPlRSBgwHavQ");
  var sheet = ss.getSheetByName("Master Matrix"); 
  var search_string = value
  var textFinder = sheet.createTextFinder(search_string)
  return textFinder.findNext();
}

function prettyArray(array){
  str = "";
  for(i = 0; i < array.length; i++){
    str += array[i] + "\n";
  }
  return str;
}
