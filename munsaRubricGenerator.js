function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
 
  .addItem('Open Menu', 'showSidebar') 
      .addToUi();
}

function onInstall(e){
  onOpen(e);
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Send Delegate Rubrics')
      .setWidth(300);
  SpreadsheetApp.getUi() 
      .showSidebar(html);
}

function sendRubrics() {
  try{
  var ui = SpreadsheetApp.getUi();
  
  var schoolName = ui.prompt("Rubrics", "Enter the school's name: ", ui.ButtonSet.OK).getResponseText();
  var column = ui.prompt("Rubrics", "Enter the school's column: ", ui.ButtonSet.OK).getResponseText();
  var lastCell = ui.prompt("Rubrics", "Enter the last row of the school's column: ", ui.ButtonSet.OK).getResponseText();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = ss.getSheets()[0];
  
  var trueSchool = data.getRange(column + "1").getValue();
  
  if(schoolName != trueSchool){ui.alert("Check school name and column", ui.ButtonSet.OK);}
  else{
    var attachments = [];
    var id = "";
    
    for(var i = 3; i <= lastCell; ++i){
      var checkCell = parseInt(lastCell) + 1;
      id = data.getRange(column+""+i.toString()).getValue();  
      if(id != ""){
        attachments.push(DriveApp.getFileById(id).getAs('application/pdf'));
        SpreadsheetApp.getActiveSheet().getRange(column+checkCell.toString()).setValue("CELL " + i.toString() + " ADDED");
      }
    }
  
  }
  
  var body = "Hello sponsors! \nAttached are your delegate's rubrics from MUNSA's non-crisis rooms. Please let us know if you are missing anything or have any questions.";
  body += " On behalf of Secretariat and all of MUNSA, thank you so much for attending our conference and we hope to see you next year! \n\n Ryan Tubbesing \n"
  var email = data.getRange(column+"2").getValue();
  //  var message = 
  //MailApp.sendEmail(message)
    GmailApp.createDraft(email, schoolName + ' -- MUNSA XXIV: Regular Room Delegate Rubrics', body, {attachments: attachments, name: 'MUNSA'});
  
  
  ui.alert("School sending complete! Email in drafts.", ui.ButtonSet.OK);
}
  catch(error){
    var ui = SpreadsheetApp.getUi();
    ui.alert("Error occured while trying to send rubrics.\n More information on error: \n\n [" + error + "]", ui.ButtonSet.OK);
  }
}
