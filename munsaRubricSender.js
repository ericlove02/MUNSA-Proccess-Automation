function onOpen(e){
  SpreadsheetApp.getUi().createAddonMenu()
  .addItem('Create Rubric IDs', 'createRubrics').addToUi();
  //.addItem('School Awards', 'getSchoolAwards').addToUi();
  
}

function onInstall(e){
  onOpen(e);
}

function createRubrics() {
  var ui = SpreadsheetApp.getUi();
  var rubricTemplate = "1qYPWV_4cSIjt8Su2ueQj53R75ProccS2I12hfFQL-YI";
  
  var startStr = ui.prompt("Total Rubrics", "Enter the first row you are creating: ", ui.ButtonSet.OK).getResponseText();
  var endStr = ui.prompt("Total Rubrics", "Enter the last row you are creating: ", ui.ButtonSet.OK).getResponseText();
  
  var start = parseInt(startStr);
  var end = parseInt(endStr);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = ss.getSheets()[0];
  
  for(var i = start; i <= end; ++i){
    var committee = data.getRange("B"+i.toString()).getValue().trim();
    var country = data.getRange("C"+i.toString()).getValue().trim();
    var auth = data.getRange("D"+i.toString()).getValue();
    var speakersList = data.getRange("E"+i.toString()).getValue();
    var mod = data.getRange("F"+i.toString()).getValue();
    var unmod = data.getRange("G"+i.toString()).getValue();
    var parlipro = data.getRange("H"+i.toString()).getValue();
    var tact = data.getRange("I"+i.toString()).getValue();
    var publicSpeaking = data.getRange("J"+i.toString()).getValue();
    var bonus = data.getRange("K"+i.toString()).getValue();
    var comments = data.getRange("L"+i.toString()).getValue().trim();
    var participation = data.getRange("M"+i.toString()).getValue();
    var diplomacy = data.getRange("N"+i.toString()).getValue();
    var total = data.getRange("O"+i.toString()).getValue();
  
    var school = "";
    var newCol = "";
    
    country = country.trim();
    
    switch(country){
      case("Canada"): 
        school = "Tom C. Clark High School";
        newCol = "B";
        break;
      case("Germany"): 
        school = "Tom C. Clark High School";
        newCol = "B";
        break;
      case("Hungary"): 
        school = "Tom C. Clark High School";
        newCol = "B";
        break;
      case("Brazil"): 
        school = "Euroamerican School of Monterrey";
        newCol = "C";
        break;
      case("Bahamas"): 
        school = "Euroamerican School of Monterrey";
        newCol = "C";
        break;
      case("Netherlands"): 
        school = "Euroamerican School of Monterrey";
        newCol = "C";
        break;
      case("Ukraine"): 
        school = "La Vernia High School";
        newCol = "D";
        break;
      case("Sweden"): 
        school = "La Vernia High School";
        newCol = "D";
        break;
      case("Islamic Republic of Iran"): 
        school = "St. Andrew's Episcopal School";
        newCol = "E";
        break;
      case("United States of America"): 
        school = "Reagan High School";
        newCol = "F";
        break;
      case("Japan"): 
        school = "Reagan High School";
        newCol = "F";
        break;
      case("Chile"): 
        school = "Reagan High School";
        newCol = "F";
        break;
      case("Colombia"): 
        school = "Reagan High School";
        newCol = "F";
        break;
      case("Cyprus"): 
        school = "Reagan High School";
        newCol = "F";
        break;          
      case("Argentina"):
        school = "Reagan High School";
        newCol = "F";
        break;
      case("Cambodia"): 
        school = "Reagan High School";
        newCol = "F";
        break;
      case("Belize"): 
        school = "Reagan High School";
        newCol = "F";
        break;
      case("Syrian Arab Republic"): 
        school = "Keystone High School";
        newCol = "G";
        break;
      case("New Zealand"): 
        school = "Keystone High School";
        newCol = "G";
        break;
      case("Turkey"): 
        school = "Westlake High School";
        newCol = "H";
        break;
      case("Uruguay"): 
        school = "Westlake High School";
        newCol = "H";
        break;
      case("Zambia"): 
        school = "Westlake High School";
        newCol = "H";
        break;
      case("Australia"): 
        school = "Colegio Nexus";
        newCol = "I";
        break;
      case("Lithuania"): 
        school = "Colegio Nexus";
        newCol = "I";
        break;
      case("Italy"): 
        school = "Great Hearts Northern Oaks";
        newCol = "J";
        break;
      case("Zimbabwe"): 
        school = "Great Hearts Northern Oaks";
        newCol = "J";
        break;
      case("Jordan"): 
        school = "Great Hearts Northern Oaks";
        newCol = "J";
        break;
      case("Venezuela"): 
        school = "Meridian World School";
        newCol = "K";
        break;
      case("France"): 
        school = "Meridian World School";
        newCol = "K";
        break;
      case("Bolivia"): 
        school = "Meridian World School";
        newCol = "K";
        break;
      case("Poland"): 
        school = "Meridian World School";
        newCol = "K";
        break;
      case("Greece"): 
        school = "Advanced Learning Academy";
        newCol = "L";
        break;
      case("Ireland"): 
        school = "Advanced Learning Academy";
        newCol = "L";
        break;
      case("Belarus"): 
        school = "Advanced Learning Academy";
        newCol = "L";
        break;
      case("Switzerland"): 
        school = "Colegio Mirasierra";
        newCol = "M";
        break;
      case("United Kingdom"): 
        school = "Austin High School";
        newCol = "N";
        break;
      case("China"): 
        school = "Austin High School";
        newCol = "N";
        break;
      case("Czech Republic"): 
        school = "Austin High School";
        newCol = "N";
        break;
      case("Haiti"):
        school = "Macarthur High School";
        newCol = "O";
        break;
      case("Cuba"):
        school = "Macarthur High School";
        newCol = "O";
        break;
      case("Costa Rica"):
        school = "Macarthur High School";
        newCol = "O";
        break;
      case("Guyana"):
        school = "Macarthur High School";
        newCol = "O";
        break;
      case("Mexico"):
        school = "Johnson High School";
        newCol = "P";
        break;
      case("Guatemala"):
        school = "Johnson High School";
        newCol = "P";
        break;
      case("Austria"):
        school = "Johnson High School";
        newCol = "P";
        break;
      case("Mali"):
        school = "Johnson High School";
        newCol = "P";
        break;
      case("Denmark"):
        school = "Houston Christian High School";
        newCol = "Q";
        break;
      case("Egypt"):
        school = "Houston Christian High School";
        newCol = "Q";
        break;
      case("Qatar"):
        school = "Houston Christian High School";
        newCol = "Q";
        break;
      case("Belgium"):
        school = "The Oakridge School";
        newCol = "R";
        break;
      case("South Sudan"):
        school = "The Oakridge School";
        newCol = "R";
        break;
      case("Israel"):
        school = "Byron P. Steele High School";
        newCol = "S";
        break;
      case("Bulgaria"):
        school = "Byron P. Steele High School";
        newCol = "S";
        break;
      case("Portugal"):
        school = "Samuel Clemens High School";
        newCol = "T";
        break;
      case("Nepal"):
        school = "Samuel Clemens High School";
        newCol = "T";
        break;
      case("Georgia"):
        school = "Liberal Arts and Science Academy";
        newCol = "U";
        break;
      case("Iraq"):
        school = "Liberal Arts and Science Academy";
        newCol = "U";
        break;
      case("Russian Federation"):
        school = "Instituto Anglia";
        newCol = "V";
        break;
      case("Tunisia"):
        school = "Instituto Anglia";
        newCol = "V";
        break;
      case("United Arab Emirates"):
        school = "Lake Travis High School";
        newCol = "W";
        break;
      case("Philippines"):
        school = "Lake Travis High School";
        newCol = "W";
        break;
      case("Saudi Arabia"):
        school = "Lee High School";
        newCol = "X";
        break;
      case("Kenya"):
        school = "Lee High School";
        newCol = "X";
        break;
      case("Democratic People's Republic of Korea"):
        school = "Lee High School";
        newCol = "X";
        break;
      case("South Africa"):
        school = "CAST Tech High School";
        newCol = "Y";
        break;
      case("Kuwait"):
        school = "Churchill High School";
        newCol = "Z";
        break;
      case("Romania"):
        school = "Churchill High School";
        newCol = "Z";
        break;
      case("Saint Lucia"):
        school = "Churchill High School";
        newCol = "Z";
        break;
      case("Antigua and Barbuda"):
        school = "Churchill High School";
        newCol = "Z";
        break;
      case("Thailand"):
        school = "IDEA Carver College Prep";
        newCol = "AA";
        break;
      case("Guinea"):
        school = "IDEA Carver College Prep";
        newCol = "AA";
        break;
      case("Cote D'lvoire"):
        school = "Westchester Academy for International Studies";
        newCol = "AB";
        break;
      case("Bosnia and Herzegovina"):
        school = "Westchester Academy for International Studies";
        newCol = "AB";
        break;
      case("Trinidad and Tobago"):
        school = "Sharpstown International School";
        newCol = "AC";
        break;
      case("Saint Vincent and the Grenadines"):
        school = "Sharpstown International School";
        newCol = "AC";
        break;
      case("Barbados"):
        school = "Sharpstown International School";
        newCol = "AC";
        break;
      case("Dominica"):
        school = "Sharpstown International School";
        newCol = "AC";
        break;
      case("Jamaica"):
        school = "Sharpstown International School";
        newCol = "AC";
        break;
      case("Saint Kitts and Nevis"):
        school = "Sharpstown International School";
        newCol = "AC";
        break;
      case("Suriname"):
        school = "Sharpstown International School";
        newCol = "AC";
        break;
      case("The former Yugoslav Republic of Macedonia"):
        school = "Pleasanton High School";
        newCol = "AD";
        break;
      case("Yemen"):
        school = "Second Baptist School";
        newCol = "AE";
        break;
      case("Madagascar"):
        school = "American Institute of Monterrey Preparatory School";
        newCol = "AF";
        break;
      case("Paraguay"):
        school = "American Institute of Monterrey Preparatory School";
        newCol = "AF";
        break;
      case("Grenada"):
        school = "American Institute of Monterrey Preparatory School";
        newCol = "AF";
        break;
      case("Sri Lanka"):
        school = "American Institute of Monterrey Preparatory School";
        newCol = "AF";
        break;
      case("India"):
        school = "Round Rock High School";
        newCol = "AG";
        break;
      case("Republic of Korea"):
        school = "Round Rock High School";
        newCol = "AG";
        break;
      case("Vietnam"):
        school = "Round Rock High School";
        newCol = "AG";
        break;
      case("Mongolia"):
        school = "Round Rock High School";
        newCol = "AG";
        break;
      case("Niger"):
        school = "Wimberely High School";
        newCol = "AH";
        break;
      case("Laos"):
        school = "Travis Early College High School";
        newCol = "AI";
        break;
      case("Panama"):
        school = "Travis Early College High School";
        newCol = "AI";
        break;
      case("Guadeloupe"):
        school = "Travis Early College High School";
        newCol = "AI";
        break;
      case("Martinique"):
        school = "Travis Early College High School";
        newCol = "AI";
        break;
      case("Indonesia"):
        school = "Plano East Senior High School";
        newCol = "AJ";
        break;
      case("Liberia"):
        school = "Plano East Senior High School";
        newCol = "AJ";
        break;
      case("Sudan"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Kazakhstan"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Palestine"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Cameroon"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Iceland"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Bangladesh"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Nicaragua"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Ecuador"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Dominican Republic"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Angola"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Eritrea"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("State of Libya"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Ethiopia"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Norway"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Rwanda"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Botswana"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Croatia"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Oman"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Burkina Faso"):
        school = "International School of the Americas";
        newCol = "AK";
        break;
      case("Honduras"):
        school = "The Colony High School";
        newCol = "AL";
        break;
      case("Malaysia"):
        school = "The Colony High School";
        newCol = "AL";
        break;
      case("Serbia"):
        school = "The Colony High School";
        newCol = "AL";
        break;
      case("Lebanon"):
        school = "The Colony High School";
        newCol = "AL";
        break;
      case("Spain"):
        school = "TMI";
        newCol = "AM";
        break;
      case("Peru"):
        school = "TMI";
        newCol = "AM";
        break;
      case("Sierra Lenone"):
        school = "Thomas Jefferson High School";
        newCol = "AN";
        break;
      case("Singapore"):
        school = "Brooks Collegiate Academy";
        newCol = "AO";
        break;
      case("Namibia"):
        school = "Brooks Collegiate Academy";
        newCol = "AO";
        break;
      case("Finland"):
        school = "Central Catholic High School";
        newCol = "AP";
        break;
      case("Armenia"):
        school = "Central Catholic High School";
        newCol = "AP";
        break;
      case("Nigeria"):
        school = "IDEA South Flores";
        newCol = "AQ";
        break;
      case("Central African Republic"):
        school = "IDEA South Flores";
        newCol = "AQ";
        break;
      case("Chad"):
        school = "John Jay";
        newCol = "AR";
        break;
      case("Seychelles"):
        school = "Saint Mary's Hall";
        newCol = "AS";
        break;
      case("Liechtenstein"):
        school = "Saint Mary's Hall";
        newCol = "AS";
        break;
      default:
        school = "UNKNOWN";
        newCol = "AT";
        break;
    }
    
    
    
    var rubricName = school + " -- " + country + " " + committee + " Delegate Rubric";
    var DelegateRubric = DriveApp.getFileById(rubricTemplate)
    .makeCopy(rubricName)
    .getId();
    var copyDoc = DocumentApp.openById(DelegateRubric);

    var copyBody = copyDoc.getActiveSection();
    
    copyBody.replaceText('keyCommittee', committee);
    copyBody.replaceText('keyDelnam', country);
    copyBody.replaceText('keySchool', school);
    copyBody.replaceText('keyAuth', auth);
    copyBody.replaceText('keySpeak', speakersList);
    copyBody.replaceText('keyDelnam', country);
    copyBody.replaceText('keyMod', mod);
    copyBody.replaceText('keyUnmod', unmod);
    copyBody.replaceText('keyParlipro', parlipro);
    copyBody.replaceText('keyTact', tact);
    copyBody.replaceText('keyPubspeak', publicSpeaking);
    copyBody.replaceText('keyParti', participation);
    copyBody.replaceText('keyDiplo', diplomacy);
    copyBody.replaceText('keyBonus', bonus);
    copyBody.replaceText('keyTotal', total);
    copyBody.replaceText('keyComments', comments);
    
    copyDoc.saveAndClose();
    
    var idSheet = "1jq6kbH6DkYd1F2GUalzPAEcthKqAKEx7UVAWIRNNqck";
    
    if(newCol != ""){
      SpreadsheetApp.openById(idSheet).getSheetByName("Sheet1").getRange(newCol+(i+1).toString()).setValue(DelegateRubric);
    }
   
    
    SpreadsheetApp.getActiveSheet().getRange("P"+i.toString()).setValue("ID TRANSFERRED");
  
  }
}
  

