function formEmailInvoice() {
//written by Eric Love XXIV
//direct any questions to eric.love@yahoo.com
var ui = SpreadsheetApp.getUi();
var response1 = ui.prompt("Invoice ID", "Enter the invoice template ID (if not default): ",
ui.ButtonSet.OK).getResponseText(); //get from url of doc to be attached
var responseYN = ui.alert("Do you want to email sponsors now?", ui.ButtonSet.YES_NO);
var response2 = ui.prompt("Emails", "Enter email to send to (not including sponsors): ",
ui.ButtonSet.OK).getResponseText();
var response3 = ui.prompt("Rows", "Enter the number of the final row on the spreadsheet: ",
ui.ButtonSet.OK).getResponseText();
var invoiceTemplate = "1IeDaVQp4XEA1bMfM_sCGHf3pk58x1QQo1kRS8brzyI8";
if(!(response1 == "")){
invoiceTemplate == response1;
}
var ss = SpreadsheetApp.getActiveSpreadsheet();
var data = ss.getSheets()[0];
var sent = "";
var sentReceipt = "";
var remaining = (response3-1);
ui.alert("Check Columns", "Ensure that data columns are correct, check code for list in
comments", ui.ButtonSet.OK_CANCEL);
var htmlOutput = HtmlService.createHtmlOutput("&lt;p&gt;Emails sent:
&lt;/p&gt;").setWidth(600).setHeight(800);
SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Emails sent");

for(var i = 2; i &lt;= response3; ++i){
var school_name = data.getRange("B"+i.toString()).getValue().trim(); // B - School
Name
var school_add = data.getRange("D"+i.toString()).getValue().trim(); // D - School
Address
var school_city = data.getRange("E"+i.toString()).getValue().trim(); // E - School
City
var school_state = data.getRange("F"+i.toString()).getValue().trim(); // F - School
State

var school_zip = data.getRange("G"+i.toString()).getValue(); // G - School
Zip Code
var school_country = data.getRange("H"+i.toString()).getValue().trim(); // H - School
Country
var sponsor_name = data.getRange("I"+i.toString()).getValue().trim(); // I -
Sponsor Name
var sponsor_email = data.getRange("J"+i.toString()).getValue().trim(); // J -
Sponsor Email
var sponsor_phone = data.getRange("K"+i.toString()).getValue(); // K -
Sponsor Phone Number
var add_sponsors = data.getRange("L"+i.toString()).getValue().trim(); // L -
Additional Sponsor Names
var total_sponsors = data.getRange("M"+i.toString()).getValue(); // M - Total
Number of Sponsors
var total_delegates = data.getRange("N"+i.toString()).getValue(); // N - Total
Number of Delegates
var OAS_students = data.getRange("O"+i.toString()).getValue(); // O -
Number of OAS Students
var hotel_rooms = data.getRange("P"+i.toString()).getValue(); // P - Number
of Hotel Rooms
var delegate_dance = data.getRange("S"+i.toString()).getValue(); // S -
Number of Delegates to Delegate Dance
var XS = data.getRange("T"+i.toString()).getValue(); // T - Number of
XS Shirts
var S = data.getRange("U"+i.toString()).getValue(); // U - Number of S
Shirts
var M = data.getRange("V"+i.toString()).getValue(); // V - Number of M
Shirts
var L = data.getRange("W"+i.toString()).getValue(); // W - Number of L
Shirts
var XL = data.getRange("X"+i.toString()).getValue(); // X - Number of
XL Shirts
var XXL = data.getRange("Y"+i.toString()).getValue(); // Y - Number of
XXL Shirts
var country1 = data.getRange("Z"+i.toString()).getValue().trim(); // Z - First
Choice Country
var country2 = data.getRange("AA"+i.toString()).getValue().trim(); // AA -
Second Choice Country
var country3 = data.getRange("AB"+i.toString()).getValue().trim(); // AB - Third
Choice Country
var country4 = data.getRange("AC"+i.toString()).getValue().trim(); // AC - Fourth
Chocie Country
var country5 = data.getRange("AD"+i.toString()).getValue().trim(); // AD - Fifth
Choice Country

var special_committee1 = data.getRange("AE"+i.toString()).getValue().toString().trim(); // AE -
First Choice Special Committee
var special_committee2 = data.getRange("AF"+i.toString()).getValue().toString().trim(); // AF -
Second Choice Special Committee
//**If any of this is different on the
spreadsheet change the letter next to the objects**//
if(OAS_students == "N/A"){OAS_students = 0;}
if(hotel_rooms == "N/A"){hotel_rooms = 0;}
if(delegate_dance == "N/A"){delegate_dance = 0;}
if(XS == 0 || XS == "N/A"){XS = "";}
if(S == 0 || S == "N/A"){S = "";}
if(M == 0 || M == "N/A"){M = "";}
if(L == 0 || L == "N/A"){L = "";}
if(XL == 0 || XL == "N/A"){XL = "";}
if(XXL == 0 || XXL == "N/A"){XXL = "";}
if(country1 == "NA" || country1 == "N/A"){country1 == "";}
if(country2 == "NA" ||country2 == "N/A"){country2 == "";}
if(country3 == "NA" ||country3 == "N/A"){country3 == "";}
if(country4 == "NA" ||country4 == "N/A"){country4 == "";}
if(country5 == "NA" ||country5 == "N/A"){country5 == "";}
if(special_committee1 == "NA" || special_committee1 == "N/A" || special_committee1 ==
"0"){special_committee1 = "";}
if(special_committee2 == "NA" || special_committee2 == "N/A" || special_committee2 ==
"0"){special_committee2 = "";}
var USGEA = response2; // usg external affairs email
// calculate vars for invoice
var delegate_cost = total_delegates * 65; // cost per delegate
var sponsor_cost = total_sponsors * 50; // cost per sponsor
var dance_cost = delegate_dance * 5; // cost per dance
var shirt_total = Number(XS) + Number(S) + Number(M) + Number(L) + Number(XL) +
Number(XXL);
var shirts_cost = shirt_total * 15; // cost per shirt
var grand_total = sponsor_cost + delegate_cost + dance_cost + shirts_cost + 100;
var curDate = new Date();
var dd = curDate.getDate();
var mm = curDate.getMonth() + 1;
var yy = curDate.getYear();

var invoiceName = school_name + " Invoice " + mm + "-" + dd + "-" + yy;

var SchoolInvoice = DriveApp.getFileById(invoiceTemplate)
.makeCopy(invoiceName)
.getId();
var copyDoc = DocumentApp.openById(SchoolInvoice);
var copyBody = copyDoc.getActiveSection();
copyBody.replaceText(&#39;keyHSpersonname&#39;, sponsor_name);
copyBody.replaceText(&#39;keyHSSchool&#39;, school_name);
copyBody.replaceText(&#39;keyHSAddress&#39;, school_add);
copyBody.replaceText(&#39;keyHSCity&#39;, school_city);
copyBody.replaceText(&#39;keyHSState&#39;, school_state);
copyBody.replaceText(&#39;keyHSZip&#39;, school_zip);
copyBody.replaceText(&#39;keyHSdelegate&#39;, total_delegates);
copyBody.replaceText(&#39;keyHSSponsor&#39;, total_sponsors);
copyBody.replaceText(&#39;keyHSdance&#39;, delegate_dance);
copyBody.replaceText(&#39;keyHSshirttotal&#39;, shirt_total);
copyBody.replaceText(&#39;keyHSdelegatcost&#39;, delegate_cost);
copyBody.replaceText(&#39;keyHSsorcost&#39;, sponsor_cost);
copyBody.replaceText(&#39;keyHSdanccost&#39;, dance_cost);
copyBody.replaceText(&#39;keyHSshirtcost&#39;, shirts_cost);
copyBody.replaceText(&#39;keyHSgrandtotal&#39;, grand_total);
copyBody.replaceText(&#39;keyHSShirtXS&#39;, XS);
copyBody.replaceText(&#39;keyHSShirtS&#39;, S);
copyBody.replaceText(&#39;keyHSShirtM&#39;, M);
copyBody.replaceText(&#39;keyHSShirtL&#39;, L);
copyBody.replaceText(&#39;keyHSShirtXL&#39;, XL);
copyBody.replaceText(&#39;keyHSShirtXXL&#39;, XXL);
copyDoc.saveAndClose();

var subject1 = "MUNSA XXIV:Envision registration -- " + school_name;

/////////////////////////////////////////////////////////////////////////////////
var body = sponsor_name + "&lt;img
src=\"https://drive.google.com/uc?export=view&amp;id=0B7kyD_PT0o14azNac3RveTFMeDRHQlhx
RTd2QVVZai1yY3Q4\" alt=\"Envision Logo\" " +
" height=\"250\" width=\"210\" align=\"right\"&gt;" + "&lt;br&gt;" + school_name + "&lt;br&gt;" +
school_add + "&lt;br&gt;" + school_city + ", " + school_state + " " + school_zip + " ";

if(school_country != "USA"){

body += school_country;
}

body += "&lt;br&gt;&lt;br&gt;&lt;h2&gt;Thank you, " + sponsor_name + ", for your interest in MUNSA XXIV:
&lt;i&gt;Envision&lt;/i&gt;.&lt;/h2&gt;" +
"We&#39;ve received your registration request for " + total_delegates + " delegate(s) and " +
total_sponsors + " sponsor(s). ";

if(!(OAS_students == "")){
body += "You have indicated that " + OAS_students + " delegate(s) speak fluent Spanish and
will be particapting in the OAS delegation. &lt;br&gt;&lt;br&gt;";
}
else{
body += "&lt;br&gt;&lt;br&gt;";
}

if(!(country1 == "" &amp;&amp; country2 == "" &amp;&amp; country3 == "" &amp;&amp; country4 == "" &amp;&amp; country5 == ""
&amp;&amp; special_committee1 == "" &amp;&amp; special_committee2 == "")){
body += "Your requests are as follows: ";
if(!(country1 == "")){
body += country1;
}
if(!(country2 == "")){
body += ", " + country2;
}
if(!(country3 == "")){
body += ", " + country3;
}
if(!(country4 == "")){
body += ", " + country4;
}
if(!(country5 == "")){
body += ", and " + country5;
}
body += ". Your specialized committee country/representative requests are as follows: ";
if(!(special_committee1 == "")){
body += special_committee1;
}
if(!(special_committee2 == "")){
body += " and " + special_committee2;

}
body += ". I will email you at a later time to confirm your assignments.&lt;br&gt;&lt;br&gt;";
}
else if(!(country1 == "" &amp;&amp; country2 == "" &amp;&amp; country3 == "" &amp;&amp; country4 == "" &amp;&amp; country5
== "") &amp;&amp; (special_committee1 == "" &amp;&amp; special_committee2 == "")){
body += "Your country requests are as follows: ";
if(!(country1 == "")){
body += country1;
}
if(!(country2 == "")){
body += ", " + country2;
}
if(!(country3 == "")){
body += ", " + country3;
}
if(!(country4 == "")){
body += ", " + country4;
}
if(!(country5 == "")){
body += ", and " + country5;
}
body += ". I will email you at a later time to confirm your country assignments.&lt;br&gt;&lt;br&gt;";
}
else if(!(country1 == "" &amp;&amp; country2 == "" &amp;&amp; country3 == "" &amp;&amp; country4 == "" &amp;&amp; country5
== "") &amp;&amp; (special_committee1 == "" &amp;&amp; special_committee2 == "")){
body += ". Your specialized committee country/representative requests are as follows: ";
if(!(special_committee1 == "")){
body += special_committee1;
}
if(!(special_committee2 == "")){
body += " and " + special_committee2;
}
body += ". I will email you at a later time to confirm your assignments.&lt;br&gt;&lt;br&gt;";
}
else{
body += "You have no country or specialized committee requests. If this needs to be
changed, please email me immediately. &lt;br&gt;&lt;br&gt;";
}

body += "You have " + delegate_dance + " delegate(s) attending the delegate
dance.&lt;br&gt;&lt;br&gt;";

if(!(XS == "" &amp;&amp; S == "" &amp;&amp; M == "" &amp;&amp; L == "" &amp;&amp; XL == "" &amp;&amp; XXL == "")){
body += "Your t-shirt order includes ";
if(!(XS == "" || XS == 0)){
body += XS + "XS ";
}
if(!(S == "" || S == 0)){
body += S + "S ";
}
if(!(M == "" || M == 0)){
body += M + "M ";
}
if(!(L == "" || L == 0)){
body += L + "L ";
}
if(!(XL == "" || XL == 0)){
body += XL + "XL ";
}
if(!(XXL == "" || XXL == 0)){
body += XXL + "XXL ";
}
body += "shirts. &lt;br&gt;&lt;br&gt;";
}
else{
body += "You did not order any MUNSA t-shirts.&lt;br&gt;&lt;br&gt;";
}

if(!(hotel_rooms == 0 || hotel_rooms == "")){
body += "You have indicated that you will be needing " + hotel_rooms + " room(s) from &lt;a
href =\"https://www.omnihotels.com/hotels/san-antonio\"&gt;the Omni Hotel&lt;/a&gt;. "
+ "Please remember that the MUNSA Staff is not in charge of booking your
rooms.&lt;br&gt;&lt;br&gt;";
}
else{
body += "You have indicated that you will not be needing rooms from &lt;a href
=\"https://www.omnihotels.com/hotels/san-antonio\"&gt;the Omni Hotel&lt;/a&gt;. ";
}

body += "If any of this information needs to be changed, please email me
immediately.&lt;br&gt;&lt;br&gt;Best regards,&lt;br&gt;&lt;a href = \"mailto:rtubbesing8964@stu.neisd.net\"&gt;"+
"Ryan Tubbesing&lt;/a&gt;&lt;br&gt;&lt;br&gt;" + mm + "-" + dd + "-" + yy;
/////////////////////////////////////////////////////////////////////////////////
var message = {to: USGEA, subject: subject1, htmlBody: body,
attachments:[DriveApp.getFileById(SchoolInvoice).getAs(&#39;application/pdf&#39;)]};
if(responseYN == ui.Button.YES){
var message2 = {to: sponsor_email, subject: subject1, htmlBody: body,
attachments:[DriveApp.getFileById(SchoolInvoice).getAs(&#39;application/pdf&#39;)]};
MailApp.sendEmail(message2);
sent += sponsor_name + " -- " + sponsor_email + ", " + school_name + "\n";
sentReceipt += sponsor_name + " -- " + sponsor_email + ", " + school_name + "&lt;br&gt;";
}
MailApp.sendEmail(message);
--remaining;
htmlOutput = HtmlService.createHtmlOutput("&lt;p&gt;Emails have been sent to: &lt;br&gt;&lt;/p&gt;" +
sentReceipt + "&lt;br&gt;&lt;br&gt;&lt;p&gt; Emails remaining: &lt;/p&gt;" +
remaining).setWidth(600).setHeight(800);
SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Emails sent");
}
var receipt = {to: USGEA, subject: "Email Invoices Sent Receipt", htmlBody: "Invoice Emails
have been sent to: &lt;br&gt;" + sentReceipt + "&lt;br&gt;&lt;br&gt; Check sent mailbox to see invoices or look
at your copies."};
MailApp.sendEmail(receipt);
ui.alert("Emails have been sent to: \n" + sent + "\n\n All Emails have been sent\nIn addition, all
emails were sent to " + USGEA + "\nCheck " + USGEA + " for receipt of sent emails");
}