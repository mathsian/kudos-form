function onFormOpen(e) {
  // get a reference to this form and the first question on it
  // assumes the first question is the email dropdown
  var form = FormApp.getActiveForm();
  var dropdown = form.getItems(FormApp.ItemType.LIST)[0].asListItem();
  
  // locate ourselves and get an id for the master sheet
  var cohortFolder = DriveApp.getFileById(form.getId()).getParents().next();
  if (cohortFolder.getFilesByName("master").hasNext()) {
    var master_id = cohortFolder.getFilesByName("master").next().getId();
    // save it for the submit function later NOT WORKING
    //PropertiesService.getScriptProperties().setProperty("master id", master_id);
    var ss = SpreadsheetApp.openById(master_id);
  } else {
    FormApp.getUi().alert("Cannot find master sheet");
    return;
  }
  
  var email_range = ss.getRangeByName("email").getDisplayValues();
  var emails = [];
  for (var i = 0; i<email_range.length; i++) {
    // the first column will be the email addresses
    // if we hit an empty string that's the end of the list, otherwise grab that email
    if (email_range[i][0].length > 0) {
      emails.push(email_range[i][0]);
    } else {
      break;
    }
  }
  // populate the dropdown
  dropdown.setChoiceValues(emails);
}
function onFormSubmit(e) {
  
  // get the response items
  var responses = e.response.getItemResponses();
  var student_email = responses[0].getResponse();
  var points = responses[1].getResponse();
  var reason = responses[2].getResponse();
  var value = responses[3].getResponse();
  var options = responses[4].getResponse();
  var respondent = e.response.getRespondentEmail();
  var date = Date();
  
  // locate ourselves and get an id for the master sheet
  var form = FormApp.getActiveForm();
  var cohortFolder = DriveApp.getFileById(form.getId()).getParents().next();
  if (cohortFolder.getFilesByName("master").hasNext()) {
    var master_id = cohortFolder.getFilesByName("master").next().getId();
    // save it for the submit function later NOT WORKING
    //PropertiesService.getScriptProperties().setProperty("master id", master_id);
    var ss = SpreadsheetApp.openById(master_id);
  } else {
    FormApp.getUi().alert("Cannot find master sheet");
    return;
  }

  var email_range = ss.getRangeByName("email").getDisplayValues();
  var team_range = ss.getRangeByName("team").getDisplayValues();
  var name_range = ss.getRangeByName("firstname").getDisplayValues();
  //var reportid_range = ss.getRangeByName("reportid").getDisplayValues();
  
  
  var i = indexOfStudent(student_email, email_range);

  if (i > -1) {
    var team_email = team_range[i][0];
    var student_name = name_range[i][0];
    //var reportid = reportid_range[i][0];
  } else {
    MailApp.sendEmail("datareporting@ada.ac.uk","Kudos form failure", "Failed for "+student_email,{noReply: true});
    return;
  }
  
  // send email
  var templ = HtmlService.createTemplateFromFile("kudos email template.html");
  templ.kudos = {student_name: student_name,
                 points: points,
                 respondent: respondent,
                 reason: reason,
                 value: value};
  var msg = templ.evaluate().getContent();
  MailApp.sendEmail({to: team_email,
                     subject: "Kudos to "+student_name,
                     htmlBody: msg,
                     noReply: true});
  if ("Copy in the student's email" in options) {
    templ = HtmlService.createTemplateFromFile("student kudos template.html");
    templ.kudos = {student_name: student_name,
                 points: points,
                 respondent: respondent,
                 reason: reason,
                 value: value};
    MailApp.sendEmail({to: student_email,
                     subject: "Kudos!",
                     htmlBody: msg,
                     noReply: true});
  }
  
  // *** Unnecessary - use linked sheet instead
  // update the student report kudos tab
  //var student_report_ss = SpreadsheetApp.openById(reportid);
  //var kudos_sheet = student_report_ss.getSheetByName("Kudos");
  //if (kudos_sheet == null){
  //    kudos_sheet = student_report_ss.insertSheet("Kudos")
  //}
  // this has simultaneous acces problems - use appendrow instead
  //var new_row_number = kudos_sheet.getLastRow()+1;
  //kudos_sheet.getRange(new_row_number, 1).setValue(date);
  //kudos_sheet.getRange(new_row_number, 2).setValue(respondent);
  //kudos_sheet.getRange(new_row_number, 3).setValue(reason);
  //kudos_sheet.appendRow([date, respondent, reason]);
}

function indexOfStudent(student_email, email_range) {
  // why are lookups so painful in apps script???
  for (var i = 0; i<email_range.length; i++) {
    if (email_range[i][0] === student_email) {
      return i;
    }
  }
  return -1;
}