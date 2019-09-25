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
  var options = [""];
  if (responses[4] !== undefined) {
    options = responses[4].getResponse();
  }
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
  
  // debugging email
  //MailApp.sendEmail("ian@ada.ac.uk","kudos debugging",options[0]);
  
  // email the student too?
  // checkboxes are handled so badly!
  // getResponse yields an array of strings of option text or error if there are none
  // if we ever have more than one option we are not guaranteed that they'll even be returned in order!?
  // Above we arrange for options to be [""] if no options are selected
  var notify_student = options.indexOf("Copy in student's email") > -1;
  
  // send email
  var templ = HtmlService.createTemplateFromFile("kudos email template.html");
  templ.kudos = {student_name: student_name,
                 points: points,
                 respondent: respondent,
                 reason: reason,
                 value: value,
                 notified: notify_student};
  var msg = templ.evaluate().getContent();
  MailApp.sendEmail({to: team_email,
                     subject: "Kudos to "+student_name,
                     htmlBody: msg,
                     noReply: true});

  if (notify_student) {
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