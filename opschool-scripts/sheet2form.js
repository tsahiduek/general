function deleteUnusedFormSheets(e) {
  // delete all existing triggers to Spreadsheet
  var all_triggers = ScriptApp.getProjectTriggers()
  for (var i = 0; i < all_triggers.length; i++) {
    ScriptApp.deleteTrigger(all_triggers[i]);
  }
  // delete all Responses sheets + unlink them from their Form
  ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var formVar = sheets[i].getName().split(" ")[0];

    if (formVar == "Form") {
      var sheet_name = ss.getSheetByName(sheets[i].getName());

      var formUrl = sheets[i].getFormUrl();
      if (formUrl) {
        FormApp.openByUrl(formUrl).removeDestination();
        form_id = FormApp.openByUrl(formUrl).getId();
        // Delete the Form
        DriveApp.getFileById(form_id).setTrashed(true);

      }
      ss.deleteSheet(sheets[i]);
    }
  }


}


/**
 * A special function that inserts a custom menu when the spreadsheet opens.
 */
function onOpen() {
  var menu = [{
    name: 'Set up OpsschoolForm',
    functionName: 'setUpOpsschoolForm'
  }];
  SpreadsheetApp.getActive().addMenu('OpsschoolForm', menu);
}


/**
 * A set-up function that uses the OpsschoolForm data in the spreadsheet to create
 * Google Calendar events, a Google Form, and a trigger that allows the script
 * to react to form responses.
 */
function setUpOpsschoolForm() {
  /* if (ScriptProperties.getProperty('calId')) {
     Browser.msgBox('Your OpsschoolForm is already set up. Look in Google Drive!');
   }*/
  deleteUnusedFormSheets();
  var ss = SpreadsheetApp.getActive();
  var values = ss.getSheetByName('opschool_exam').getDataRange().getValues();;
  //  var range = sheet.getDataRange();
  //var values = range.getValues();

  //setUpCalendar_(values, range);
  setUpForm_(ss, values);
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit()
    .create();
  //  ss.removeMenu('OpsschoolForm');
}

/**
 * Creates a Google Form that allows respondents to select which OpsschoolForm
 * sessions they would like to attend, grouped by date and start time.
 *
 * @param {Spreadsheet} ss The spreadsheet that contains the OpsschoolForm data.
 * @param {Array<String[]>} values Cell values for the spreadsheet range.
 */
function setUpForm_(ss, values) {

  var form = FormApp.create('OpsSchool-exam');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  form.addTextItem().setTitle('Name').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true).setValidation(FormApp.createTextValidation().requireTextIsEmail().build());


  for (var i = 1; i < values.length; i++) {
    var question = values[i];
    var q = question[0];
    var an1 = question[1];
    var an2 = question[2];
    var an3 = question[3];
    var an4 = question[4];
    // Browser.msgBox(an1 + an2 + an3 + an4);s

    var header = form.addSectionHeaderItem().setTitle('Question' + i);
    var item = form.addMultipleChoiceItem()
    item.setTitle(q)
      .setChoices([
        item.createChoice(an1),
        item.createChoice(an2),
        item.createChoice(an3),
        item.createChoice(an4)
      ])
      .showOtherOption(false);


  }
}

/**
 * A trigger-driven function that sends out calendar invitations and a
 * personalized Google Docs itinerary after a user responds to the form.
 *
 * @param {Object} e The event parameter for form submi§ion to a spreadsheet;
 *     see https://developers.google.com/apps-script/understanding_events
 */
function onFormSubmit(e) {
  var user = {
    name: e.namedValues["Name"][0],
    email: e.namedValues["Email"][0],
    correct_answers: {
      devops: 0,
      system: 0
    },
    devops_score: 0,
    system_score: 0,
    pass: false
  };

  var questions_count = {
    devops: 0,
    system: 0,
    total: 0
  }

  // Grab the session data again so that we can match it to the user's choices.
  var exam_values = SpreadsheetApp.getActive().getSheetByName('opschool_exam')
    .getDataRange().getValues();
  var results_ss = SpreadsheetApp.getActive().getSheetByName('opschool_exam_results');
  var results_values = results_ss.getDataRange().getValues();

  // last_row = results_ss.getLastRow()
  // results_ss.getRange("A1").setValue(last_row)

  // results_ss.appendRow([user.name, user.email, user.devops_score, user.system_score, user.pass]);
  // results_ss.appendRow([user.name, user.email, user.devops_score, user.system_score, user.pass]);


  for (var i = 1; i < exam_values.length; i++) {
    var question_data = exam_values[i];
    question = question_data[0]
    question_related_field = question_data[6]
    correct_answer = question_data[5]
    // results_ss.appendRow(["QUESTION", question, question_related_field, correct_answer]);

    questions_count[question_related_field]++
    questions_count.total = questions_count.total + 1;

    // results_ss.appendRow(["hash", questions_count["devops"], questions_count["system"], questions_count["total"], questions_count.total]);


    if (e.namedValues[question] == correct_answer) {
      user.correct_answers[question_related_field] = user.correct_answers[question_related_field] + 1;
    }
  }
  user.devops_score = user.correct_answers.devops / questions_count.devops
  user.system_score = user.correct_answers.system / questions_count.system
  if (user.devops_score > 0.4 && user.system_score > 0.4) {
    user.pass = true
  }
  results_ss.appendRow([user.name, user.email, user.devops_score, user.system_score, user.pass]);
  if (user.pass) {
    body_message = 'Thanks for taking the OpsSchool test! We\'re happy to inform you that you had passed the test. Here is the link for the second assignment'
  } else {
    body_message = 'Thanks for taking the OpsSchool test! Unfortunately you didn\'t pass the test'
  }
  MailApp.sendEmail({
    to: user.email,
    subject: user.name,
    body: body_message
    // attachments: doc.getAs(MimeType.PDF)
  });
}