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
    var an5 = question[5];
    // Browser.msgBox(an1 + an2 + an3 + an4);s

    var header = form.addSectionHeaderItem().setTitle('Question' + i);
    var item = form.addMultipleChoiceItem()
    item.setTitle(q)
      .setChoices([
        item.createChoice(an1),
        item.createChoice(an2),
        item.createChoice(an3),
        item.createChoice(an4),
        item.createChoice(an5)
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
  var cnst_other_question = "I am not familiar with this topic";
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
  var devops_overqualified_result = 0.4;
  var system_underqualified_result = 0.4;

  // Grab the session data again so that we can match it to the user's choices.
  var param_values = SpreadsheetApp.getActive().getSheetByName('params')
  var exam_values = SpreadsheetApp.getActive().getSheetByName('opschool_exam')
    .getDataRange().getValues();
  var results_ss = SpreadsheetApp.getActive().getSheetByName('opschool_exam_results');
  var results_values = results_ss.getDataRange().getValues();

  // get constants from params sheet
  for (var i = 0; i < param_values.length; i++) {
    var param_data = param_values[i];
    switch (param_data[0]) {
      case "pass_message":
        pass_message = param_data[1];
        break;
      case "fail_message":
        fail_message = param_data[1];
        break;
      case "devops_overqualified_result":
        devops_overqualified_result = param_data[1];
        break;
      case "system_underqualified_result":
        system_underqualified_result = param_data[1];
        break;
    }


  }
  // calculate all answers from attendees
  for (var i = 1; i < exam_values.length; i++) {
    var question_data = exam_values[i];
    question = question_data[0]
    question_related_field = question_data[7]
    correct_answer = question_data[6]

    questions_count[question_related_field]++;
    questions_count.total++;
    // questions_count.total = questions_count.total + 1;


    if (e.namedValues[question] == correct_answer) {
      // user.correct_answers[question_related_field] = user.correct_answers[question_related_field] + 1;
      user.correct_answers[question_related_field]++;
    }

    if (e.namedValues[question] == cnst_other_question) {
      user.correct_answers[question_related_field] = user.correct_answers[question_related_field] + 0.25;

    }
  }
  user.devops_score = user.correct_answers.devops / questions_count.devops
  user.system_score = user.correct_answers.system / questions_count.system
  if (user.devops_score > 0.4 && user.system_score > 0.4) {
    user.pass = true
  }


  switch (true) {
    case user.devops_score > devops_overqualified_result:
      user.pass = false
      user.pass_reason = "devops-overqualified"
      break;
    case user.system_score < system_underqualified_result:
      user.pass = false
      user.pass_reason = "system-underqualified"
      break;
    case user.system_score >= system_underqualified_result:
      user.pass = true
      user.pass_reason = "system-underqualified"
      break;
  }

  results_ss.appendRow([user.name, user.email, user.devops_score, user.system_score, user.pass, user.pass_reason]);

  if (user.pass) {
    body_message = pass_message
  } else {
    body_message = fail_message
  }
  MailApp.sendEmail({
    to: user.email,
    subject: user.name,
    body: body_message
    // attachments: doc.getAs(MimeType.PDF)
  });
}