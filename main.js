// Copyright 2019 Google LLC
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     https://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
// Modifications made be Scott Morgan <smorgan@bridgend.ac.uk>

var SPREADSHEET_ID = '19sL72pKG0sCKR6v1HQ7JQLFBT4iJ4UHnFXM78QUu_D8';
var REQUEST_NOTIFICATION_EMAIL = 'smorgan@bridgend.ac.uk';

var CAMPUSES = [
  'Cowbridge Road',
  'Pencoed',
  "Queen's Road",
  "Maesteg",
];

var AVAILABLE_LOCATIONS = [
  'Den01',
  'Own space',
];

// Form field titles, used for creating the form and as keys when handling
// responses.
/**
 * Add custom menu items when opening the sheet.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Resources request')
      .addItem('Initialise', 'setup_')
      .addItem('Update', 'update_')
      .addToUi();
}

/**
 * Set up the form and triggers for the workflow.
 */
function setup_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (ss.getFormUrl()) {
    var msg = 'Form already exists. Try running update instead.';
    SpreadsheetApp.getUi().alert(msg);
    return;
  }
  var form = FormApp.create('Resources Request Form')
      .setCollectEmail(true)
      .setDestination(FormApp.DestinationType.SPREADSHEET, SPREADSHEET_ID)
      .setLimitOneResponsePerUser(false);
  form.addListItem().setTitle('Campus').setRequired(true).setChoiceValues(CAMPUSES);
  form.addListItem().setTitle('Location').setRequired(true).setChoiceValues(AVAILABLE_LOCATIONS);
  form.addDateTimeItem().setTitle('Date & time required').setRequired(true);

  // Adds multiple choice question, assistance question and new sections
  var item_resource = form.addMultipleChoiceItem();
  var item_lt = form.addListItem();
  var sectionExplain = form.addPageBreakItem();
  form.addParagraphTextItem()
      .setTitle('Please explain what exactly you would like to be able to do and one of the Learning Technologists will be in touch as soon as possible.')
      .setRequired(true)
  var sectionRedirect = form.addPageBreakItem();
  form.addSectionHeaderItem().setTitle('You have requested a training session. Please use the specific training request form, available here: https://forms.gle/w9iPVA23VPJJ4LYMA');
  var sectionUnused = form.addPageBreakItem();

  sectionExplain.setGoToPage(FormApp.PageNavigationType.CONTINUE); // Required explicitly
  sectionRedirect.setGoToPage(FormApp.PageNavigationType.SUBMIT);
  sectionUnused.setGoToPage(FormApp.PageNavigationType.RESTART);

  // Sets up resources list
  setResourceList_(ss,item_resource,sectionExplain,sectionRedirect)

  item_lt.setTitle('Would you like assistance from a learning technologist before/while using the resource?')
         .setChoiceValues(['Yes','No'])
         .setRequired(true);

  // Start workflow on each form submit
  ScriptApp.newTrigger('onFormSubmit_')
      .forForm(form)
      .onFormSubmit()
      .create();

  // Archive completed items every 5m.
  //ScriptApp.newTrigger('processCompletedItems_')
  //    .timeBased()
  //    .everyMinutes(5)
  //    .create();
}

/**
 * Cleans up the project (stop triggers, form submission, etc.)
 */
function update_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID)
  var formUrl = ss.getFormUrl();
  if (!formUrl) {
    return;
  }
  var title = 'Do you know what this does?';
  var msg = 'This script will update the resources and quantities in the from with data from the \'Available\' tab.';
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(title,msg,ui.ButtonSet.OK_CANCEL);

  if (response == ui.Button.OK) {
    var form = FormApp.openByUrl(formUrl);

    var items = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE)
    var sections = form.getItems(FormApp.ItemType.PAGE_BREAK)

    var item_resource = items[0].asMultipleChoiceItem();
    var sectionExplain = sections[0].asPageBreakItem();
    var sectionRedirect = sections[1].asPageBreakItem();

    setResourceList_(ss,item_resource,sectionExplain,sectionRedirect)
  } else {
    return;
  }

}

function setResourceList_(ss,item_resource,sectionExplain,sectionRedirect) {

  // Sets up resources list
  item_resource.setTitle('Which Den01 resources would you like to request?').setRequired(true);
  item_resource.createChoice('Training session',sectionRedirect);

  var sheet_available = ss.getSheetByName('Available')
  var data = sheet_available.getRange(1, 1, sheet_available.getLastRow() - 1, 2).getValues();

  var choiceArray = [];
  choiceArray.push(item_resource.createChoice('Training session',sectionRedirect))
  for (var i=0; i < data.length; i++){
    var resourceInd = data[i][0]; // resource
    var numAvailableInd = +data[i][1]; // number available

    if (numAvailableInd == 0) {
      quString = resourceInd;
    } else {
      quString = resourceInd + " (" + numAvailableInd + " available)";
    }

    choiceArray.push(item_resource.createChoice(quString,sectionExplain));

  }

  item_resource.setChoices(choiceArray);
}

/**
 * Sends email notifying team a new equipment request has been submitted.
 *
 * @param {Object} request - Request details
 */
function sendNewEquipmentRequestEmail_(request) {
  var template = HtmlService.createTemplateFromFile('resource-request.html');
  template.request = request;
  template.sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var msg = template.evaluate();
  MailApp.sendEmail({
    to: REQUEST_NOTIFICATION_EMAIL,
    subject: 'New resource request',
    htmlBody: msg.getContent(),
  });
}

/**
 * Creates calendar event in Den01 calendar
 */
function createCalendarEvent(response) {

  var calendar = CalendarApp.getCalendarById('bridgend.ac.uk_f9tbdustnbeoqscf21tcf747t4@group.calendar.google.com');

  var eventName = response['Which Den01 resources would you like to request?'];
  var startDateTime = new Date(response['Date & time required']);
  var endDateTime = new Date(startDateTime + 60*60*1000);
  var eventDescription = 'Request sent by: ' + response.email + '\n\nRequest: \n' + response['Please explain what exactly you would like to be able to do and one of the Learning Technologists will be in touch as soon as possible.'];

  calendar.createEvent(eventName,startDateTime,endDateTime,{description: eventDescription});
}

/**
 * Handle new form submissions to trigger the workflow.
 *
 * @param {Object} event - Form submit event
 */
function onFormSubmit_(event) {
  var response = mapResponse_(event.response);
  sendNewEquipmentRequestEmail_(response);
  createCalendarEvent(response);
  /**
  var equipmentDetails = Utilities.formatString('%s\n%s\n%s',
      response['Laptop'],
      response['Desktop'],
      response['Monitor']);
  var row = ['New',
    '',
    response['Due Date'],
    response['Employee name'],
    response['Desk location'],
    equipmentDetails,
    response['email']];
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Pending');
  sheet.appendRow(row);
  */
}

/**
 * Converts a form response to an object keyed by the item titles. Allows easier
 * access to response values.
 *
 * @param {FormResponse} response
 * @return {Object} Form values keyed by question title
 */
function mapResponse_(response) {
  var initialValue = {
    email: response.getRespondentEmail(),
    timestamp: response.getTimestamp(),
  };
  return response.getItemResponses().reduce(function(obj, itemResponse) {
    var key = itemResponse.getItem().getTitle();
    obj[key] = itemResponse.getResponse();
    return obj;
  }, initialValue);
}

/**
 * Sweeps completed events, notifying the requestors and archiving them
 * to the completed
 *
 * @param {Object} event
 */
/**
function processCompletedItems_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var pending = ss.getSheetByName('Pending');
  var completed = ss.getSheetByName('Completed');
  var rows = pending.getDataRange().getValues();
  rows.forEach(function(row, index) {
    var status = row[0];
    if (status === 'Completed' || status == 'Cancelled') {
      pending.deleteRow(index + 1);
      completed.appendRow(row);
      sendEquipmentRequestCompletedEmail_({
        'Employee name': row[3],
        'Desk location': row[4],
        'email': row[6],
      });
    }
  });
}
*/

/**
 * Sends email notifying requestor that the equipment has been provided.
 *
 * @param {Object} request - Request details
 */
 /**
function sendEquipmentRequestCompletedEmail_(request) {
  var template = HtmlService.createTemplateFromFile('request-complete.html');
  template.request = request;
  var msg = template.evaluate();
  MailApp.sendEmail({
    to: request.email,
    subject: 'Equipment request completed',
    htmlBody: msg.getContent(),
  });
}
*/
