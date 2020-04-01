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

var request_notification_email = 'smorgan@bridgend.ac.uk';

var resources_available = [
  'Discussion with Learning Technologist',
  'Training session',
  'Virtual reality headsets - Oculus Go',
  'Virtual reality headsets - Oculus Quest',
  'Indoor drones',
  'Outdoor drone',
  'Raspberry Pi',
  'Micro:bits',
  'Robots',
  '360° camera',
  '3D scanner',
  'DSLR camera',
  'Swivl lecture capture',
  '3D printing',
];
  
var campuses = [
  'Cowbridge Road',
  'Pencoed',
  "Queen's Road",
  "Maesteg",
];

var available_locations = [
  'Den01',
  'Own space',
];

// Form field titles, used for creating the form and as keys when handling
// responses.
/**
 * Add custom menu items when opening the sheet.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Equipment requests')
      .addItem('Set up', 'setup_')
      .addItem('Clean up', 'cleanup_')
      .addToUi();
}

/**
 * Set up the form and triggers for the workflow.
 */
function setup_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getFormUrl()) {
    var msg = 'Form already exists. Unlink the form and try again.';
    SpreadsheetApp.getUi().alert(msg);
    return;
  }
  var form = FormApp.create('Resources Request Form - Apps Script Version')
      .setCollectEmail(true)
      .setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId())
      .setLimitOneResponsePerUser(false);
  form.addListItem().setTitle('Campus').setRequired(true).setChoiceValues(campuses);
  form.addListItem().setTitle('Location').setRequired(true).setChoiceValues(available_locations);
  form.addDateTimeItem().setTitle('Date & time required').setRequired(true);
  //form.addTimeItem().setTitle('Time required').setRequired(true);
  var item = form.addMultipleChoiceItem();
  item.setTitle('Would you like assistance from a learning technologist before/while using the resource?')
      .setChoices([
        item.createChoice('Yes'),
        item.createChoice('No')
       ]);

  // Section 2
  form.addParagraphTextItem().setTitle('Please explain what exactly you would like to be able to do and one of the Learning Technologists will be in touch as soon as possible.').setRequired(true);
  
  // Section 3
  form.addSectionHeaderItem().setTitle('You have requested a training session. Please use the specific training request form, available here: https://forms.gle/w9iPVA23VPJJ4LYMA');

  // Hide the raw form responses.
  ss.getSheets().forEach(function(sheet) {
    if (sheet.getFormUrl() == ss.getFormUrl()) {
      sheet.hideSheet();
    }
  });
  // Start workflow on each form submit
  ScriptApp.newTrigger('onFormSubmit_')
      .forForm(form)
      .onFormSubmit()
      .create();
  // Archive completed items every 5m.
  ScriptApp.newTrigger('processCompletedItems_')
      .timeBased()
      .everyMinutes(5)
      .create();
}

/**
 * Cleans up the project (stop triggers, form submission, etc.)
 */
function cleanup_() {
  var formUrl = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
  if (!formUrl) {
    return;
  }
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });
  FormApp.openByUrl(formUrl)
      .deleteAllResponses()
      .setAcceptingResponses(false);
}

/**
 * Handle new form submissions to trigger the workflow.
 *
 * @param {Object} event - Form submit event
 */
function onFormSubmit_(event) {
  var response = mapResponse_(event.response);
  sendNewEquipmentRequestEmail_(response);
  var equipmentDetails = Utilities.formatString('%s\n%s\n%s',
      response['Laptop'],
      response['Desktop'],
      response['Monitor']);
  var row = ['New',
    '',
    response['Desk location'],
    response['Employee name'],
    response['Desk location'],
    equipmentDetails,
    response['email']];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Pending requests');
  sheet.appendRow(row);
}

/**
 * Sweeps completed events, notifying the requestors and archiving them
 * to the completed
 *
 * @param {Object} event
 */
function processCompletedItems_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pending = ss.getSheetByName('Pending requests');
  var completed = ss.getSheetByName('Completed requests');
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

/**
 * Sends email notifying team a new equipment request has been submitted.
 *
 * @param {Object} request - Request details
 */
function sendNewEquipmentRequestEmail_(request) {
  var template = HtmlService.createTemplateFromFile('new-equipment-request.html');
  template.request = request;
  template.sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var msg = template.evaluate();
  MailApp.sendEmail({
    to: REQUEST_NOTIFICATION_EMAIL,
    subject: 'New equipment request',
    htmlBody: msg.getContent(),
  });
}

/**
 * Sends email notifying requestor that the equipment has been provided.
 *
 * @param {Object} request - Request details
 */
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