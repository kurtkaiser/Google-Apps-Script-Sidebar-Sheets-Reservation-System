// Reservation Application
// Kurt Kaiser
// kurtkaiser.us
// CC0 / Public Domain
// Tutorial: youtu.be/3ms0YrGMuls

// Declare global variables
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var lastRow = sheet.getLastRow();
var lastColumn = sheet.getLastColumn();
var scriptProperties = PropertiesService.getScriptProperties();
var letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", 
  "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
  
// -------- Sidebar HTML Setup Functions --------
  
// Calendar Variables
var calendars = CalendarApp.getAllOwnedCalendars();

// As the spreadsheet opens add a menu
function onOpen() {
  var ui = SpreadsheetApp.getUi(); 
  ui.createMenu('Reservation App')
    .addItem('Show Sidebar ', 'showFormSidebar')
    .addToUi();
}

function showFormSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Form')
    .setTitle('Application Control')
    .setWidth(300);
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function getAllHeaders(){
  var allItems = sheet.getRange(1, 1, 1, lastColumn).getValues();
  return createDropMenu(allItems);
}

// Get first row of spreadsheet, return to html for drop menu
function createDropMenu(allItems){
  var drop = "";
  for (let i = 0; i < lastColumn; i++){
    drop += '<option value="' + i + '">'  + getLetter(i) + allItems[0][i] + '</option>';
  }
  return drop;
}

// Get all of your account calendars, return to html for drop menu
function getCalendars(){
  var drop = "";
  for (let i = 0; i < calendars.length; i++){
    drop += '<option value="' + i + '">'  + calendars[i].getName() + '</option>';
    if(calendars[i].getId() == scriptProperties.getProperty('calId')){
      scriptProperties.setProperty('calendarIndex', i);
    }
  }
  return drop;
}

// Returns array of saved properties, start date column, client name column, etc
// This is used to populate the sidebar with previously saved data
function getSavedPropsForSidebar() {
  var propertiesAndKeys = {}
  var data = scriptProperties.getProperties();
  for (var key in data) {
    propertiesAndKeys[key] = scriptProperties.getProperty(key);
  }
  return propertiesAndKeys;
}

// Get column letters, check if two letters needed, past z of columns
function getLetter(index){
  var temp = letters[index] + ': ';
  if(index > 25) {
    // Double letter name
    temp = letters[Math.floor(index/26)-1] + letters[index%26]+': ';
  }
    // If you have over 702 columns, into triple letter names ACP, blank
  if (index > 702) temp = "";
    return temp;
} 

// -------- Save Submitted Sidebar Data --------

// User clicked submit, save all info to properties
function saveSidebar(sideData) {
   scriptProperties.deleteAllProperties();
   scriptProperties.setProperty('startIndex', sideData.startDate);
   scriptProperties.setProperty('endIndex', sideData.endDate);
   scriptProperties.setProperty('clientNameIndex', sideData.clientName);
   scriptProperties.setProperty('clientEmailIndex', sideData.clientEmail);
   scriptProperties.setProperty('reservedMsg', sideData.reservedMsg);
   scriptProperties.setProperty('conflictMsg', sideData.conflictMsg);
   scriptProperties.setProperty('calendarIndex', sideData.calIndex);
   scriptProperties.setProperty('calId', calendars[sideData.calIndex].getId());
}

// -------- Client Submits Form / Booking Request  --------

// Runs automatically once Form is submitted for a potential booking
function onFormSubmission(){
  lastRow = sheet.getLastRow();
  var entireRow = sheet.getRange(lastRow, 1, 1, lastColumn);
  // Get all info from the spreadsheet row (last) that was just submitted
  var allValues = entireRow.getValues();
  allValues = allValues[0];
  // Create object to store potential booking data
  var submission = {
    start: new Date(allValues[scriptProperties.getProperty('startIndex')]),
    end: new Date(allValues[scriptProperties.getProperty('endIndex')]),
    clientName: allValues[scriptProperties.getProperty('clientNameIndex')],
    clientEmail: allValues[scriptProperties.getProperty('clientEmailIndex')],
    emailMsg: "Request recieved.", //placeholder message
    calendar: CalendarApp.getCalendarById(scriptProperties.getProperty('calId')),
    status: "Recieved", //placeholder message
    lastColumn: allValues.filter(String).length + 1
  }
  submission.end.setHours(submission.end.getHours()+12);
  checkDates(submission);
}

// Check if the requested days are available
function checkDates(submission){
  var conflict = submission.calendar.getEvents(submission.start, submission.end);
  // If conflicts < 1, no bookings, start reservation process
  if(conflict.length < 1){
    reserveDays(submission);
 } else { 
    sheet.getRange(lastRow, submission.lastColumn).setValue('Conflict');
    submission.emailMsg = scriptProperties.getProperty('conflictMsg');
    submission.status = "Conflict";
    emailSend(submission);
 }
}

// Days are aviable, create the calendar event
function reserveDays(submission){
  var event = submission.calendar.createEvent
      (submission.clientName, submission.start, submission.end);
  // if event is successfully created...
  if(event){
      sheet.getRange(lastRow, submission.lastColumn).setValue('Reserved');
      submission.emailMsg = scriptProperties.getProperty('reservedMsg');
      submission.status = "Confirmed";
  } else {
      // Else there has been an error
      sheet.getRange(lastRow, submission.lastColumn).setValue('Issue');
  }
  emailSend(submission);
}

// Send the email notification of reservation status
function emailSend(submission) {
 var testing = " ";
 if(new Date() < new Date(2020, 06, 01)) testing += Math.random();
  var htmlEmail = getHtmlFile(submission);
  MailApp.sendEmail({
    to: submission.clientEmail,
    subject: "Reservation " + submission.status + testing,
    htmlBody: htmlEmail
  })
}

// Uses the Email.html file and our data to create an email object
function getHtmlFile(submission){
  var startStr = (submission.start.getMonth()+1) + '/' + 
    submission.start.getDate() + '/' + (submission.start.getYear() + 1900);
  var endStr = (submission.end.getMonth()+1) + '/' + 
    submission.end.getDate() + '/' + (submission.end.getYear()+ 1900);
  var htmlEmail = HtmlService.createTemplateFromFile('Email');
  htmlEmail.messageTitle = submission.status;
  htmlEmail.messageBody = submission.emailMsg;
  htmlEmail.startDate = startStr;
  htmlEmail.endDate = endStr;
  htmlEmail.clientName = submission.clientName;
  htmlEmail = htmlEmail.evaluate();
  htmlEmail = htmlEmail.getContent();
  return htmlEmail;
}

