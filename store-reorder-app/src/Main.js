/**
 * Main.js
 * Weekly pipeline entry point + trigger installer.
 * Run installWeeklyTrigger() once from the Apps Script editor after
 * CONFIG.SPREADSHEET_ID / IMPORT_FOLDER_ID are filled in.
 */

function runWeeklyPipeline() {
  ownerOnly_();
  importAllFiles();
  recomputeMinMax();
  buildReorderQueue();
  sendReorderEmails();
  logActivity_('runWeeklyPipeline', 'ok', 'weekly run complete');
}

function installWeeklyTrigger() {
  ownerOnly_();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runWeeklyPipeline' || t.getHandlerFunction() === 'remindUnansweredReorders') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('runWeeklyPipeline').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  ScriptApp.newTrigger('remindUnansweredReorders').timeBased().everyDays(1).atHour(9).create();
}

function logActivity_(step, status, message) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVITY_LOG) || ss.insertSheet(CONFIG.SHEETS.ACTIVITY_LOG);
  if (sheet.getLastRow() === 0) sheet.appendRow(['timestamp', 'step', 'status', 'message']);
  sheet.appendRow([new Date(), step, status, message]);
}
