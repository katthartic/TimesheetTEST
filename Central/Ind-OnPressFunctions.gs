// @ts-nocheck

/**
 * Activated on press of submit button
 * Runs approveTimesheet from TimeSheetScriptMain
 *
 * @return none
 * @customfunction
 */
function onSubmit(){
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getActiveSheet()

  TimesheetScriptMAIN.approveTimesheet(ss,sheet)
}
