// @ts-nocheck
/**
 * Activated on edit of spreadsheet
 * Runs onEditCalled from TimeSheetScriptMain
 */
function onEdit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getActiveSheet()
  // const sheetName = sheet.getSheetName()
  const selection = sheet.getSelection()
  const activeRange = selection.getActiveRange()

  TimesheetScriptMAIN.onEditCalled(ss, sheet, activeRange, e)
}
