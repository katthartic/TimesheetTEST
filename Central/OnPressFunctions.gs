// @ts-nocheck

/**
 * Enters pay period total in cover summary.
 * Emails summary of timesheet
 *
 * @return none
 * @customfunction
 */
function approveTimesheet(ss, sheet) {
  // const ss = SpreadsheetApp.getActiveSpreadsheet()

  //check timesheet is accurate
  _checkTimesheet(ss,sheet)

  //get pay period from timesheet
  const payPeriod = ss.getSheetByName('Timesheet').getRange('C2').getValue()

  //get total from timesheet
  const total = ss.getSheetByName('Timesheet').getRange('D24').getValue()

  //set total for pay pay period on cover sheet
  const tableArray = ss.getRangeByName('CoverPayPeriod').getValues()
  const row = tableArray.findIndex(item => item[0] === payPeriod) + 1
  ss.getRangeByName('CoverPayPeriod').getCell(row,2).setValue(total)

  //email timesheet
  const name = ss.getSheetByName('Cover').getRange('B2').getValue()
  const emailAddress = ss.getSheetByName('Cover').getRange('B3').getValue()
  // const approverEmailAddresses = 'sara@bax.org, ashley@bax.org'
  const approverEmailAddress = ss.getRangeByName('ApproverEmail').getValue()
  const subject = `BAX Timesheet: ${name} submitted for ${payPeriod}`
  const filteredSummary = ss.getRangeByName('TimesheetSummary').getValues().filter(row => row[2])
        //row has hours data
  const body = _createEmailBody(filteredSummary)

  MailApp.sendEmail({
    to: emailAddress,
    cc: approverEmailAddress,
    subject,
    body
  })
}

/**
 * Converts an array of timesheet summary rows into a string.
 * Return string as body of an email
 *
 * @param {summaryArray} the array of timesheet summary rows
 * @return body of email as string
 * @customfunction
 */
function _createEmailBody(summaryArray) {
  let body = "SUMMARY OF TIMESHEET\n\n"

  summaryArray.forEach(row => {
    const [dept, jobType, hours, pay] = row
    const roundHours = Math.round(hours * 100) / 100
    const roundPay = Math.round(pay * 100) / 100

    if(jobType !== 'TOTAL'){
      body += `${dept} - ${jobType}\n$${roundPay} for ${roundHours} ${roundHours > 1 ? 'hours' : 'hour'}\n\n`
    } else {
      body += `${jobType}\n$${roundPay} for ${roundHours} ${roundHours > 1 ? 'hours' : 'hour'}\n\n`
    }
  })

  return body
}

/**
 * Checks if rate or time/session is missing
 * Sends alert
 *
 * @return none
 * @customfunction
 */
function _checkTimesheet(ss,sheet){
  const timesheetHoursTotal = ss.getRangeByName('TimesheetHoursTotal').getValues()
  
  for(i = 0; i < timesheetHoursTotal.length; i++){
    const row = timesheetHoursTotal[i]
    if((row[0] > 0 && row[1] <=0) || (row[1] > 0 && row[0] <=0)){
      const ui = SpreadsheetApp.getUi()
      const issueRow = i + 1
      const rateCell = sheet.getRange(rateColNum,issueRow)

      if(rateCell.getValue() <= 0) {
        ui.alert(`Please enter rate on row ${issueRow}.`)
      } else {
        ui.alert(`Please enter time or session on row ${issueRow}.`)
      }
      break
    }
  }
}
