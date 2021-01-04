// @ts-nocheck

/** Variables*/
const deptColNum = 7
const typeColNum = 8
const rateColNum = 10
const startTimeColNum = 11
const sessionColNum = 13
const editableBackground = '#d9ead3'
const uneditableBackground = '#d8d8d8'

/**
 * Sets data validation dropdown in cell(s) based on list.
 *
 * @param {range} the cell(s) the dropdown appears in
 * @param {souceRange} the list for the dropdown
 * @return none
 * @customfunction
 */
function _setDataValid(range, sourceRange) {
  //sets data validation for job types dropdown
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(sourceRange, true).build();
  range.setDataValidation(rule);
}

/**
 * Returns rate for given job dept and job type.
 *
 * @param {dept} the dept of job
 * @param {type} the type of job
 * @return number representation of rate based on job dept and type
 * @customfunction
 */
function _findRate(dept, type){

  const rateTable = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('CoverJobTable')
  const rate = rateTable.getValues().find(item => {
    const [itemdept, itemType, itemRate] = item
    return (itemdept === dept && itemType === type)
  })[2]

  if(rate) return rate
  else return null

}

/**
 * Activated when selections are made in the Timesheet Dept column and Job Type column
 * Runs _setDataValid on edit in Dept column
 * Runs _findRate on edit in Job Type column
 */
function onEdit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const sheetName = ss.getSheetName()
  // const aCell = ss.getActiveCell()
  // const acCol = aCell.getColumn()
  // const acRow = aCell.getRow()

  const selection = ss.getSelection()
  const activeRange = selection.getActiveRange()
  const numCols = activeRange.getNumColumns()
  const numRows = activeRange.getNumRows()

  //loop through all cells in selected range
  for (let i = 1; i <= numCols; i++) {
    for (let j = 1; j <= numRows; j++) {
      const aCell = activeRange.getCell(j, i)
      const acCol = aCell.getColumn()
      const acRow = aCell.getRow()
      const value = aCell.getValue()
      
      if(value && sheetName === 'Timesheet'){
        
        //if dept is set... fetches and sets type dropdown
        if(acCol === deptColNum && acRow > 1){
          const range = ss.getRange(acRow,typeColNum)
          const sourceRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(value)
          _setDataValid(range,sourceRange)
        }
        
        //if type is set... fetches and sets rate
        if(acCol === typeColNum && acRow > 1) {
          const entereddept = ss.getRange(acRow,deptColNum).getValue()
          const enteredType = value
          const rate = _findRate(entereddept,enteredType)
          ss.getRange(acRow,rateColNum).setValue(rate)
        }

        if(acCol === typeColNum && (value === 'AIE' || value === 'Stipend')){
          const timeRange = ss.getRange(acRow,startTimeColNum,1,2)
          timeRange.setBackground(uneditableBackground)
          const sessionRange = ss.getRange(acRow,sessionColNum)
          sessionRange.setValue(1)
        }
    
      } else {
        //if dept is cleared... clears type dropdown
        if(acCol === deptColNum){
        ss.getRange(acRow,typeColNum).setDataValidation(null)
        }

        //if type is cleared... clears rate, session and resets background color
        if(acCol === typeColNum){
        ss.getRange(acRow,rateColNum).setValue(null)
        ss.getRange(acRow,sessionColNum).setValue(null)
        ss.getRange(acRow,startTimeColNum,1,2).setBackground(editableBackground)
        }
      }
    }
  }
  
  if(e.range.getColumn() === typeColNum && (e.oldValue === 'AIE' || e.oldValue === 'Stipend')){
    const eRow = e.range.getRow()
    ss.getRange(eRow,sessionColNum).setValue(null)
    ss.getRange(eRow,startTimeColNum,1,2).setBackground(editableBackground)
  }
}

function test() {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const sheetName = ss.getSheetName()

  const selection = ss.getSelection()
  const activeRange = selection.getActiveRange()
  const numCols = activeRange.getNumColumns()
  const numRows = activeRange.getNumRows()

  for (let i = 1; i <= numCols; i++) {
    for (let j = 1; j <= numRows; j++) {
      const aCell = activeRange.getCell(j, i)
      const acCol = aCell.getColumn()
      const acRow = aCell.getRow()

    }
  }
}

