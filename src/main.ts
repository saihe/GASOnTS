const dateCell = 'B'
const dayCell = 'C'
const startHoursCell = 'D'
const startMinutesCell = 'F'
const endHoursCell = 'G'
const endMinutesCell = 'I'
const restHoursCell = 'J'
const restMinutesCell = 'L'
const descriptionCell = 'Q'
const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土']

function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  if (e.range.getRow() !== 5 || e.range.getColumn() !== 2 && e.range.getColumn() !== 5) {
    return
  }

  if (Browser.msgBox('作業報告書を初期化しますか？', Browser.Buttons.OK_CANCEL) === 'cancel') {
    return
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('作業報告書')
  if (!sheet) {
    return
  }

  init(sheet)
}

function init(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  const year = sheet.getRange('B5').getValue()
  const month = sheet.getRange('E5').getValue()
  const defaultStart = sheet.getRange('C9').getDisplayValue().toString().split(':')
  const defaultEnd = sheet.getRange('E9').getDisplayValue().toString().split(':')
  const defaultRest = sheet.getRange('G9').getDisplayValue().toString().split('.')
  const lastOfMonth = new Date(year, month, 0)

  for(var row = 12; row < 43; row++) {
    const date = row - 11
    if (date > lastOfMonth.getDate()) {
      sheet.getRange(`${dateCell}${row}`).setValue('')
      sheet.getRange(`${dayCell}${row}`).setValue('')
      sheet.getRange(`${startHoursCell}${row}`).setValue('')
      sheet.getRange(`${startMinutesCell}${row}`).setValue('')
      sheet.getRange(`${endHoursCell}${row}`).setValue('')
      sheet.getRange(`${endMinutesCell}${row}`).setValue('')
      sheet.getRange(`${restHoursCell}${row}`).setValue('')
      sheet.getRange(`${restMinutesCell}${row}`).setValue('')
      sheet.getRange(`${descriptionCell}${row}`).setValue('')
      continue
    }

    const currentDates = new Date(year, month-1, date)
    sheet.getRange(`${dateCell}${row}`).setValue(date)
    sheet.getRange(`${dayCell}${row}`).setValue(dayOfWeek[currentDates.getDay()])
    if (this.JapaneseHolidays.isHoliday(currentDates) || currentDates.getDay() === 0 || currentDates.getDay() === 6) {
      sheet.getRange(`${startHoursCell}${row}`).setValue('')
      sheet.getRange(`${startMinutesCell}${row}`).setValue('')
      sheet.getRange(`${endHoursCell}${row}`).setValue('')
      sheet.getRange(`${endMinutesCell}${row}`).setValue('')
      sheet.getRange(`${restHoursCell}${row}`).setValue('')
      sheet.getRange(`${restMinutesCell}${row}`).setValue('')
    } else {
      sheet.getRange(`${startHoursCell}${row}`).setValue(defaultStart[0])
      sheet.getRange(`${startMinutesCell}${row}`).setValue(defaultStart[1])
      sheet.getRange(`${endHoursCell}${row}`).setValue(defaultEnd[0])
      sheet.getRange(`${endMinutesCell}${row}`).setValue(defaultEnd[1])
      sheet.getRange(`${restHoursCell}${row}`).setValue(defaultRest[0])
      sheet.getRange(`${restMinutesCell}${row}`).setValue(defaultRest[1])
    }
    sheet.getRange(`${descriptionCell}${row}`).setValue(
      (this.JapaneseHolidays.isHoliday(currentDates) && currentDates.getDay() !== 0 && currentDates.getDay() !== 6) ? '祝日' : ''
    )
  }
}
