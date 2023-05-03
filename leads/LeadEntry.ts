// function sheetProtection() {
//     let name = "leads_management"
//     let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name.toLowerCase())
//     let row = sheet.getActiveCell().getRow()
//     let col = sheet.getActiveCell().getColumn()
//     let range = sheet.getRange(row, col)
//     let value = range.getValue()
//     let items = ["OPEN", "USELESS", "WON", "WON DEALER", "LOST", , "POTENTIAL"]

//     if (value != "") {
//         sheet.getRange(row, 1).setValue(new Date().toLocaleString())
//         let setProtection = true
//         items.forEach((item) => {
//             if (item.toLowerCase() === String(value).toLowerCase()) {
//                 setProtection = false
//                 return

//             }
//         })
//         if (setProtection) {
//             let protection = range.protect()
//             protection.removeEditors(protection.getEditors());
//             if (protection.canDomainEdit())
//                 protection.setDomainEdit(false)
//         }
//     }
// }

// function sheetBackup() {
//     Logger.log("backup is waiting")
//     Utilities.sleep(5 * 1000)
//     Logger.log("backup is running")

//     // responsible for backup sheet

//     let sheetName1 = "IMPORTANT"
//     let sheetName2 = "IMPORTANT BACKUP"
//     let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName1.toLowerCase())
//     let backupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName2.toLowerCase())
//     let sheet_range = sheet.getDataRange()
//     let data = sheet_range.getValues()
//     for (var i = 1; i <= sheet_range.getNumRows(); i++) {
//         for (var j = 1; j <= sheet_range.getNumColumns(); j++) {
//             backupSheet.getRange(i, j).setValue(data[i - 1][j - 1])
//         }
//     }

// }




