function sheetProtection() {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("production")
    if (sheet) {
        let row = sheet.getActiveCell().getRow()
        let col = sheet.getActiveCell().getColumn()
        let range = sheet.getRange(row, col)
        let value = range.getValue()
        let items = []

        if (value != "") {
            let setProtection = true
            items.forEach((item) => {
                if (String(item) === String(value)) {
                    setProtection = false
                    return
                }
            })
            if (setProtection) {
                sheet.getRange(row, 1).setValue(new Date().toLocaleString())
                let protection = range.protect()
                protection.removeEditors(protection.getEditors());
                if (protection.canDomainEdit())
                    protection.setDomainEdit(false)
            }
        }
    }

}



