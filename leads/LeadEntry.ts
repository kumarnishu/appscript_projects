function LeadsManagement() {
    let name = "leads_management"
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name.toLowerCase())
    if (sheet) {
        let row = sheet.getActiveCell().getRow()
        let col = sheet.getActiveCell().getColumn()
        let range = sheet.getRange(row, col)
        let value = range.getValue()
        let items = ["OPEN", "USELESS", "WON", "WON DEALER", "LOST", , "POTENTIAL"]

        if (value != "") {
            sheet.getRange(row, 1).setValue(new Date().toLocaleString())
            let setProtection = true
            items.forEach((item) => {
                if (item?.toLowerCase() === String(value).toLowerCase()) {
                    setProtection = false
                    return

                }
            })
            if (setProtection) {
                let protection = range.protect()
                protection.removeEditors(protection.getEditors());
                if (protection.canDomainEdit())
                    protection.setDomainEdit(false)
            }
        }
    }

}
