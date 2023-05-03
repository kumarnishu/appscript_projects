
//protecting cell for editors after creating table
function ProtectCell(range: GoogleAppsScript.Spreadsheet.Range) {
    let protection = range.protect()
    let user = Session.getActiveUser()
    protection.addEditor(user)
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit())
        protection.setDomainEdit(false)
}

//check owner
function CheckOwner() {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('production')
    if (sheet) {
        let owner_email = SpreadsheetApp?.getActive()?.getOwner()?.getEmail()
        if (owner_email !== Session.getActiveUser().getEmail()) {
            var ui = SpreadsheetApp.getUi();
            ui.alert("must be sheet owner to execute this task")
            return false
        }
        else {
            return true
        }
    }
    return false
}

//launch one month table menu
function CreateOneMonthProductionMenu() {
    if (!CheckOwner()) return
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('You Want to Create One Month Production Table ?',
        ui.ButtonSet.YES_NO);

    // Process the user's response.
    if (response == ui.Button.YES)
        CreateOneMonthProduction(30)
    else
        return
}

//create one month table
function CreateOneMonthProduction(days: number) {
    if (!CheckOwner()) return
    let fisrtFocus = true
    for (let i = 0; i < days; i++) {
        if (fisrtFocus) {
            OneDayProduction(true)
            fisrtFocus = false
        }
        else
            OneDayProduction(false)
    }
}

//create one day table
function OneDayProduction(focus: boolean) {
    if (!CheckOwner()) return
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('production')
    if (sheet) {
        let lastrow = sheet.getDataRange().getLastRow()
        let date = new Date(sheet.getRange(lastrow, 1).getValue())
        // add a day
        date.setDate(date.getDate() + 1);
        let machines = [
            "VER-1  (GF)",
            "VER-2  (GF)",
            "VER-3  (GF)",
            "VER-4  (GF)",
            "VER-5  (GF)",
            "VER-6  (GF)",
            "LYM-7",
            "LYM-8",
            "LYM-9",
            "VER-10  (GF)",
            "VER-11  (GF)",
            "VER-12  (SF)",
            "VER-13  (SF)",
            "VER-14  (SF)",
            "VER-15  (SF)",
            "VER-16  (SF)",
            "VER-17  (SF)",
            "GBOOT-18",
            "GBOOT-19",
            "PU MACHINE - 20"
        ]
        let firstrow = true
        for (let i = 0; i < machines.length; i++) {
            if (firstrow) {
                sheet.getRange(lastrow + i + 1, 1).setValue("Date").setFontWeight('bold')
                sheet.getRange(lastrow + i + 1, 2).setValue("Machines").setFontWeight('bold')
                sheet.getRange(lastrow + i + 1, 3).setValue("Productions").setFontWeight('bold')
                ProtectCell(sheet.getRange(lastrow + i + 1, 3))
                firstrow = false

            }
            sheet.getRange(lastrow + i + 2, 1).setValue(date.toLocaleDateString())
            sheet.getRange(lastrow + i + 2, 2).setValue(machines[i])
            sheet.getRange(lastrow + i + 2, 3).setValue(0)
        }
        ProtectCell(sheet.getRange(`A${lastrow + 1}:B${lastrow + 1 + machines.length}`))
        if (focus) {
            sheet.setActiveRange(sheet.getRange(lastrow + 2, 3))
            focus = false
        }
    }
}

//protect after edit
function onEditProductionProtect() {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("production")
    if (sheet) {
        let row = sheet.getActiveCell().getRow()
        let col = sheet.getActiveCell().getColumn()
        let range = sheet.getRange(row, col)
        let value = range.getValue()
        let items = []

        if (value !=="") {
            let setProtection = true
            items.forEach((item) => {
                if (String(item) === String(value)) {
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