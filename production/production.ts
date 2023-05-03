function ProductionMenu() {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Are you sure you want to continue?',
        ui.ButtonSet.YES_NO);

    // Process the user's response.
    if (response == ui.Button.YES)
        ProductionEntry()
    else
        return
}

function ProductionEntry() {
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
                firstrow = false
            }
            sheet.getRange(lastrow + i + 2, 1).setValue(date)
            sheet.getRange(lastrow + i + 2, 2).setValue(machines[i])
            sheet.getRange(lastrow + i + 2, 3).setValue(0)
        }
        sheet.setActiveRange(sheet.getRange(lastrow + 2, 3))
    }
}
