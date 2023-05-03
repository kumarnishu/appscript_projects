type Sheet = {
    sheet: GoogleAppsScript.Drive.File,
    owner: GoogleAppsScript.Drive.User,
    last_edit: string,
    link: string,
    editors: GoogleAppsScript.Drive.User[],
    viewers: GoogleAppsScript.Drive.User[]
}
function SheetWise() {
    let data = DriveApp.getFiles()
    let sheetMimeType = "application/vnd.google-apps.spreadsheet"
    let Files: GoogleAppsScript.Drive.File[] = []
    let sheets: Sheet[] = []

    // convert iterator to array
    while (data.hasNext()) {
        Files.push(data.next())
    }

    //filter google sheets
    Files.map((file) => {
        if (file.getMimeType() === sheetMimeType) {
            sheets.push({
                sheet: file,
                owner: file.getOwner(),
                last_edit: file.getLastUpdated().toLocaleString(),
                link: file.getUrl(),
                editors: file.getEditors(),
                viewers: file.getViewers()
            })
        }
    })

    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet Wise")
    if (sheet) {
        sheet.clear()
        for (let i = 0; i < sheets.length; i++) {
            sheet.getRange(i * 2 + 1, 1).setValue(sheets[i].sheet.getName())
            sheet.getRange(i * 2 + 1, 2).setValue(sheets[i].owner.getName() + " : " + sheets[i].owner.getEmail())
            sheet.getRange(i * 2 + 1, 3).setValue(sheets[i].link)
            sheet.getRange(i * 2 + 1, 4).setValue("editors")
            sheet.getRange(i * 2 + 2, 4).setValue("viewers")
            for (let j = 0; j < sheets[i].editors.length; j++) {
                sheet.getRange(i * 2 + 1, j + 5).setValue(sheets[i].editors[j].getName() + " : " + sheets[i].editors[j].getEmail())
            }
            for (let k = 0; k < sheets[i].viewers.length; k++) {
                sheet.getRange(i * 2 + 2, k + 5).setValue(sheets[i].viewers[k].getName() + " : " + sheets[i].viewers[k].getEmail())
            }
        }
    }

}
