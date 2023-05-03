type User = {
    name: string,
    email: string,
    resources: {
        sheet: string,
        access_type: string,
        link: string
    }[]
}

function EmailWise() {
    let data = DriveApp.getFiles()
    let sheetMimeType = "application/vnd.google-apps.spreadsheet"
    let Files: GoogleAppsScript.Drive.File[] = []
    let users: {
        user: GoogleAppsScript.Drive.User,
        email: string
    }[] = []
    let sheets: GoogleAppsScript.Drive.File[] = []

    // convert iterator to array
    while (data.hasNext()) {
        Files.push(data.next())
    }

    //filter google sheets
    Files.map((file) => {
        if (file.getMimeType() === sheetMimeType) {
            sheets.push(file)
            users.push({ user: file.getOwner(), email: file.getOwner().getEmail() })
            file.getEditors().forEach((editor) => users.push({ user: editor, email: editor.getEmail() }))
            file.getViewers().forEach((viewer) => users.push({ user: viewer, email: viewer.getEmail() }))
        }
    })

    //filter users
    let newusers: {
        user: GoogleAppsScript.Drive.User,
        email: string
    }[] = []

    let tempEmails: string[] = []
    users.forEach((user) => {
        if (!tempEmails.includes(user.email)) {
            newusers.push(user)
            tempEmails.push(user.email)
        }
    })

    let finalUsers: User[] = []
    let resources: User['resources'] = []
    for (let i = 0; i < newusers.length; i++) {
        sheets.map((file) => {
            let access = "viewer"
            if (file.getOwner().getEmail() === newusers[i].email)
                access = "owner"
            file.getEditors().forEach((editor) => {
                if (editor.getEmail() === newusers[i].email)
                    access = "editor"
            })
            resources.push({
                sheet: file.getName(),
                access_type: access,
                link: file.getUrl()
            })
        })
        finalUsers.push({
            name: newusers[i].user.getName(),
            email: newusers[i].email,
            resources: resources
        })
        resources = []
    }


    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Wise")
    if (sheet) {
        sheet.clear()
        let rowindex = 1
        finalUsers.forEach((user) => {
            let firstrow = true
            for (let i = 0; i < user.resources.length; i++) {
                if (firstrow) {
                    sheet?.getRange(rowindex, 1).setValue(user.name)
                    sheet?.getRange(rowindex, 2).setValue(user.email)
                    sheet?.getRange(rowindex, 3).setValue(user.resources[i].sheet)
                    sheet?.getRange(rowindex, 4).setValue(user.resources[i].access_type)
                    sheet?.getRange(rowindex, 5).setValue(user.resources[i].link)
                    firstrow = false
                }
                else {
                    sheet?.getRange(rowindex, 3).setValue(user.resources[i].sheet)
                    sheet?.getRange(rowindex, 4).setValue(user.resources[i].access_type)
                    sheet?.getRange(rowindex, 5).setValue(user.resources[i].link)
                }
                rowindex++
            }
        })
    }

}
