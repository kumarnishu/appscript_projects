function PreventRenamingSheets() {
    var ss = SpreadsheetApp.getActive();
    var sheets = ss.getSheets();
    var editors :string[]= [];
    ss.getEditors().forEach(function (user) {
        editors.push(user.getEmail());

    });

    for (var i = 0; i < sheets.length; i++) {
        var ws = sheets[i];
        if (ws.getProtections(SpreadsheetApp.ProtectionType.SHEET).length === 0) {

            var cells = ws.getRange(1, 1, ws.getMaxRows(), ws.getMaxColumns());

            // Protect the sheet with all cells unprotected, in this way the sheet name can not be edited by other editors

            ws.protect().setUnprotectedRanges([cells]).removeEditors(editors);

        }

    }

}