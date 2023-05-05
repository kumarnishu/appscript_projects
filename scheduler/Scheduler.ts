// Compiled using appscript_projects 1.0.0 (TypeScript 4.9.5)
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("scheduler");
var date = new Date(sheet === null || sheet === void 0 ? void 0 : sheet.getRange(2, 9).getValue());
var ScriptProperty = PropertiesService.getScriptProperties()
ScriptProperty.setProperty('count', '0')
let count = Number(ScriptProperty.getProperty('count'))

//setup menu and task status for the  trigger
function CreateEmailAutomationMenu() {
    SpreadsheetApp.getUi().createMenu("Email Automation").addItem("Start", 'runEmailTrigger').addItem("Stop", 'DeleteTrigger').addToUi();
    CountTasks();
}

function runEmailTrigger() {
    if (date < new Date()) {
        SpreadsheetApp.getUi().alert("date and time should be greter than now");
        return;
    }
    ScriptApp.newTrigger('SetUpEmailTrigger').timeBased().at(date).create()
    ScriptProperty.setProperty('count', String(count++))
    CountTasks();
}
//setup trigger
function SetUpEmailTrigger() {
    var wf = sheet?.getRange(2, 11).getValue();
    var df = sheet?.getRange(2, 12).getValue();
    var hf = sheet?.getRange(2, 13).getValue();
    var mf = sheet?.getRange(2, 14).getValue();
    var weekdayf = sheet?.getRange(2, 15).getValue();
    var monthdayf = sheet?.getRange(2, 16).getValue();
    var tempARR = [wf, df, hf, mf, monthdayf, weekdayf];
    var count = 0;
    tempARR.forEach(function (item) {
        if (item > 0)
            count++;
    });
    if (count > 1) {
        DisplayAlert("please choose one between weeks,hours,min,weekday,monthday--all connot work together");
    }
    if (wf > 0) {
        ScriptApp.newTrigger('SendEmail').timeBased().everyWeeks(wf).create();
    }
    if (df > 0) {
        ScriptApp.newTrigger('SendEmail').timeBased().everyDays(df).create();
    }
    if (hf > 0) {
        ScriptApp.newTrigger('SendEmail').timeBased().everyHours(hf).create();
    }
    if (mf > 0) {
        ScriptApp.newTrigger('SendEmail').timeBased().everyMinutes(mf).create();
    }
    if (weekdayf > 0) {
        ScriptApp.newTrigger('SendEmail').timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(date.getHours()).create();
    }
    if (monthdayf > 0) {
        ScriptApp.newTrigger('SendEmail').timeBased().onMonthDay(monthdayf).atHour(date.getHours()).create();
    }
    ScriptProperty.setProperty('count', String(count++))
    CountTasks();
}
function SendEmail() {
    GmailApp.sendEmail("kumarnishu437@gmail.com", "testing mail from scheduler", "this is automated message");
    ScriptProperty.setProperty('count', String(count--))
    CountTasks();
}
//delete trigger
function DeleteTrigger() {
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (trigger.getHandlerFunction() === "SendEmail" || trigger.getHandlerFunction() === "SetUpEmailTrigger") {
            ScriptApp.deleteTrigger(trigger);
            ScriptProperty.setProperty('count', String(count--))
        }
    });
    DisplayAlert("task stopped");
    CountTasks();
}
function DisplayAlert(message) {
    SpreadsheetApp.getUi().alert(message);
    return;
}
function CountTasks() {
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (trigger.getHandlerFunction() === "SendEmail") {
            ScriptProperty.setProperty('count', String(count++))
        }
    });
    sheet?.getRange(2, 10).setValue(`${String(count)} task running`);
}
