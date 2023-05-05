// Compiled using appscript_projects 1.0.0 (TypeScript 4.9.5)
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("whatsapp scheduler");
var date = new Date(sheet === null || sheet === void 0 ? void 0 : sheet.getRange(2, 9).getValue());
var ScriptProperty = PropertiesService.getScriptProperties()
ScriptProperty.setProperty('whatsappcount', '0')
ScriptProperty.setProperty('whatsappcountfailed', '0')
ScriptProperty.setProperty('whatsappcountsuccess', '0')
let whatsappcount = Number(ScriptProperty.getProperty('whatsappcount'))
let whatsappcountfailed = Number(ScriptProperty.getProperty('whatsappcountfailed'))
let whatsappcountsuccess = Number(ScriptProperty.getProperty('whatsappcountsuccess'))

//setup menu and task status for the  trigger
function CreateWhatsappAutomationMenu() {
    SpreadsheetApp.getUi().createMenu("Whatsapp Automation").addItem("Start", 'runWhatsappTrigger').addItem("Stop", 'DeleteWhatsappTrigger').addToUi();
    whatsappCountTasks();
}

function runWhatsappTrigger() {
    if (date < new Date()) {
        SpreadsheetApp.getUi().alert("date and time should be greter than now");
        return;
    }
    ScriptApp.newTrigger('SetUpWhatsappTrigger').timeBased().at(date).create()
    ScriptProperty.setProperty('whatsappcount', String(whatsappcount++))
    sheet?.getRange(2, 11).setValue("task started").setFontColor("green");
}
//setup trigger
function SetUpWhatsappTrigger() {
    var wf = sheet?.getRange(2, 12).getValue();
    var df = sheet?.getRange(2, 13).getValue();
    var hf = sheet?.getRange(2, 14).getValue();
    var mf = sheet?.getRange(2, 15).getValue();
    var weekdayf = sheet?.getRange(2, 16).getValue();
    var monthdayf = sheet?.getRange(2, 17).getValue();
    var tempARR = [wf, df, hf, mf, monthdayf, weekdayf];
    var whatsappcount = 0;
    tempARR.forEach(function (item) {
        if (item > 0)
            whatsappcount++;
    });
    if (whatsappcount > 1) {
        DisplayAlertWhatsapp("please choose one between weeks,hours,min,weekday,monthday--all connot work together");
    }
    if (wf > 0) {
        ScriptApp.newTrigger('SendWhatsapp').timeBased().everyWeeks(wf).create();
    }
    if (df > 0) {
        ScriptApp.newTrigger('SendWhatsapp').timeBased().everyDays(df).create();
    }
    if (hf > 0) {
        ScriptApp.newTrigger('SendWhatsapp').timeBased().everyHours(hf).create();
    }
    if (mf > 0) {
        ScriptApp.newTrigger('SendWhatsapp').timeBased().everyMinutes(mf).create();
    }
    if (weekdayf > 0) {
        ScriptApp.newTrigger('SendWhatsapp').timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(date.getHours()).create();
    }
    if (monthdayf > 0) {
        ScriptApp.newTrigger('SendWhatsapp').timeBased().onMonthDay(monthdayf).atHour(date.getHours()).create();
    }
    ScriptProperty.setProperty('whatsappcount', String(whatsappcount++))
    sheet?.getRange(2, 11).setValue("task started").setFontColor("green");
    whatsappCountTasks()
}
function SendWhatsapp() {
    var url = "https://graph.facebook.com/v16.0/103342876089967/messages";
    var data = {
        "messaging_product": "whatsapp",
        "to": "917056943283",
        "type": "template",
        "template": {
            "name": "hello_world",
            "language": { "code": "en_US" }
        }
    }
    var options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        "method": "post",
        "headers": {
            "Authorization": "Bearer " + "EAANSh4Xqjb4BAI4g2QtGTGHne1YdxbRz2yncJCEQN4jE2TxYu3hzXCZAK9oT6cRYwc5UIHLTtNMmJ567cZBJUc7Yqu9Kb7k7LfE6fXUHM9N523BDYYZAuFeaTZA5yWiSZC6tDGYYv6mnJi8MtHDL7CDjwVQbyOSM8NnT4GuEcge0WEVFCrgMMX0JfX28HEmxoXKhGsZBv9KEgc1h7NEqyWGg1QY3b9tecZD"
        },
        "contentType": "application/json",
        "payload": JSON.stringify(data) 
    };
    var response = UrlFetchApp.fetch(url, options)
    if (response.getResponseCode()===200)
        ScriptProperty.setProperty('whatsappcountsuccess', String(whatsappcountsuccess++))
    else
        ScriptProperty.setProperty('whatsappcountfailed',String(whatsappcountfailed++))
    ScriptProperty.setProperty('whatsappcount', String(whatsappcount--))
}
//delete trigger
function DeleteWhatsappTrigger() {
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (trigger.getHandlerFunction() === "SendWhatsapp" || trigger.getHandlerFunction() === "SetUpWhatsappTrigger") {
            ScriptApp.deleteTrigger(trigger);
            ScriptProperty.setProperty('whatsappcount', String(0))
        }
    });
    DisplayAlertWhatsapp("task stopped");
    whatsappCountTasks();
}
function DisplayAlertWhatsapp(message) {
    SpreadsheetApp.getUi().alert(message);
    return;
}
function whatsappCountTasks() {
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (trigger.getHandlerFunction() === "SendWhatsapp") {
            ScriptProperty.setProperty('whatsappcount', String(whatsappcount++))
        }
    });
    whatsappcountfailed = Number(ScriptProperty.getProperty('whatsappcountfailed'))
    whatsappcountsuccess = Number(ScriptProperty.getProperty('whatsappcountsuccess'))
    sheet?.getRange(2, 10).setValue(`${String(whatsappcount)} task running`);
    sheet?.getRange(2, 11).setValue(`${String(whatsappcountfailed)} failed`).setFontColor("red");
    sheet?.getRange(2, 11).setValue(`${String(whatsappcountsuccess)} success`).setFontColor("green");
}
