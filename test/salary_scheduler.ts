let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("monthly salary scheduler")
let date = sheet?.getRange(2, 12).getValue()
let refreshDate = sheet?.getRange(2, 4).getValue()

function onOpen() {
    SpreadsheetApp.getUi().createMenu("Salary Scheduler").addItem("Start", 'runSalaryTrigger').addItem("Stop", 'DeleteSalaryTrigger').addToUi();
}

function runSalaryTrigger() {
    let status = false
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (trigger.getHandlerFunction() === "SetUpSalaryTrigger") {
            status = true
        }
    });
    if (status) {
        DisplayAlert("scheduler running, please stop first");
        return
    }
    status = false
    if (date < new Date()) {
        DisplayAlert("scheduler date and time should be greter than now");
        return;
    }
    if (refreshDate < new Date()) {
        DisplayAlert("refesh date and time should be greter than now");
        return;
    }
    ScriptApp.newTrigger('SetUpSalaryTrigger').timeBased().at(date).create()
    ScriptApp.newTrigger('RefreshSalaryStatusTrigger').timeBased().at(refreshDate).create()
    sheet?.getRange(2, 7).setValue("running").setFontColor('green')
}

//refersh salry status
function RefreshSalaryStatusTrigger() {
    sheet?.getRange(2, 3).setValue("not paid")
}

//setup trigger
function SetUpSalaryTrigger() {
    let mf = sheet?.getRange(2, 13).getValue();
    let df = sheet?.getRange(2, 14).getValue();
    let tempARR = [mf, df];
    let count = 0
    tempARR.forEach(function (item) {
        if (item > 0)
            count++;
    });
    if (count > 1) {
        DisplayAlert("minute and day frequency canot work together");
    }
    else if (df > 0) {
        ScriptApp.newTrigger('SendSalaryNotification').timeBased().everyDays(df).create();
    }
    else if (mf > 0) {
        ScriptApp.newTrigger('SendSalaryNotification').timeBased().everyMinutes(mf).create();
    }
    else
        ScriptApp.newTrigger('SendSalaryNotification').timeBased().at(date).create();
}


function SendSalaryNotification() {
    let personName = sheet?.getRange(2, 2).getValue()
    let token = PropertiesService.getScriptProperties().getProperty('accessToken')
    let url = "https://graph.facebook.com/v16.0/103342876089967/messages";
    let data = {
        "messaging_product": "whatsapp",
        "recipient_type": "individual",
        "to": "917056943283",
        "type": "template",
        "template": {
            "name": "salary_reminder",
            "language": {
                "code": "en_US"
            },
            "components": [
                {
                    "type": "header",
                    "parameters": [
                        {
                            "type": "image",
                            "image": {
                                "link": "https://fplogoimages.withfloats.com/tile/605af6c3f7fc820001c55b20.jpg"
                            }
                        }
                    ]
                },
                {
                    "type": "body",
                    "parameters": [
                        {
                            "type": "text",
                            "text": personName
                        }
                    ]
                }
            ]
        }
    }
    let options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        "method": "post",
        "headers": {
            "Authorization": `Bearer ${token}`
        },
        "contentType": "application/json",
        "payload": JSON.stringify(data)
    };
    UrlFetchApp.fetch(url, options)
}

function DeleteSalaryTrigger() {
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (trigger.getHandlerFunction() === "SendSalaryNotification" || trigger.getHandlerFunction() === "SetUpSalaryTrigger" || trigger.getHandlerFunction() === "RefreshSalaryStatusTrigger") {
            ScriptApp.deleteTrigger(trigger);
        }
    });
    DisplayAlert("task stopped");
    sheet?.getRange(2, 7).setValue("stoped").setFontColor('red')
}

function DisplayAlert(message) {
    SpreadsheetApp.getUi().alert(message);
    return;
}