let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test")
let schedule_time = sheet?.getRange(3, 8).getValue()
let date = new Date(sheet?.getRange(3, 14).getValue())
let work_title = sheet?.getRange(3, 3).getValue()
let work_detail = sheet?.getRange(3, 4).getValue()
let personName = sheet?.getRange(3, 5).getValue()
let phoneno = sheet?.getRange(3, 6).getValue()
//refresh date and autoStop value  picked from sheet always fresh
let refreshDate = sheet?.getRange(3, 15).getValue()
let autoStop = sheet?.getRange(3, 13).getValue()
let work_status = sheet?.getRange(3, 7).getValue()
let tsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('triggers')
let triggerDate = tsheet?.getRange(2, 1)
let triggerId = tsheet?.getRange(2, 2).getValue()


//setup scheduler menu
function onOpen() {
    SpreadsheetApp.getUi().createMenu("Scheduler").addItem("Start", 'StartScheduler').addItem("Stop", 'StopScheduler').addToUi();
}

//start scheduler
function StartScheduler() {
    let status = false
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (trigger.getHandlerFunction() === "setUpTrigger") {
            status = true
        }
    });
    if (status) {
        Alert("scheduler already running, please stop first");
        return
    }
    status = false
    if (date < new Date()) {
        Alert("scheduler date and time should be greater than now");
        return;
    }
    let trigger = ScriptApp.newTrigger('setUpTrigger').timeBased().at(date).create()
    tsheet?.getRange(2, 1).setValue(date)
    tsheet?.getRange(2, 2).setValue(trigger.getUniqueId())
    tsheet?.getRange(2, 3).setValue(phoneno)
    sheet?.getRange(3, 1).setValue("running").setFontColor('green')
}


// setup trigger based on frequency and repeatation
function setUpTrigger() {
    let frequency = sheet?.getRange(3, 12).getValue()
    let date = sheet?.getRange(3, 14).getValue()
    let autoStop = sheet?.getRange(3, 13).getValue()
    let weekdays=[]
    if (!autoStop && work_status == "pending") {
        refreshDate = new Date(date)
        var mf = sheet?.getRange(3, 16).getValue();
        var hf = sheet?.getRange(3, 17).getValue();
        var df = sheet?.getRange(3, 18).getValue();
        var wf = sheet?.getRange(3, 19).getValue();
        var monthf = sheet?.getRange(3, 20).getValue();
        let TmpArr = [mf, hf, df, wf, monthf]
        let count = 0
        TmpArr.forEach((item) => {
            if (item > 0)
                count++;
        });
        Logger.log(count)
        if (count > 1) {
            Alert("Not allowed  here anyone from from month,weeks,days,hour and minute repeatation")
        }
        if (frequency === "every minute") {
            if (count > 0) {
                Alert("Not allowed  here anyone from from month,weeks,days,hour and minute repeatation")
            }
            refreshDate = new Date(refreshDate.getTime() + 1 * 60000)
            sheet?.getRange(3, 15).setValue(refreshDate)
        }
        if (frequency === "every hour") {
            if (count > 0) {
                Alert("Not allowed  here anyone from from month,weeks,days,hour and minute repeatation")
            }
            refreshDate = new Date(refreshDate.getTime() + 60 * 60000)
            sheet?.getRange(3, 15).setValue(refreshDate)
        }
        if (frequency === "every day") {
            if (count > 0) {
                Alert("Not allowed  here anyone from from month,weeks,days,hour and minute repeatation")
            }
            refreshDate.setDate(refreshDate.getDate() + 1)
            sheet?.getRange(3, 15).setValue(refreshDate)
        }
        if (frequency === "every week") {
            let repeatationf = null
            let repeatationType = ""
            if (mf > 0) {
                repeatationf = mf
                repeatationType = "mf"
            }
            if (hf > 0) {
                repeatationf = hf
                repeatationType = "hf"
            }
            if (df > 0) {
                repeatationf = df
                repeatationType = "df"
            }
            if (repeatationf && repeatationType === "")
                SetUpRepeatedTrigger(frequency, repeatationType, repeatationf)
            refreshDate.setDate(refreshDate.getDate() + 7)
            sheet?.getRange(3, 15).setValue(refreshDate)
        }
        if (frequency === "every month") {
            let repeatationf = null
            let repeatationType = ""
            if (mf > 0) {
                repeatationf = mf
                repeatationType = "mf"
            }
            if (hf > 0) {
                repeatationf = hf
                repeatationType = "hf"
            }
            if (df > 0) {
                repeatationf = df
                repeatationType = "df"
            }
            if (wf > 0) {
                repeatationf = wf
                repeatationType = "wf"
            }
            if (repeatationf && repeatationType === "")
                SetUpRepeatedTrigger(frequency, repeatationType, repeatationf)
            refreshDate.setDate(refreshDate.getDate() + GetMonthDays(refreshDate.getFullYear(), refreshDate.getMonth()))
            sheet?.getRange(3, 15).setValue(refreshDate)
        }
        if (frequency === "every year") {
            let repeatationf = null
            let repeatationType = ""
            if (mf > 0) {
                repeatationf = mf
                repeatationType = "mf"
            }
            if (hf > 0) {
                repeatationf = hf
                repeatationType = "hf"
            }
            if (df > 0) {
                repeatationf = df
                repeatationType = "df"
            }
            if (wf > 0) {
                repeatationf = wf
                repeatationType = "wf"
            }
            if (monthf > 0) {
                repeatationf = monthf
                repeatationType = "monthf"
            }
            if (repeatationf && repeatationType === "")
                SetUpRepeatedTrigger(frequency, repeatationType, repeatationf)
            refreshDate.setDate(refreshDate.getDate() + 365)
            sheet?.getRange(3, 15).setValue(refreshDate)
        }
    }
    return

}


//send whatsapp message with response buttons
function SendWhatsappMessageWithButtons() {
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

//showing alert message in sheet
function Alert(message: string) {
    SpreadsheetApp.getUi().alert(message);
    return;
}

//get month days
function GetMonthDays(year: number, month: number) {
    Logger.log(month + " :" + year)
    let febDays = 28
    if (year % 4 === 0) {
        febDays = 29
    }
    let day31 = [1, 3, 5, 7, 8, 10, 12]
    let day30 = [4, 6, 9, 11]
    if (day31.includes(month))
        return 31
    if (day30.includes(month))
        return 30
    return febDays
}

//refresh work status
function RefreshWorkStatus(frequency) {
    let last_date = new Date(sheet?.getRange(3, 9).getValue())
    let count = 0
    if (frequency === "every minute") {
        if (count > 0) {
            Alert("Not allowed  here anyone from from month,weeks,days,hour and minute repeatation")
        }
        last_date = new Date(last_date.getTime() + 1 * 60000)
        sheet?.getRange(3, 9).setValue(last_date)
    }
    if (frequency === "every hour") {
        if (count > 0) {
            Alert("Not allowed  here anyone from from month,weeks,days,hour and minute repeatation")
        }
        last_date = new Date(last_date.getTime() + 60 * 60000)
        sheet?.getRange(3, 9).setValue(last_date)
    }
    if (frequency === "every day") {
        if (count > 0) {
            Alert("Not allowed  here anyone from from month,weeks,days,hour and minute repeatation")
        }
        last_date.setDate(last_date.getDate() + 1)
        sheet?.getRange(3, 9).setValue(last_date)
    }
    if (frequency === "every week") {
        last_date.setDate(last_date.getDate() + 7)
        sheet?.getRange(3, 9).setValue(last_date)
    }
    if (frequency === "every month") {
        last_date.setDate(last_date.getDate() + GetMonthDays(last_date.getFullYear(), last_date.getMonth()))
        sheet?.getRange(3, 9).setValue(last_date)
    }
    if (frequency === "every year") {
        last_date.setDate(last_date.getDate() + 365)
        sheet?.getRange(3, 9).setValue(last_date)
    }
    sheet?.getRange(3, 7).setValue("pending").setFontColor('red')
    sheet?.getRange(3, 13).setValue(true)
}

//stop scheduler
function StopScheduler() {
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (trigger.getHandlerFunction() === "SendWhatsappMessageWithButtons" || trigger.getHandlerFunction() === "setUpTrigger") {
            ScriptApp.deleteTrigger(trigger);
        }
    });
    Alert("task stopped");
}