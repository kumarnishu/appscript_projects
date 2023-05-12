let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test")

//setup scheduler menu
function onOpen() {
    SpreadsheetApp.getUi().createMenu("Scheduler").addItem("Start", 'StartScheduler').addItem("Stop", 'StopScheduler').addToUi();
}


//start scheduler
function StartScheduler() {
    let status = false;
    ScriptApp.getProjectTriggers().forEach((trigger) => {
        if (trigger.getHandlerFunction() === "CheckSchedulerStatusGlobally") {
            status = true;
        }
    });
    if (status) {
        Alert("scheduler already running, please stop first");
        return;
    }

    if (sheet) {
        //trigger error handler
        for (let i = 3; i <= sheet.getLastRow(); i++) {
            let autoStop = sheet?.getRange(i, 12).getValue()
            let work_status = sheet?.getRange(i, 3).getValue()
            if (String(autoStop).toLowerCase() !== "stop" && String(work_status).toLowerCase() !== "done")
                if (TriggerErrorHandler(i))
                    return
        }

        Alert("Congrats !! validation successful,no Error found, we will start setting up scheduler for each row now ?")

        //setup trigger
        for (let i = 3; i <= sheet.getLastRow(); i++) {
            let autoStop = sheet?.getRange(i, 12).getValue()
            let work_status = sheet?.getRange(i, 3).getValue()
            let date = new Date(sheet?.getRange(i, 13).getValue())
            let phone = sheet?.getRange(i, 7).getValue()
            let name = sheet?.getRange(i, 6).getValue()
            if (String(autoStop).toLowerCase() !== "stop" && String(work_status).toLowerCase() !== "done")
                SetUpTrigger1(date, phone, name, i)
        }
        Alert("Congrats !! We have setup scheduler for each row now Successfully ?")
    }

    ScriptApp.newTrigger('CheckSchedulerStatusGlobally').timeBased().at(new Date()).create()
}

//dummy function to check status globally
function CheckSchedulerStatusGlobally() {
    return
}

//setup repeated trigger for each row
function SetUpRepeatedTrigger() {
    if (sheet) {
        for (let i = 3; i <= sheet?.getLastRow(); i++) {
            let autoStop = sheet?.getRange(i, 12).getValue()
            let work_status = sheet?.getRange(i, 3).getValue()
            let date = new Date(sheet?.getRange(i, 13).getValue())
            let phone = sheet?.getRange(i, 7).getValue()
            let name = sheet?.getRange(i, 6).getValue()
            if (String(autoStop).toLowerCase() !== "stop" && String(work_status).toLowerCase() !== "done")
                SetUpTrigger2(date, phone, name, i)
        }
    }
}

//setup trigger 1
function SetUpTrigger1(date: Date, phoneno: number, name: string, index: number) {
    let trigger = ScriptApp.newTrigger('SetUpRepeatedTrigger').timeBased().at(date).create()
    TrackTrigger(date, phoneno, name, trigger.getUniqueId(), trigger.getHandlerFunction())
    sheet?.getRange(index, 1).setValue("ready").setFontColor("orange").setFontWeight('bold')
}

//setup trigger 2
function SetUpTrigger2(date: Date, phoneno: number, name: string, index: number) {
    let trigger = ScriptApp.newTrigger('SendWhatsappMessageWithButtons').timeBased().at(date).create()
    TrackTrigger(date, phoneno, name, trigger.getUniqueId(),trigger.getHandlerFunction())
    sheet?.getRange(index, 1).setValue("running").setFontColor("green").setFontWeight('bold')
}


//tracking triggers with phone and their id
function TrackTrigger(date: Date, phone: number, name: string, trigger_id: string,trigger_type:string) {
    Logger.log(date + " : " + phone + " : " + trigger_id + " : " + name)
}

//find trigger
function FindTrigger(phoneno: number, trigger_id: string,trigger_type:string) {
    let sheett=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("triggers")
    let trigger: any=undefined
    if (sheett){
        for (let i = 2; i <= sheett.getLastRow(); i++) {
            let trigger_id = sheet?.getRange(i,2).getValue()
            let trigger_type = sheet?.getRange(i, 3).getValue()
            let mobile = sheet?.getRange(i, 4).getValue()
            if (trigger_type === trigger_type && phoneno===mobile){
                trigger=ScriptApp.getProjectTriggers().find(trigger=>trigger.getUniqueId()===trigger_id)
            }
        }
    }
    return trigger
}

//trigger error handler
function TriggerErrorHandler(index) {
    let mf = sheet?.getRange(index, 15).getValue();
    let hf = sheet?.getRange(index, 16).getValue();
    let df = sheet?.getRange(index, 17).getValue();
    let wf = sheet?.getRange(index, 18).getValue();
    let monthf = sheet?.getRange(index, 19).getValue();
    let yearf = sheet?.getRange(index, 20).getValue();
    let weekdays = sheet?.getRange(index, 21).getValue();
    let monthdays = sheet?.getRange(index, 22).getValue()
    let phoneno = sheet?.getRange(index, 7).getValue()

    let TmpArr = [mf, hf, df, wf, monthf, yearf, weekdays, monthdays]

    if (!phoneno) {
        Alert(`Select Phone no first : Error comes from Row No ${index} In Data Range`)
        return true
    }
    let count = 0
    TmpArr.forEach((item) => {
        if (Number(item) > 0) {
            count++;
        }
        else if (String(item) !== "") {
            count++
        }
    });
    if (count > 1) {
        Alert(`Select only one from from hour,minutes,days,weeks,year,weekdays and monthdays repeatation : Error comes from Row No ${index} In Data Range`)
        return true
    }
    return false
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
// function RefreshWorkStatus(frequency) {
//     let last_date = new Date(sheet?.getRange(3, 9).getValue())
//     let count = 0
//     if (frequency === "every minute") {
//         if (count > 0) {
//             Alert("Not allowed  here anyone from from month,weeks,days,hour and minute repeatation")
//         }
//         last_date = new Date(last_date.getTime() + 1 * 60000)
//         sheet?.getRange(3, 9).setValue(last_date)
//     }
//     if (frequency === "every hour") {
//         if (count > 0) {
//             Alert("Not allowed  here anyone from from month,weeks,days,hour and minute repeatation")
//         }
//         last_date = new Date(last_date.getTime() + 60 * 60000)
//         sheet?.getRange(3, 9).setValue(last_date)
//     }
//     if (frequency === "every day") {
//         if (count > 0) {
//             Alert("Not allowed  here anyone from from month,weeks,days,hour and minute repeatation")
//         }
//         last_date.setDate(last_date.getDate() + 1)
//         sheet?.getRange(3, 9).setValue(last_date)
//     }
//     if (frequency === "every week") {
//         last_date.setDate(last_date.getDate() + 7)
//         sheet?.getRange(3, 9).setValue(last_date)
//     }
//     if (frequency === "every month") {
//         last_date.setDate(last_date.getDate() + GetMonthDays(last_date.getFullYear(), last_date.getMonth()))
//         sheet?.getRange(3, 9).setValue(last_date)
//     }
//     if (frequency === "every year") {
//         last_date.setDate(last_date.getDate() + 365)
//         sheet?.getRange(3, 9).setValue(last_date)
//     }
//     sheet?.getRange(3, 7).setValue("pending").setFontColor('red')
//     sheet?.getRange(3, 13).setValue(true)
// }

// stop scheduler
function StopScheduler() {
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (trigger.getHandlerFunction() === "SendWhatsappMessageWithButtons" || trigger.getHandlerFunction() === "setUpDateTrigger") {
            ScriptApp.deleteTrigger(trigger);
        }
    });
    Alert("task stopped");
}