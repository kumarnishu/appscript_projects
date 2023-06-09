let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test")

type Trigger = {
    date: Date, trigger_id: string, trigger_type: string, phone: string, name: string, work_title: string, work_detail: string,
    mf: number, hf: number, df: number, wf: number, monthf: number, yearf: number, weekdays: string, monthdays: string
}

//setup scheduler menu
function onOpen() {
    SpreadsheetApp.getUi().createMenu("Scheduler").addItem("Start", 'StartScheduler').addItem("Stop", 'DeleteAllTriggers').addToUi();
}

//start scheduler
function StartScheduler() {
    if (sheet) {
        //trigger error handler
        for (let i = 3; i <= sheet.getLastRow(); i++) {
            let scheduler_status = String(sheet?.getRange(i, 1).getValue())
            let autoStop = String(sheet?.getRange(i, 12).getValue())
            let work_status = String(sheet?.getRange(i, 3).getValue())
            if (autoStop.toLowerCase() !== "stop" && work_status.toLowerCase() !== "done" && scheduler_status.toLowerCase() !== "running" && scheduler_status.toLowerCase() !== "ready") {
                if (TriggerErrorHandler(i))
                    return
            }
        }
        DuplicatePhoneChecker()
        //setup trigger
        for (let i = 3; i <= sheet.getLastRow(); i++) {
            let scheduler_status = String(sheet?.getRange(i, 1).getValue())
            let autoStop = String(sheet?.getRange(i, 12).getValue())
            let work_status = String(sheet?.getRange(i, 3).getValue())

            if (autoStop.toLowerCase() !== "stop" && work_status.toLowerCase() !== "done" && scheduler_status.toLowerCase() !== "running" && scheduler_status.toLowerCase() !== "ready") {
                DateTrigger(i)
                RefreshTrigger(i)
            }
        }
    }
}


//setup whatsapp trigger for each row
function SetUpWhatsappTrigger() {
    if (sheet) {
        for (let i = 3; i <= sheet?.getLastRow(); i++) {
            let autoStop = sheet?.getRange(i, 12).getValue()
            let work_status = sheet?.getRange(i, 3).getValue()
            if (String(autoStop).toLowerCase() !== "stop" && String(work_status).toLowerCase() !== "done") {
                WhatsappTrigger(i)
            }
        }
    }
}

// date trigger
function DateTrigger(index: number) {
    let date = new Date(sheet?.getRange(index, 13).getValue())
    let phone = sheet?.getRange(index, 7).getValue()
    let name = sheet?.getRange(index, 6).getValue()
    let work_title = sheet?.getRange(index, 4).getValue()
    let work_detail = sheet?.getRange(index, 5).getValue()
    let mf = Number(sheet?.getRange(index, 15).getValue())
    let hf = Number(sheet?.getRange(index, 16).getValue())
    let df = Number(sheet?.getRange(index, 17).getValue())
    let wf = Number(sheet?.getRange(index, 18).getValue())
    let monthf = Number(sheet?.getRange(index, 19).getValue())
    let yearf = Number(sheet?.getRange(index, 20).getValue())
    let weekdays = String(sheet?.getRange(index, 21).getValue())
    let monthdays = String(sheet?.getRange(index, 22).getValue())
    let trigger = ScriptApp.newTrigger('SetUpWhatsappTrigger').timeBased().at(date).create()

    SaveTrigger({
        date: date, trigger_id: trigger.getUniqueId(), trigger_type: trigger.getHandlerFunction(), phone: phone, name: name, work_title: work_title, work_detail: work_detail, mf: mf, hf: hf, df: df, wf: wf, monthf: monthf, yearf: yearf, weekdays: weekdays,
        monthdays: monthdays
    })
    sheet?.getRange(index, 1).setValue("ready").setFontWeight('bold')
}

// whatsapp trigger 
function WhatsappTrigger(index: number) {
    let date = new Date(sheet?.getRange(index, 13).getValue())
    let phone = sheet?.getRange(index, 7).getValue()
    let name = sheet?.getRange(index, 6).getValue()
    let work_title = sheet?.getRange(index, 4).getValue()
    let work_detail = sheet?.getRange(index, 5).getValue()
    let mf = sheet?.getRange(index, 15).getValue();
    let hf = sheet?.getRange(index, 16).getValue();
    let df = sheet?.getRange(index, 17).getValue();
    let wf = sheet?.getRange(index, 18).getValue();
    let monthf = sheet?.getRange(index, 19).getValue();
    let yearf = sheet?.getRange(index, 20).getValue();
    if (!mf || typeof (mf) !== "number") mf = 0
    if (!hf || typeof (hf) !== "number") hf = 0
    if (!df || typeof (df) !== "number") df = 0
    if (!wf || typeof (wf) !== "number") wf = 0
    if (!monthf || typeof (monthf) !== "number") monthf = 0
    if (!yearf || typeof (yearf) !== "number") yearf = 0
    let weekdays = String(sheet?.getRange(index, 21).getValue())
    let monthdays = String(sheet?.getRange(index, 22).getValue())
    let triggers: GoogleAppsScript.Script.Trigger[] = []
    if (mf > 0) {
        let tr = ScriptApp.newTrigger('SendWhatsappMessageWithButtons').timeBased().everyMinutes(mf).create();
        triggers.push(tr)
    }
    if (hf > 0) {
        let tr = ScriptApp.newTrigger('SendWhatsappMessageWithButtons').timeBased().everyHours(hf).create();
        triggers.push(tr)
    }
    if (df > 0) {
        let tr = ScriptApp.newTrigger('SendWhatsappMessageWithButtons').timeBased().everyDays(df).create();
        triggers.push(tr)
    }
    if (wf > 0) {
        let tr = ScriptApp.newTrigger('SendWhatsappMessageWithButtons').timeBased().everyWeeks(wf).create();
        triggers.push(tr)
    }
    if (monthf > 0) {
        let tr = ScriptApp.newTrigger('SendWhatsappMessageWithButtons').timeBased().everyDays(GetMonthDays(date.getFullYear(), date.getMonth())).create()
        triggers.push(tr)
    }
    if (yearf > 0) {
        let tr = ScriptApp.newTrigger('SendWhatsappMessageWithButtons').timeBased().everyDays(365 * yearf).create();
        triggers.push(tr)
    }
    if (weekdays.length > 0) {
        weekdays.split(",").forEach((wd) => {
            if (wd.toLowerCase() === "sun") {
                let tr = ScriptApp.newTrigger('SendWhatsappMessageWithButtons').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(date.getHours()).create();
                triggers.push(tr)
            }
            if (wd.toLowerCase() === "mon") {
                let tr = ScriptApp.newTrigger('SendWhatsappMessageWithButtons').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(date.getHours()).create();
                triggers.push(tr)
            }
            if (wd.toLowerCase() === "tue") {
                let tr = ScriptApp.newTrigger('SendWhatsappMessageWithButtons').timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(date.getHours()).create();
                triggers.push(tr)
            }
            if (wd.toLowerCase() === "wed") {
                let tr = ScriptApp.newTrigger('SendWhatsappMessageWithButtons').timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(date.getHours()).create();
                triggers.push(tr)
            }
            if (wd.toLowerCase() === "thu") {
                let tr = ScriptApp.newTrigger('SendWhatsappMessageWithButtons').timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(date.getHours()).create();
                triggers.push(tr)
            }
            if (wd.toLowerCase() === "fri") {
                let tr = ScriptApp.newTrigger('SendWhatsappMessageWithButtons').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(date.getHours()).create();
                triggers.push(tr)
            }
            if (wd.toLowerCase() === "sat") {
                let tr = ScriptApp.newTrigger('SendWhatsappMessageWithButtons').timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(date.getHours()).create();
                triggers.push(tr)
            }
        })
    }
    if (monthdays.length > 0) {
        monthdays.split(",").forEach((md) => {
            let tr = ScriptApp.newTrigger('SendWhatsappMessageWithButtons').timeBased().onMonthDay(Number(md)).atHour(date.getHours()).create();
            triggers.push(tr)
        })
    }
    let tr = ScriptApp.newTrigger('SendWhatsappMessageWithButtons').timeBased().at(date).create();
    triggers.push(tr)
    triggers.forEach((trigger) => {
        SaveTrigger({
            date: date, trigger_id: trigger.getUniqueId(), trigger_type: trigger.getHandlerFunction(), phone: phone, name: name, work_title: work_title, work_detail: work_detail,
            mf: mf, hf: hf, df: df, wf: wf, monthf: monthf, yearf: yearf, weekdays: weekdays,
            monthdays: monthdays
        })
    })
    sheet?.getRange(index, 1).setValue("running").setFontWeight('bold')
}

//Saving to track triggers with phone and their id
function SaveTrigger(trigger: Trigger) {
    let tsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('triggers')
    let row = tsheet?.getDataRange().getLastRow()
    if (row)
        row = row + 1
    if (row) {
        tsheet?.getRange(row, 1).setValue(trigger.date)
        tsheet?.getRange(row, 2).setValue(trigger.trigger_id)
        tsheet?.getRange(row, 3).setValue(trigger.trigger_type)
        tsheet?.getRange(row, 4).setValue(trigger.phone)
        tsheet?.getRange(row, 5).setValue(trigger.name)
        tsheet?.getRange(row, 6).setValue(trigger.work_title)
        tsheet?.getRange(row, 7).setValue(trigger.work_detail)
        tsheet?.getRange(row, 8).setValue(trigger.mf)
        tsheet?.getRange(row, 9).setValue(trigger.hf)
        tsheet?.getRange(row, 10).setValue(trigger.df)
        tsheet?.getRange(row, 11).setValue(trigger.wf)
        tsheet?.getRange(row, 12).setValue(trigger.monthf)
        tsheet?.getRange(row, 13).setValue(trigger.yearf)
        tsheet?.getRange(row, 14).setValue(trigger.weekdays)
        tsheet?.getRange(row, 15).setValue(trigger.monthdays)
    }
}

//find trigger
function FindAllTriggers(phoneno: number, trigger_type: string) {
    let sheett = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("triggers")
    let triggers: GoogleAppsScript.Script.Trigger[] = []
    if (sheett) {
        for (let i = 2; i <= sheett.getLastRow(); i++) {
            let trigger_id = sheett?.getRange(i, 2).getValue()
            let trigger_t = sheett?.getRange(i, 3).getValue()
            let mobile = sheett?.getRange(i, 4).getValue()
            if (trigger_type === trigger_t && phoneno === mobile) {
                let tr = ScriptApp.getProjectTriggers().find(trigger => trigger.getUniqueId() === trigger_id)
                if (tr)
                    triggers.push(tr)
            }
        }
    }
    return triggers
}

//all triggers
function findAllTriggersFromTriggersSheet() {
    let tsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("triggers")
    let triggers: Trigger[] = []
    if (tsheet) {
        for (let i = 2; i <= tsheet.getLastRow(); i++) {
            if (tsheet) {
                let date = tsheet.getRange(i, 1).getValue()
                let trigger_id = tsheet.getRange(i, 2).getValue()
                let trigger_type = tsheet.getRange(i, 3).getValue()
                let phone = tsheet.getRange(i, 4).getValue()
                let name = tsheet.getRange(i, 5).getValue()
                let work_title = tsheet.getRange(i, 6).getValue()
                let work_detail = tsheet.getRange(i, 7).getValue()
                let mf = tsheet.getRange(i, 8).getValue()
                let hf = tsheet.getRange(i, 9).getValue()
                let df = tsheet.getRange(i, 10).getValue()
                let wf = tsheet.getRange(i, 11).getValue()
                let monthf = tsheet.getRange(i, 12).getValue()
                let yearf = tsheet.getRange(i, 13).getValue()
                let weekdays = tsheet.getRange(i, 14).getValue()
                let monthdays = tsheet.getRange(i, 15).getValue()
                triggers.push({
                    date: date, trigger_id: trigger_id, trigger_type: trigger_type, phone: phone, name: name, work_title: work_title, work_detail: work_detail,
                    mf: mf, hf: hf, df: df, wf: wf, monthf: monthf, yearf: yearf, weekdays: weekdays,
                    monthdays: monthdays
                })
            }
        }
    }
    return triggers
}

//delete trigger
function DeleteTrigger(trigger: GoogleAppsScript.Script.Trigger, index: number) {
    ScriptApp.deleteTrigger(trigger)
    sheet?.getRange(index, 1).setValue("stopped").setFontWeight('bold')
}

//trigger error handler
function TriggerErrorHandler(index) {
    let errorStatus = false
    let mf = sheet?.getRange(index, 15).getValue();
    let hf = sheet?.getRange(index, 16).getValue();
    let df = sheet?.getRange(index, 17).getValue();
    let wf = sheet?.getRange(index, 18).getValue();
    let monthf = sheet?.getRange(index, 19).getValue();
    let yearf = sheet?.getRange(index, 20).getValue();
    let mins = [0, 1, 5, 10, 15, 30]
    if (!mf) mf = 0
    if (typeof (mf) !== "number" || !mins.includes(mf)) {
        Alert(`Select valid  minutes ,choose one from 0,1,5,10,15,30: Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    if (!hf || typeof (hf) !== "number") hf = 0
    if (!df || typeof (df) !== "number") df = 0
    if (!wf || typeof (wf) !== "number") wf = 0
    if (!monthf || typeof (monthf) !== "number") monthf = 0
    if (!yearf || typeof (yearf) !== "number") yearf = 0
    let weekdays = String(sheet?.getRange(index, 21).getValue())
    let monthdays = String(sheet?.getRange(index, 22).getValue())
    let phoneno = sheet?.getRange(index, 7).getValue()
    let date = new Date(sheet?.getRange(index, 13).getValue())
    if (date < new Date()) {
        Alert(`Select valid  date ,date could not be in the past: Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    let TmpArr = [mf, hf, df, wf, monthf, yearf]
    if (!phoneno) {
        Alert(`Select Phone no first : Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    if (String(phoneno).length < 12) {
        Alert(`Select Phone no in correct format : Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    let count = 0
    TmpArr.forEach((item) => {
        if (item > 0) {
            count++;
        }
    });
    let tmpWeekdays = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
    if (weekdays.length > 0) {
        let weekDays = weekdays.split(",")
        weekDays.forEach((item) => {
            if (!tmpWeekdays.includes(item.toLowerCase())) {
                Alert(`Select week days in correct format : Error comes from Row No ${index} In Data Range`)
                errorStatus = true
            }
        })
        count++
    }
    if (String(monthdays).length > 0) {
        let monthDays = monthdays.split(",")
        monthDays.forEach((item) => {
            if (Number(item) === 0 || item.length > 2 || Number(item) > 28) {
                Alert(`Select month days in correct format less than 29 and more than 0 : Error comes from Row No ${index} In Data Range`)
                errorStatus = true
            }

        })
        count++
    }
    if (count > 1) {
        Alert(`Select only one from from hour,minutes,days,weeks,year ,week days, and month days repeatation : Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    if (errorStatus)
        return true
}

//send whatsapp message with response buttons
function SendWhatsappMessageWithButtons(e: GoogleAppsScript.Events.TimeDriven) {
    let phone_id=PropertiesService.getScriptProperties().getProperty('phone_id')
    let triggers = findAllTriggersFromTriggersSheet().filter((trigger) => {
        if (trigger.trigger_id === e.triggerUid) {
            return trigger
        }
    })

    if (triggers.length > 0) {
        if (!AutoStop(String(triggers[0].phone))) {
            try {

                let token = PropertiesService.getScriptProperties().getProperty('accessToken')
                let url = `https://graph.facebook.com/v16.0/${phone_id}/messages`;
                let data = {
                    "messaging_product": "whatsapp",
                    "recipient_type": "individual",
                    "to": triggers[0].phone,
                    "type": "template",
                    "template": {
                        "name": "salary_reminder    ",
                        "language": {
                            "code": "en_US"
                        },
                        "components": [
                            {
                                "type": "header",
                                "parameters": [
                                    {
                                        "type": "image",
                                        // "text": triggers[0].work_title
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
                                        "text": triggers[0].work_title + " : " + triggers[0].work_detail
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
            catch (err) {
                console.log(err)
            }
        }
    }
}

//alert box
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

// refresh work status
function RefreshTrigger(index: number) {
    let date = new Date(sheet?.getRange(index, 13).getValue())
    let refresh_date: Date | null = null
    let phone = sheet?.getRange(index, 7).getValue()
    let name = sheet?.getRange(index, 6).getValue()
    let work_title = sheet?.getRange(index, 4).getValue()
    let work_detail = sheet?.getRange(index, 5).getValue()
    let mf = sheet?.getRange(index, 15).getValue();
    let hf = sheet?.getRange(index, 16).getValue();
    let df = sheet?.getRange(index, 17).getValue();
    let wf = sheet?.getRange(index, 18).getValue();
    let monthf = sheet?.getRange(index, 19).getValue();
    let yearf = sheet?.getRange(index, 20).getValue();
    if (!mf || typeof (mf) !== "number") mf = 0
    if (!hf || typeof (hf) !== "number") hf = 0
    if (!df || typeof (df) !== "number") df = 0
    if (!wf || typeof (wf) !== "number") wf = 0
    if (!monthf || typeof (monthf) !== "number") monthf = 0
    if (!yearf || typeof (yearf) !== "number") yearf = 0
    let weekdays = String(sheet?.getRange(index, 21).getValue())
    let monthdays = String(sheet?.getRange(index, 22).getValue())
    let triggers: GoogleAppsScript.Script.Trigger[] = []
    if (mf === 1) {
        let miliseconds = 1 * 30000
        refresh_date = new Date(date.getTime() + miliseconds)
    }
    if (mf > 1) {
        let miliseconds = (mf - 2) * 60000
        refresh_date = new Date(date.getTime() + miliseconds)
    }
    if (hf > 0) {
        let miliseconds = (hf * 60 - 10) * 60000
        refresh_date = new Date(date.getTime() + miliseconds)
    }
    if (df > 0) {
        let miliseconds = (df * 24 * 60 - 30) * 60000
        refresh_date = new Date(date.getTime() + miliseconds)
    }
    if (wf > 0) {
        let miliseconds = (wf * 7 * 24 * 60 - 30) * 60000
        refresh_date = new Date(date.getTime() + miliseconds)
    }
    if (monthf > 0) {
        let miliseconds = (monthf * GetMonthDays(date.getFullYear(), date.getMonth()) * 24 * 60 - 30) * 60000
        refresh_date = new Date(date.getTime() + miliseconds)
    }
    if (yearf > 0) {
        let miliseconds = (yearf * 365 * 24 * 60 - 30) * 60000
        refresh_date = new Date(date.getTime() + miliseconds)
    }
    if (weekdays) {
        let miliseconds = (7 * 24 * 60 - 30) * 60000
        refresh_date = new Date(date.getTime() + miliseconds)
    }
    if (monthdays) {
        let miliseconds = (GetMonthDays(date.getFullYear(), date.getMonth()) * 24 * 60 - 30) * 60000
        refresh_date = new Date(date.getTime() + miliseconds)
    }
    if (refresh_date) {
        let tr = ScriptApp.newTrigger('GetRefresh').timeBased().at(refresh_date).create();
        sheet?.getRange(index, 14).setValue(new Date(refresh_date))
        triggers.push(tr)
        triggers.forEach((trigger) => {
            SaveTrigger({
                date: date, trigger_id: trigger.getUniqueId(), trigger_type: trigger.getHandlerFunction(), phone: phone, name: name, work_title: work_title, work_detail: work_detail,
                mf: mf, hf: hf, df: df, wf: wf, monthf: monthf, yearf: yearf, weekdays: weekdays,
                monthdays: monthdays
            })
        })
    }

}

//refresh last date,refresh date,work status
function GetRefresh(e: GoogleAppsScript.Events.TimeDriven) {
    let triggers = findAllTriggersFromTriggersSheet().filter((trigger) => {
        if (trigger.trigger_id === e.triggerUid && trigger.trigger_type === "GetRefresh") {
            return trigger
        }
    })
    if (triggers.length > 0) {
        RefreshData(triggers[0].phone)
    }
}

//refresh data
function RefreshData(phone: string) {
    if (sheet) {
        for (let i = 3; i <= sheet.getLastRow(); i++) {
            let mobile = String(sheet?.getRange(i, 7).getValue())
            if (mobile === phone) {
                sheet?.getRange(i, 3).setValue("pending")
            }
        }
    }

}

// delete all triggers
function DeleteAllTriggers() {
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (trigger.getHandlerFunction() === "SendWhatsappMessageWithButtons" || trigger.getHandlerFunction() === "SetUpWhatsappTrigger" || trigger.getHandlerFunction() === "GetRefresh") {
            ScriptApp.deleteTrigger(trigger);
        }
    });
    sheet?.getRange(3, 1, sheet.getLastRow() - 2).clearContent()
    sheet?.getRange(3, 2, sheet.getLastRow() - 2).clearContent()
    let tsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('triggers')
    tsheet?.clear()
    tsheet?.getRange(1, 1).setValue("date").setBackground('yellow').setFontWeight('bold')
    tsheet?.getRange(1, 2).setValue("trigger id").setBackground('yellow').setFontWeight('bold')
    tsheet?.getRange(1, 3).setValue("trigger type").setBackground('yellow').setFontWeight('bold')
    tsheet?.getRange(1, 4).setValue("phone").setBackground('yellow').setFontWeight('bold')
    tsheet?.getRange(1, 5).setValue("name").setBackground('yellow').setFontWeight('bold')
    tsheet?.getRange(1, 6).setValue("work title").setBackground('yellow').setFontWeight('bold')
    tsheet?.getRange(1, 7).setValue("work detail").setBackground('yellow').setFontWeight('bold')
    tsheet?.getRange(1, 8).setValue("minutes").setBackground('yellow').setFontWeight('bold')
    tsheet?.getRange(1, 9).setValue("hours").setBackground('yellow').setFontWeight('bold')
    tsheet?.getRange(1, 10).setValue("days").setBackground('yellow').setFontWeight('bold')
    tsheet?.getRange(1, 11).setValue("weeks").setBackground('yellow').setFontWeight('bold')
    tsheet?.getRange(1, 12).setValue("months").setBackground('yellow').setFontWeight('bold')
    tsheet?.getRange(1, 13).setValue("years").setBackground('yellow').setFontWeight('bold')
    tsheet?.getRange(1, 14).setValue("weekdays").setBackground('yellow').setFontWeight('bold')
    tsheet?.getRange(1, 15).setValue("monthdays").setBackground('yellow').setFontWeight('bold')
}

function DuplicatePhoneChecker() {
    if (sheet) {
        let prev = "0066"
        for (let i = 3; i <= sheet.getLastRow(); i++) {
            let mobile = String(sheet?.getRange(i, 7).getValue())
            if (mobile === prev) {
                Alert(`Duplicate phone no exists :${mobile} Error comes from Row No ${i} In Data Range`)
            }
            prev = mobile
        }
    }
}


function AutoStop(phone: string) {
    let dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('test')
    let stop = false
    if (dataSheet) {
        for (let i = 3; i <= dataSheet.getLastRow(); i++) {
            let mobile = String(dataSheet?.getRange(i, 7).getValue())
            let autostop = String(dataSheet?.getRange(i, 12).getValue())
            if (mobile === phone) {
                if (autostop.toLowerCase() === "stop") {
                    stop = true
                }
            }
        }
    }
    return stop
}

