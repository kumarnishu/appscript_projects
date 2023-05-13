function doGet(e: GoogleAppsScript.Events.DoGet) {
    const ScriptProperty = PropertiesService.getScriptProperties()
    //Displays the text on the webpage.
    let mode = e.parameter["hub.mode"];
    let challange = e.parameter["hub.challenge"];
    let token = e.parameter["hub.verify_token"];
    if (mode && token) {
        if (mode === "subscribe" && token === ScriptProperty.getProperty('myToken')) {
            return ContentService.createTextOutput(challange)
        } else {
            return ContentService.createTextOutput(JSON.stringify({ error: 'Error message' })).setMimeType(ContentService.MimeType.JSON);
        }
    }
}

function doPost(e: GoogleAppsScript.Events.DoPost) {
    const ScriptProperty = PropertiesService.getScriptProperties()
    let token = ScriptProperty.getProperty('accessToken')
    const { entry } = JSON.parse(e.postData.contents)
    ServerLog(e.postData.contents)
    try {
        if (entry.length > 0 && token) {
            let message = ""
            let from = ""
            if (entry[0].changes[0].value.messages){
                let type = entry[0].changes[0].value.messages[0].type
                if(type){
                    switch (type) {
                        case "button": {
                            from = entry[0].changes[0].value.messages[0].from
                            let btnRes = entry[0].changes[0].value.messages[0].button.text
                            UpdateWorkStatus(from, btnRes)
                            sendText(`Response saved`, from, token)
                        }
                            break;
                        case "text": {
                            from = entry[0].changes[0].value.messages[0].from
                            message = entry[0].changes[0].value.messages[0].text.body
                            sendText(`hi we have recieved your message ${message}`, from, token)
                        }
                            break;
                        default: sendText(`failed to parse message ${message}`, from, token)
                    }
                }
                if (entry[0].changes[0].value.statuses) {
                    let status = String(entry[0].changes[0].value.statuses[0].status)
                    let phone = String(entry[0].changes[0].value.statuses[0].recipient_id)
                    UpdateWhatsappStatus(phone, status)
                }
            }
        
        }
    }
    catch (error) {
        ServerLog(error)
    }

}

function sendText(message: string, from: string, token: string) {
    let url = "https://graph.facebook.com/v16.0/103342876089967/messages";
    let data = {
        "messaging_product": "whatsapp",
        "to": from,
        "type": "text",
        "text": {
            "body": message
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


function UpdateWorkStatus(phone: string, response: string) {
    let dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('test')
    if (dataSheet) {
        for (let i = 3; i <= dataSheet.getLastRow(); i++) {
            let mobile = String(dataSheet?.getRange(i, 7).getValue())
            if (mobile === phone) {
                dataSheet?.getRange(i, 3).setValue(response.toLowerCase())
            }
        }
    }
}

function UpdateWhatsappStatus(phone: string, response: string) {
    let dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('test')
    if (dataSheet) {
        for (let i = 3; i <= dataSheet.getLastRow(); i++) {
            let mobile = String(dataSheet?.getRange(i, 7).getValue())
            if (mobile === phone) {
                dataSheet?.getRange(i, 2).setValue(response)
            }
        }
    }
}

function sendTemplate2(token: string) {
    let url = "https://graph.facebook.com/v16.0/103342876089967/messages";
    let data = {
        "messaging_product": "whatsapp",
        "recipient_type": "individual",
        "to": "917056943283",
        "type": "template",
        "template": {
            "name": "product_announcement",
            "language": {
                "code": "en_US"
            },
            "components": [
                {
                    "type": "body",
                    "parameters": [
                        {
                            "type": "text",
                            "text": "Nishu kumar"
                        },
                        {
                            "type": "text",
                            "text": "RockFord Article"
                        },
                        {
                            "type": "text",
                            "text": "https://www.agarson.com/"
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


function ServerLog(msg: string) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('logs');
    sheet?.getRange(sheet.getLastRow() + 1, 1).setValue(new Date().toLocaleString() + " : " + msg)
}
