
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
    const { entry } = JSON.parse(e.postData.contents)
    let message = ""
    let from = ""
    ServerLog(entry[0].changes[0].value)
    ServerLog(entry[0].changes[0].value.messages[0])
    let type = entry[0].changes[0].value.messages[0].type
    ServerLog("type" + type)
    ServerLog(JSON.stringify(entry))
    let token = ScriptProperty.getProperty('accessToken')
    if (entry.length > 0 && token) {
        message = entry[0].changes[0].value.messages[0].text.body
        from = entry[0].changes[0].value.messages[0].from
        sendTemplate1(token)
        if (type === "text")
         sendText(`hi we have recieved your message ${message}`, from, token)
        if (type === "button")
            {
            ServerLog(JSON.stringify(entry[0].changes[0].value.messages[0].button))
            ServerLog("successful")
            sendText(`Thanks for salary ${entry[0].changes[0].value.messages[0].button.text}`, from, token)
            }
        else{
            ServerLog("failed")
        }
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

function sendTemplate1(token: string) {
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
                            "text": "Sandeep"
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
