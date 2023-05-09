
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
    let buttonMessage = ""
    let token = ScriptProperty.getProperty('accessToken')
    ServerLog(message + buttonMessage)
    if (entry.length > 0 && token) {
        message = entry[0].changes[0].value.messages[0].text.body
        buttonMessage = entry[0].changes[0].value.messages[0].button.payload
        from = entry[0].changes[0].value.messages[0].from
        ServerLog(message + buttonMessage)
        if (buttonMessage === "Sent")
            sendTextMessage(`thanks for sending salary`, from, token)
        sendTextMessage(`hi we have recieved your message ${message}`, from, token)
    }
}

function sendTextMessage(message: string, from: string, token: string) {
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


function ServerLog(msg: string) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('logs');
    sheet?.getRange(sheet.getLastRow() + 1, 1).setValue(new Date().toLocaleString() + " : " + msg)
}
