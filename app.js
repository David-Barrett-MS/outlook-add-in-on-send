/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady();

var statusInfo = "";

/**
 * Set notification on MailItem (overwrites any previous notification)
 * @param {Notification message to be set} message 
 */
async function SetNotification(message) {
    var infoMessage =
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: message,
      icon: "icon2",
      persistent: true
    };    
    Office.context.mailbox.item.notificationMessages.replaceAsync('OnSendInfo', infoMessage);
}

/**
 * Append the given status to the notification for the MailItem
 * @param {Message to be added to the status} message 
 * @returns 
 */
async function SetStatus(message) {
    if (statusInfo != "") {
        statusInfo = statusInfo + " | ";    
    }
    statusInfo = statusInfo + message;
    console.log(message);
    return SetNotification(statusInfo);
}

/**
 * Entry point for message send event handling code.  Can be called from OnMessageSend or ItemSend event.
 * @param {} event 
 */
function onMessageSendHandler(event) {
    SetStatus("onMessageSendHandler called");
    invokeWebAPIOnSend(event);
}

 /**
 * Makes the call to the requested REST API.
 * jsonData must be a string in JSON format.
 * @param {*} urlData The URL to call.
 * @param {*} jsonData The JSON data to pass to the REST API.
 * @returns 
 */
async function postRestApi(urlData, postData) {
    SetStatus("postRestAPI");

    const response = await fetch(urlData, {
        method: "POST",
        body: postData,
        headers: {
            "Content-type": "application/json; charset=UTF-8"
        }
    });
    SetStatus("postRestAPI done");
    return response;
}

 /**
 * Makes the call to the requested REST API.
 * jsonData must be a string in JSON format.
 * @param {*} urlData The URL to call.
 * @param {*} jsonData The JSON data to pass to the REST API.
 * @returns 
 */
 async function getRestApi(urlData) {
    const response = await fetch(urlData, {
        method: "GET",
        headers: {
          "Content-type": "text/plain; charset=UTF-8"
        }
      })
      return response;
}    


/**
 * Invoke web API and mark send event complete only when response received (or error occurs)
 * @param {*} event 
 * @returns 
 */
async function invokeWebAPIOnSend(event) {
    const apiURL = "https://apps1.daves.tips/TestAPI?SecondsToWait=5";
    const apigetURL = "https://apps1.daves.tips/TestAPI?ReplyDelay=15";
    const dialogURL = "https://apps1.daves.tips/onSendTest/dialog.html";
    
    const useFetch = true;
    const postData = "{\"FieldName\": \"Some data.\"}";

    SetStatus("invokeWebAPIOnSend");

    if (useFetch)
    {
        SetStatus("Fetch");

        postRestApi(apiURL,postData).then(result => {
            SetStatus("Fetch complete");
            if (result != null)
            {
                SetStatus("Fetch body");
                // Process the response
                var responseBody = result.text();
                SetStatus("Fetch done");
                console.log("Response: " + responseBody);
                event.completed({ allowEvent: true });
                SetStatus("Event complete");
            }
            else
            {
                console.log("No API response");
                Office.context.mailbox.item.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Failed to contact API' });
                event.completed({ allowEvent: false, errorMessage:"Failed to contact API" });
                SetStatus("Event complete");
            }
        });
        
        SetStatus("End invokeWebAPIOnSend");
        return;
    }

    var xhr = new XMLHttpRequest();
    SetStatus("XMLHttpRequest");
    xhr.onreadystatechange = function () {
        console.log("Ready state:" + this.readyState);
        console.log("Status: " + this.status);
        if (this.readyState == 4) {
            if (this.status == 200 ) {
                console.log("Status: " + this.status);
                console.log("ResponseText: " + this.responseText);

                event.completed({ allowEvent: true });
            }
            else {
                console.log("FAILED" + this.readyState);
                console.log("ResponseText: " + this.responseText);
                console.log("Status: " + this.status);
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Failed to contact API' });
                event.completed({ allowEvent: false, errorMessage:"Failed to contact API" });
            }
        }
    }     
    xhr.timeout = 30000;
    xhr.open("POST", apiURL, true);
    xhr.setRequestHeader("Content-Type", "application/json"); 
    xhr.send(postData);
    console.log("MessageURL", apiURL);
    console.log("Request", postData);
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);