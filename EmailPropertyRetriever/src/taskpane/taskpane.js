/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
 // Get a reference to the current message
 const item = Office.context.mailbox.item;

 // Write message property value to the task pane

 document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
 document.getElementById("item-from").innerHTML = "<b>From:</b> <br/>" + item.sender.emailAddress + " " + item.sender.displayName
 const recipients = item.to.reduce((a, b) => a + b.emailAddress + ", ", "")
 document.getElementById("item-to").innerHTML = "<b>To:</b> <br/>" + recipients.slice(0, -2)
 document.getElementById("item-timestamp").innerHTML = "<b>Timestamp:</b> <br/>" + item.dateTimeCreated


  item.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
   if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
     document.getElementById("item-body").innerHTML = "<b>Body:</b> <br/>" + bodyResult.value;
   }
 })
}
