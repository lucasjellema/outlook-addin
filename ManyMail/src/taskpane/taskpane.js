/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = processData;
    document.getElementById("sendEmails").onclick = sendEmails;
    findPlaceHolders()
    //  document.getElementById("newMail").onclick = askForNewMail;
  }
});

// find all unique occurrences of {{}} in the subject and body
function findPlaceHolders() {
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const subject = Office.context.mailbox.item.subject;
      const body = result.value;
      const regex = /\{\{.*?\}\}/g;
      const uniqueSubjectMatches = [...subject.matchAll(regex)].map(match => match[0]);
      const uniqueBodyMatches = [...body.matchAll(regex)].map(match => match[0]);
      const uniqueMatches = [...new Set([...uniqueSubjectMatches, ...uniqueBodyMatches])];
      document.getElementById("placeholders").innerHTML = "{{email}}, "+uniqueMatches.join(", ");
    }
  });
}

// Function to convert CSV string to object array
function csvToObjectArray(csvString) {
  const rows = csvString.trim().split('\n');
  const headers = rows[0].split(',');
  return rows.slice(1).map(row => {
    const values = row.split(',');
    let obj = {};
    values.forEach((value, index) => {
      obj[headers[index]] = value;
    });
    return obj;
  });
}

// Function to create an HTML table from an object array
function createTableFromData(data) {
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');

  // Create header row
  const headerRow = document.createElement('tr');
  Object.keys(data[0]).forEach(key => {
    const th = document.createElement('th');
    th.textContent = key;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  // Create data rows
  data.forEach(item => {
    const row = document.createElement('tr');
    Object.values(item).forEach(value => {
      const td = document.createElement('td');
      td.textContent = value;
      row.appendChild(td);
    });
    tbody.appendChild(row);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  return table;
}

// Handle paste event on textarea
const processData = () => {


  const recipientsTextarea = document.getElementById('recipients');

  const data = recipientsTextarea.value;
  const parsedData = csvToObjectArray(data);
  recipientsData = parsedData
  const tableContainer = document.getElementById('tableContainer');
  tableContainer.innerHTML = ''; // Clear previous table
  const table = createTableFromData(parsedData);
  tableContainer.appendChild(table);
};

let recipientsData
/*
 
firstName,lastName,email,company
John,Smith,4hjO5@example.com,Microsoft
Jane,Doe,3lT9H@example.com,Google
*/


const askForNewMail = () => {
  // send message from iframe to parent - hopefully consumed by Chrome Extension to click on New Mail button in UI
  parent.postMessage({ action: "NEW_MAIL", sender: "MY_OUTLOOK_ADDIN", eventType: "outlookMailEvent", data: { subject: "THE SUBJECT" } }, '*');
}


function run() {

  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
}

const personalize = (theString, recipient) => {
  // for all properties in the recipient replace each occurrence of {{property}} in theString with the value of that property
  let workstring = theString
  Object.keys(recipient).forEach((key) => {
    const value = recipient[key];
    const regex = new RegExp(`{{${key}}}`, 'g');
    workstring = workstring.replace(regex, value);
  })
  return workstring
}


function sendEmails() {
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emailBody = result.value;
      const subject = Office.context.mailbox.item.subject;
      recipientsData.forEach((recipient, index) => {
        const personalizedBody = personalize(emailBody, recipient)
        const personalizedSubject = personalize(subject, recipient)
        // wait some time between mails
        setTimeout(() => {
          notify("Sending email to " + recipient.email + " (" + index + " of " + recipientsData.length + ")");
          saveEmail(recipient.email, personalizedSubject, personalizedBody)
        }, index * 1200); // saveEmail(recipient, personalizedBody);
      });
    }
  });
}

const notify = (message) => {
  const notification = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: message,
    icon: "Icon.80x80",
    persistent: true,
  };
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", notification);
}

function saveEmail(recipient, subject, body) {

  Office.context.mailbox.displayNewMessageFormAsync({
    toRecipients: [recipient], // Copies the To line from current item
    // ccRecipients: ["sam@contoso.com"],
    bccRecipients: ["autosend@contoso.com"], // signal to Chrome Extension that it can automatically send this mail
    subject: subject,
    htmlBody: body
    // , just an attempt - does not work
    // attachments: [
    //   {
    //     type: "file",
    //     name: "image.png",
    //     url: "http://www.cutestpaw.com/wp-content/uploads/2011/11/Cute-Black-Dogs-s.jpg",
    //     isInline: true
    //   }
    // ]
  }, (asyncResult) => {
    console.log(JSON.stringify(asyncResult));
    // write asyncResult to task pane
    // find DIV element with id == notification
    // set its textcontent tpo asyncResult
    // const div = document.getElementById("notification");
    // div.textContent = JSON.stringify(asyncResult);


    //  Office.context.mailbox.item.to.setAsync([{ emailAddress: recipient }], (result) => {
    //   if (result.status === Office.AsyncResultStatus.Succeeded) {
    //   Office.context.mailbox.item.body.setAsync(body, { coercionType: Office.CoercionType.Html }, (result) => {
    //   message.message = "update body for "+recipient;
    // Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
    // if (result.status === Office.AsyncResultStatus.Succeeded) {
    // message.message = "update body succes, now save  "+recipient;

    // Office.context.mailbox.item.saveAsync((result) => {
    //   if (result.status === Office.AsyncResultStatus.Succeeded) {
    //     notify("save success ");
    //     Office.context.mailbox.item.close();
    //     notify("Closed, next");
    //   }
    // });
  })
}


//       });
//     }
//   });
// }




function sendEmail(recipient, subject, body) {

  const soapMessage = `
  <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Header>
      <t:RequestServerVersion Version="Exchange2013" />
    </soap:Header>
    <soap:Body>
      <m:CreateItem MessageDisposition="SendAndSaveCopy">
        <m:Items>
          <t:Message>
            <t:Subject>${subject}</t:Subject>
            <t:Body BodyType="HTML">${body}</t:Body>
            <t:ToRecipients>
              <t:Mailbox><t:EmailAddress>${recipient}</t:EmailAddress></t:Mailbox>
            </t:ToRecipients>
          </t:Message>
        </m:Items>
      </m:CreateItem>
    </soap:Body>
  </soap:Envelope>`;

  Office.context.mailbox.makeEwsRequestAsync(soapMessage, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log('Email sent successfully.');
    } else {
      console.error('Error sending email:', result.error.message);
    }
  });
}


