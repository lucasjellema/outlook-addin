Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = processData;
    document.getElementById("sendEmails").onclick = sendEmails;
    findPlaceHolders()
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
      document.getElementById("placeholders").innerHTML = "{{email}}, " + uniqueMatches.join(", ");
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
  }, (asyncResult) => {
    console.log(JSON.stringify(asyncResult));
  })
}

const askForNewMail = () => {
  // send message from iframe to parent - hopefully consumed by Chrome Extension to click on New Mail button in UI
  parent.postMessage({ action: "NEW_MAIL", sender: "MY_OUTLOOK_ADDIN", eventType: "outlookMailEvent", data: { subject: "THE SUBJECT" } }, '*');
}
