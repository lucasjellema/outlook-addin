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
    updateMeetingList()
  }
});


// set list of meetings in select
function updateMeetingList() {
  const meetingList = document.getElementById("meetings");
  let listOfMeetings = getFromLocalStorage(listOfMeetingsKey);
  if (listOfMeetings) {
    listOfMeetings = JSON.parse(listOfMeetings);
    if (!listOfMeetings || listOfMeetings.length === 0) return
    for (let meeting of listOfMeetings.sort((a, b) => a.timeOfCreation - b.timeOfCreation)) {
      const meetingOption = document.createElement("option");
      meetingOption.value = meeting;
      meetingOption.text = meeting.meeting;
      meetingList.appendChild(meetingOption);
    }
  }
}
function run() {
  const item = Office.context.mailbox.item;

  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const bodyText = result.value;
      const meeting = item.subject.substring(item.subject.lastIndexOf(": ") + 1);
      const rsvp = item.itemClass.substring(item.itemClass.lastIndexOf(".") + 1);

      updateMeetingDetails(meeting, item.sender, item.dateTimeModified, rsvp, bodyText, item.conversationId)
      document.getElementById("rsvpResult").innerHTML = "<b>RSVP status:</b> <br/>" + rsvp;
      document.getElementById("meeting").innerHTML = "<b>Meeting:</b> <br/>" + meeting + " conv id " + item.conversationId;
      document.getElementById("sender").innerHTML = "<b>Sender:</b> <br/>" + item.sender.displayName + "<br/>" + item.sender.emailAddress;
      document.getElementById("timestamp").innerHTML = "<b>Timestamp:</b> <br/>" + item.dateTimeModified;
      // The message class of the response you send depends on the value you specify in the RespondType parameter. It is IPM.Schedule.Meeting.Resp.Pos if you accept, IPM.Schedule.Meeting.Resp.Neg if you decline, or IPM.Schedule.Meeting.Resp.Tent if you accept tentatively.
      document.getElementById("messageProperties").innerHTML = bodyText
      displayMeetingDetails(meeting, item.conversationId)

      document.getElementById("copyAttendees").onclick = () => { copyAttendees(item.conversationId) }
      document.getElementById("copyAttendees").style = "display:block"
    } else {
      console.error('Failed to get the body of the email.');
    }
  });
}

const keyPostFix = "MeetingDetailsKey"
const listOfMeetingsKey = "ListOfMeetingsKey"

const addMeetingToList = (meeting, conversationId) => {
  let listOfMeetings = getFromLocalStorage(listOfMeetingsKey);
  if (!listOfMeetings) {
    listOfMeetings = [];
  } else {
    listOfMeetings = JSON.parse(listOfMeetings);
  }
  const meetingObject = {
    conversationId: conversationId,
    meeting: meeting,
    status: "new",
    timeOfCreation: new Date()
  }
  const index = listOfMeetings.findIndex((obj) => obj.conversationId === conversationId);
  if (index !== -1) {
    // listOfMeetings[index] = meetingObject  update the existing entry or not? I think not
  } else {
    listOfMeetings.push(meetingObject)
  }
  setInLocalStorage(listOfMeetingsKey, JSON.stringify(listOfMeetings));
}

const updateMeetingDetails = (meeting, sender, timestamp, rsvp, message, conversationId) => {
  const meetingDetailsKey = conversationId + keyPostFix
  let meetingDetails = getFromLocalStorage(meetingDetailsKey);
  if (!meetingDetails) {
    meetingDetails = {
      rsvps: {},
    };

  } else {
    meetingDetails = JSON.parse(meetingDetails);
  }
  meetingDetails.rsvps[sender.emailAddress] = {
    sender: sender,
    timestamp: timestamp,
    rsvp: rsvp,
    message: message
  }
  setInLocalStorage(meetingDetailsKey, JSON.stringify(meetingDetails));
  // determine number of rsvps, per type (pos, neg, ten)

  let rsvpCount = 0
  let count = { Pos: 0, Neg: 0, Ten: 0 }
  Object.keys(meetingDetails.rsvps).forEach((key) => {
    rsvpCount += 1
    count[meetingDetails.rsvps[key].rsvp] += 1
  })
  notify("Details updated - # of rsvps " + rsvpCount + " " + JSON.stringify(count))

  addMeetingToList(meeting, conversationId)
}

const displayMeetingDetails = (meeting, conversationId) => {
  const meetingDetailsKey = conversationId + keyPostFix
  let meetingDetails = getFromLocalStorage(meetingDetailsKey);
  if (meetingDetails) {
    document.getElementById("meetingProperties").innerHTML = meetingDetails
  }
}


function setInLocalStorage(key, value) {
  const myPartitionKey = Office.context.partitionKey;
  // Check if local storage is partitioned. 
  // If so, use the partition to ensure the data is only accessible by your add-in.
  if (myPartitionKey) {
    localStorage.setItem(myPartitionKey + key, value);
  } else {
    localStorage.setItem(key, value);
  }
}

function getFromLocalStorage(key) {
  const myPartitionKey = Office.context.partitionKey;
  // Check if local storage is partitioned.
  if (myPartitionKey) {
    return localStorage.getItem(myPartitionKey + key);
  } else {
    return localStorage.getItem(key);
  }
}

// write informational message on top of message pane
const notify = (message) => {
  const notification = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: message,
    icon: "Icon.80x80",
    persistent: true,
  };
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", notification);
}


const copyAttendees = (conversationId) => {
  const meetingDetailsKey = conversationId + keyPostFix
  const includeDeclined = document.getElementById('includeDeclined').checked
  // copy eventResponses to clipboard 
  let text = ""
  const fieldSeparator = "\t"
  const recordSeparator = "\n"

  let meetingDetails = getFromLocalStorage(meetingDetailsKey);
  rsvps = JSON.parse(meetingDetails).rsvps;

  Object.keys(rsvps).forEach((key) => {
    console.log("key",key)
    console.log("rsvps", rsvps)
    const response = rsvps[key]
    if (response) {

    // for (const attendee of attendees.sort((a, b) => a.appointmentResponse.localeCompare(b.appointmentResponse) 
    //                                                 || a.displayName.localeCompare(b.displayName) ) ){
    //
    if (response.rsvp === 'Pos' || response.rsvp === 'Ten' || (response.rsvp === 'Neg' && includeDeclined)) {

      text += response.sender.displayName.split(" ")[0] + fieldSeparator + response.sender.displayName + fieldSeparator + response.sender.emailAddress
        + fieldSeparator + response.rsvp+ fieldSeparator + response.message + recordSeparator
    }
  }
  })

  document.getElementById('csv').value = text
  copyTextToClipboard(text)

  
}


function copyTextToClipboard(text) {
  if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(() => {
          console.log('Text copied to clipboard successfully!');
      }).catch(err => {
          console.error('Failed to copy text to clipboard: ', err);
          fallbackCopyTextToClipboard(text);
      });
  } else {
      fallbackCopyTextToClipboard(text);
  }
}

function fallbackCopyTextToClipboard(text) {
  const textArea = document.createElement("textarea");
  textArea.value = text;
  
  // Avoid scrolling to bottom
  textArea.style.top = "0";
  textArea.style.left = "0";
  textArea.style.position = "fixed";

  document.body.appendChild(textArea);
  textArea.focus();
  textArea.select();

  try {
      const successful = document.execCommand('copy');
      const msg = successful ? 'successful' : 'unsuccessful';
      console.log('Fallback: Copying text command was ' + msg);
  } catch (err) {
      console.error('Fallback: Oops, unable to copy', err);
  }

  document.body.removeChild(textArea);
}