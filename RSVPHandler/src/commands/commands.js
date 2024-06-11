/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const item = Office.context.mailbox.item;
  console.log("action", item);
  console.log("action", "fetch body ");

  item.body.getAsync(Office.CoercionType.Html, (result) => {
    console.log("body text", item);
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log("body text success", item);
      const bodyText = result.value;
      const meeting = item.subject.substring(item.subject.lastIndexOf(": ") + 1);
      const rsvp = item.itemClass.substring(item.itemClass.lastIndexOf(".") + 1);

      updateMeetingDetails(meeting, item.sender, item.dateTimeModified, rsvp, bodyText, item.conversationId)
      console.log("updated meeting", item);

      // Be sure to indicate when the add-in command function is complete.
      event.completed();

    } else {
      console.error('Failed to get the body of the email.');
      event.completed();
    }
  })
  
}

// Register the function with Office.
Office.actions.associate("action", action);



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
