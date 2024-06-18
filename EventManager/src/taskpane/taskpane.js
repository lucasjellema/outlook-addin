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

function run() {
  const item = Office.context.mailbox.item;
  const itemType = Office.context.mailbox.item.itemType;
  switch (itemType) {
    case Office.MailboxEnums.ItemType.Appointment:
      console.log(`Current item is an ${itemType}.`);
      writeEventDetails(item);
      break;
    case Office.MailboxEnums.ItemType.Message:
      console.log(`Current item is a ${itemType}. A message could be an email, meeting request, meeting response, or meeting cancellation.`);
      break;
  }
}

function writeEventDetails(item) {
  Office.context.mailbox.item.subject.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const subject = asyncResult.value;
      document.getElementById("event-subject").innerHTML = "<b>Subject:</b> <br/>" + subject
    }
  });

  Office.context.mailbox.item.start.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const start = asyncResult.value;
      document.getElementById("event-start").innerHTML = "<b>Start:</b> <br/>" + start
    }
  });
  Office.context.mailbox.item.end.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const end = asyncResult.value;
      document.getElementById("event-end").innerHTML = "<b>End:</b> <br/>" + end
    }
  });
  Office.context.mailbox.item.location.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const location = asyncResult.value;
      document.getElementById("event-location").innerHTML = "<b>Location:</b> <br/>" + location
    }
  });


  Office.context.mailbox.item.organizer.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const apptOrganizer = asyncResult.value;
      document.getElementById("event-organizer").innerHTML = "Organizer: " + apptOrganizer.displayName + " (" + apptOrganizer.emailAddress + ")"
    }
  });

  getAttendees()
}

const eventResponses = []

function getAttendees() {
  // This snippet gets an appointment's required and optional attendees and groups them by their response.
  const appointment = Office.context.mailbox.item;
  let attendees;
  // clear event responses
  eventResponses.length = 0
  if (Object.keys(appointment.organizer).length === 0) {
    // Get attendees as the meeting organizer.
    // Ensure this runs in the context of an appointment
    if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      appointment.requiredAttendees.getAsync((result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.log(result.error.message);
          return;
        }
        attendees = result.value;
        appointment.optionalAttendees.getAsync((result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.log(result.error.message);
            return;
          }
          attendees = attendees.concat(result.value);
          eventResponses.push(attendees)
          let acceptedCount = 0
          let declinedCount = 0
          let tentativeCount = 0
          for (const attendee of attendees) {
            if (attendee.appointmentResponse === Office.MailboxEnums.ResponseType.Accepted) {
              acceptedCount++
            }
            if (attendee.appointmentResponse === Office.MailboxEnums.ResponseType.Declined) {
              declinedCount++
            }
            if (attendee.appointmentResponse === Office.MailboxEnums.ResponseType.Tentative) {
              tentativeCount++
            }
          }
          document.getElementById("event-attendee-count").innerHTML = "<b>Attendee count:</b> <br/> Accepted: "
            + acceptedCount + " / Tentative:" + tentativeCount
            + " / Declined:" + declinedCount + "  / " + attendees.length
          document.getElementById("copyAttendees").onclick = function () { copyAttendees(attendees) }
          document.getElementById("copyAttendees").style = "display:block"

        });
      });
    } else {
      // Get attendees as a meeting attendee.
      attendees = appointment.requiredAttendees;
      attendees = attendees.concat(appointment.optionalAttendees);
      eventResponses.push(attendees)
    }
  }
}

const copyAttendees = (attendees) => {

  const includeNone = document.getElementById('includeNone').checked
  const includeDeclined = document.getElementById('includeDeclined').checked
  // copy eventResponses to clipboard 
  let text = ""
  const fieldSeparator = "\t"
  const recordSeparator = "\n"


  if (attendees.length > 0)
    if (attendees[0].appointmentResponse) {


      for (const attendee of attendees.sort((a, b) => a.appointmentResponse.localeCompare(b.appointmentResponse)
        || a.displayName.localeCompare(b.displayName))) {
        if (attendee.appointmentResponse === Office.MailboxEnums.ResponseType.Accepted
          || attendee.appointmentResponse === Office.MailboxEnums.ResponseType.Tentative
          || (attendee.appointmentResponse === Office.MailboxEnums.ResponseType.None && includeNone)
          || (attendee.appointmentResponse === Office.MailboxEnums.ResponseType.Declined && includeDeclined)
        )
          text += attendee.displayName.split(" ")[0] + fieldSeparator + attendee.displayName + fieldSeparator + attendee.emailAddress
            + fieldSeparator + attendee.appointmentResponse + recordSeparator
      }
      document.getElementById("eventDetails").innerHTML = text
      window.focus();
      navigator.clipboard.writeText(text)
    } else {  // in case we are looking at an appointment in editing mode - when no RSVPs are available, only a list of all people invited
      try {
        if (attendees[0].recipientType) {
          const fieldSeparatorAlt = ","
          const recordSeparatorAlt = ";"

          for (const attendee of attendees.filter(attendee => attendee.recipientType === "other").sort((a, b) => a.displayName.localeCompare(b.displayName))) {
            text += attendee.displayName.split(" ")[0] + fieldSeparatorAlt + attendee.displayName + fieldSeparatorAlt + attendee.emailAddress
              + fieldSeparatorAlt + "unknown" + recordSeparatorAlt
          }
          window.focus();
          document.getElementById("eventDetails").innerHTML = text
          navigator.clipboard.writeText(text)
        }
      } catch (error) {
        console.log('alternative processing of invited attendees failed', error)
      }
    }
}

function organizeByResponse(attendees) {
  const accepted = [];
  const declined = [];
  const noResponse = [];
  const tentative = [];
  attendees.forEach(attendee => {
    switch (attendee.appointmentResponse) {
      case Office.MailboxEnums.ResponseType.Accepted:
        accepted.push(attendee);
        break;
      case Office.MailboxEnums.ResponseType.Declined:
        declined.push(attendee);
        break;
      case Office.MailboxEnums.ResponseType.None:
        noResponse.push(attendee);
        break;
      case Office.MailboxEnums.ResponseType.Tentative:
        tentative.push(attendee);
        break;
      case Office.MailboxEnums.ResponseType.Organizer:
        console.log(`Organizer: ${attendee.displayName}, ${attendee.emailAddress}`);
        break;
    }
  });

}

// function printAttendees(attendees) {
//   let text = "Accepted: " + attendees.length
//   if (attendees.length === 0) {
//     console.log("None");
//   } else {
//     for (const attendee of attendees) {
//       text += `XX ${attendee.displayName}, ${attendee.emailAddress}, ${attendee.appointmentResponse}`
//     }
//     // attendees.forEach(attendee => {
//     //   console.log(` ${attendee.displayName}, ${attendee.emailAddress}`);
//     //   text += `XX ${attendee.displayName}, ${attendee.emailAddress}`
//     //   document.getElementById("eventDetails").innerHTML = text
//     // });

//   }
//   document.getElementById("eventDetails").innerHTML = text
// }