function addDogfoodSignature(eventObj) {
  let logoContent = "iVBORw0KGgoAAAANSUhEUgAAACIAAAAiCAYAAAA6RwvCAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAIISURBVFhH7Ze9bxNBEMXfbC5OkyIUSO44KUmNS6C6dEATu0JIUQiICkUi1BQxVaAiBXQgPgoiKuIyNCQNSun8AUgXKR2WsCkiBNwOc9aAyZH7QPZe5V/hm7eW5ae9N3ezhKJsdG7A8BpANV0ZEm7Lxy4qXhP3zvSKGXnY2QThrqoRI4YqXjChKp2NL77sxJYqB1AVkf1mVKVjbF0rl9TzjcDOaOEQqhUwQnGonFMwrJ/bIDqvygkFdkSY8gIwH6hywp8d8S+s+IbsoiHMsEX4af/1K/1qwKOOBJf//znSPfLpuHtO1an0jcxdWl6XSzOuf8OMMKKfjfDjm6EzYm69bRIh/o9UzOzFpUW5njARIz/0PXgvVDrHENEDrU+jJkbLeI7EYc3uBiIzondLNsW6pgTGRpKMjSQZG0kyNpKkFCOWkPPi5FY5O/L82jaDD1X9Q0T8uLRbYyNTT5qxQA9MN/Hs+h5Vr97flamE/5qRTnA8XQ2/np0PVRaH5bzy9HJL1YDbW4GxJiBQGNnJd3jZ6MXLhNUdMZFmYzhkuGpjsrKAzYWuLqUik6EbEzEyXNXw4/sHlZk4z0jfzJ33uTNNOWElG2iVimFOb6uRQZQbdiMB2dbaCRLYHrxKbk4MvKl1aRt3hyemNemafotmMWiY1Z343NGQhZEcLeWWtyBPTDy5sqdLGQC/AJM9h+Epch8hAAAAAElFTkSuQmCC";
  let logoName = "op-logo.png";

  let afterHoursDisclaimer = "";
  let today = new Date();
  let time = today.getHours();
  let day = today.getDay();
  if (day == 0 || day == 6 || time < 8 || time > 16) {
    afterHoursDisclaimer += "<br/>";
    afterHoursDisclaimer += "<span style='font-size:7.0pt'>After-hours responses are not required or expected!</span>";
  }

  let signature = "";
  signature += "<table>";
  signature +=   "<tr>";
  signature +=     "<td style='border-right: 1px solid #888888; padding-right: 5px;'><img src='cid:" + logoName + "' alt='MS Logo' width='24' height='24' /></td>";
  signature +=     "<td style='padding-left: 5px;'>" + Office.context.mailbox.userProfile.displayName + afterHoursDisclaimer + "</td>";
  signature +=   "</tr>";
  signature += "</table>";

  Office.context.mailbox.item.addFileAttachmentFromBase64Async(logoContent, logoName, { isInline: true }, function(result) {
    Office.context.mailbox.item.body.setSignatureAsync
    (
      signature,
      {
        "coercionType": "html",
        "asyncContext" : eventObj
      },
      function (asyncResult)
      {
        asyncResult.asyncContext.completed({ "key00" : "val00" });
      }
    );
  });
}

function set_signature(eventObj, signature)
{
    try
    {
      Office.context.mailbox.item.body.setSignatureAsync
      (
        signature,
        {
          "coercionType": "html",
          "asyncContext": eventObj
        },
        function (asyncResult)
        {
          console.log("Home.js - set_signature callback - invoked!");
        }
      );
    }
    catch (ex)
    {
      console.log(JSON.stringify(ex));
    }
}

function body_prepend_async(eventObj, text)
{
    try
    {
      Office.context.mailbox.item.body.prependAsync
      (
        text,
        {
          "coercionType": Office.CoercionType.Html,
          "asyncContext": eventObj
        },
        function (asyncResult)
        {
          console.log("Home.js - body_prepend_async callback - invoked!");
        }
      );
    }
    catch (ex)
    {
      console.log(JSON.stringify(ex));
    }
}

function onMessageCompose(eventObj)
{
  console.log("Home.js - onMessageCompose - got called at " + (new Date()).toString());
  body_prepend_async(eventObj, "<br>on_message_compose_created");
}

function onAppointmentOrganizer(eventObj)
{
  console.log("Home.js - onAppointmentOrganizer - got called at " + (new Date()).toString());
  body_prepend_async(eventObj, "<br>on_appointment_organizer_created");
}

function onMessageAttachmentsChanged(eventObj)
{
  console.log("Home.js - onMessageAttachmentsChanged - got called at " + (new Date()).toString());
  body_prepend_async(eventObj, "<br>on_message_attachments_changed");
}

function onAppointmentAttachmentsChanged(eventObj)
{
  console.log("Home.js - onAppointmentAttachmentsChanged - got called at " + (new Date()).toString());
  body_prepend_async(eventObj, "<br>on_appointment_attachments_changed");
}

function onMessageSend(eventObj)
{
  console.log("Home.js - onMessageSend - got called at " + (new Date()).toString());
  body_prepend_async(eventObj, "<br>on_message_send");
}

function onAppointmentSend(eventObj)
{
  console.log("Home.js - onAppointmentSend - got called at " + (new Date()).toString());
  body_prepend_async(eventObj, "<br>on_appointment_send");
}

function onMessageRecipientsChanged(eventObj)
{
  console.log("Home.js - onMessageRecipientsChanged - got called at " + (new Date()).toString());
  body_prepend_async(eventObj, "<br>on_message_recipients_changed");
}

function onAppointmentAttendeesChanged(eventObj)
{
  console.log("Home.js - onAppointmentAttendeesChanged - got called at " + (new Date()).toString());
  body_prepend_async(eventObj, "<br>on_appointment_attendees_changed");
}

function onAppointmentTimeChanged(eventObj)
{
  console.log("Home.js - onAppointmentTimeChanged - got called at " + (new Date()).toString());
  body_prepend_async(eventObj, "<br>on_appointment_time_changed");
}

function onAppointmentRecurrenceChanged(eventObj)
{
  console.log("Home.js - onAppointmentRecurrentChanged - got called at " + (new Date()).toString());
  body_prepend_async(eventObj, "<br>on_appointment_recurrence_changed");
}

function onInfoBarDismissClicked(eventObj)
{
  console.log("Home.js - onInfoBarDismissClicked - got called at " + (new Date()).toString());
  body_prepend_async(eventObj, "<br>on_infoBar_dismiss_clicked");
}

if(Office.actions){
  Office.actions.associate("taskpane.addDogfoodSignature", addDogfoodSignature);

  Office.actions.associate("taskpane.onMessageCompose", onMessageCompose);
  Office.actions.associate("taskpane.onAppointmentOrganizer", onAppointmentOrganizer);
  Office.actions.associate("taskpane.onMessageAttachmentsChanged", onMessageAttachmentsChanged);
  Office.actions.associate("taskpane.onAppointmentAttachmentsChanged", onAppointmentAttachmentsChanged);
  Office.actions.associate("taskpane.onMessageSend", onMessageSend);
  Office.actions.associate("taskpane.onAppointmentSend", onAppointmentSend);
  Office.actions.associate("taskpane.onMessageRecipientsChanged", onMessageRecipientsChanged);
  Office.actions.associate("taskpane.onAppointmentAttendeesChanged", onAppointmentAttendeesChanged);
  Office.actions.associate("taskpane.onAppointmentTimeChanged", onAppointmentTimeChanged);
  Office.actions.associate("taskpane.onAppointmentRecurrenceChanged", onAppointmentRecurrenceChanged);
  Office.actions.associate("taskpane.onInfoBarDismissClicked", onInfoBarDismissClicked);
}
