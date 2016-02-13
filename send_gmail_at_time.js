// This script sends email on a delayed schedule.
// You write an draft email in gmail, and give it one of the labels defined
// in LABELS_BY_HOUR_OF_DAY. The script should be set up to run on an hourly
// trigger (the mainFunction() function). If the current hour in the
// TIMEZONE timezone is equal to any of the hours defined in LABELS_BY_HOUR_OF_DAY,
// all draft emails with that label will be send.

// A map from hour of day to the label whose messages should be sent during
// that hour of day. Add and remove labels as needed.
var LABELS_BY_HOUR_OF_DAY = {10 : "Delayed Send/10am",
                             14 : "Delayed Send/2pm"}

// The timezone where labels will apply. I live in this timezone, but you can
// change the value.
var TIMEZONE = "America/New_York"

// Returns the 'full' username to use in the from field, so recipients see
// more than the simple email address.
function getUserFullName() {
  // I found this approach in StackOverflow:
  // http://stackoverflow.com/a/32081910/250096
  var email = Session.getActiveUser().getEmail();
  var contact = ContactsApp.getContact(email);
  return contact.getFullName();
}

// Sends the given message by forwarding it to the same person it was
// initially addressed to, 
function sendMessage(message) {
  // Just forward the message with all the same parameters.
  // I found this approach in StackOverflow:
  // http://stackoverflow.com/a/32081910/250096
  message.forward(message.getTo(), {
    cc: message.getCc(),
    bcc: message.getBcc(),
    subject: message.getSubject(),
    name: getUserFullName()
  });
}

// This is the main function that will be run hourly.
function mainFunction() {
  // 1 - get current hour of day.
  var date = new Date()
  var hour_string = Utilities.formatDate(date, TIMEZONE, "HH")
  var hour = parseInt(hour_string)
  Logger.log("Current hour of day: " + hour)
  
  // 2 - search map for labels you can fire in the same hour of day.
  if (!(hour in LABELS_BY_HOUR_OF_DAY)) {
    Logger.log("Did not find hour (" + hour + 
               ") in LABELS_BY_HOUR_OF_DAY, returning.")
    return
  }
  var label_name = LABELS_BY_HOUR_OF_DAY[hour]
  Logger.log("Found label for hour of day: " + hour + " --> " + label_name)
  
  // 3 - look up threads with label.
  var label = GmailApp.getUserLabelByName(label_name);
  if (label == null) {
    Logger.log("Could not find label: " + label_name)
    return 
  }
  var label_threads = label.getThreads()
  Logger.log("Found " + label_threads.length + " threads in label " + label_name)
  
  // 4 - for each thread in label, for each message in thread.
  for (var i = 0; i < label_threads.length; ++i) {
    var label_thread = label_threads[i]
    for (var j = 0; j < label_thread.getMessages().length; ++j) {
      var message = label_thread.getMessages()[j] 
      // 5 - Only send drafs. We don't want to duplicate anything.
      if (message.isDraft()) {
        Logger.log("Found draft! Subject: " + message.getSubject())
        // 6 - Don't send message if it's not addressed to anyone.
        if (!message.getTo().trim() && !message.getCc().trim() && 
            !message.getBcc().trim()) {
          Logger.log("Message was not to anyone. Ignoring.")
          continue
        }
        // 7 - Send!
        sendMessage(message)
        Logger.log("Sent message " + message.getSubject())
        // 8 - remove label. This may not be needed considering we are
        //     moving the message to the trash, but we want to get it
        //     definitively out of the way.
        label_thread.removeLabel(label)
        // 9 - Move to trash. We've got the send email, so we don't care
        //     about this any more.
        message.moveToTrash()
      }
    }
  }
}
