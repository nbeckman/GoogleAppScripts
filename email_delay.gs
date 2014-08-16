// Email Delay:
// Email delay computes the mean, median and std. deviation of your email
// response times. You can use it to track how busy you are (or at least how
// long it takes you to reply to emails!). 
// - Response time is computed from actual responses in inbox, and emails in
//   priority inbox that have not yet been responded to.
// - Output is written to a spreadsheet of your choosing.
// - To use, update the email address in 'isFromMe()' and the spreadsheet in
//   'processInbox()'
// - You can set it up to run daily, as I have, using "Resources" -> "Curent Project's Triggers"
//   and adding a daily (or whatever) trigger.
// - Needs spreadsheet & gmail permissions.

function isFromMe(message) {
  // Use your email address here.
  return message.getFrom().indexOf("me@gmail.com") > -1
};

function getThreads() {
  return GmailApp.getInboxThreads()
};

function  getExpectedResponseThreads() {
  return GmailApp.getPriorityInboxThreads() 
}

// Attempts to compute the time outstanding for email responses. For each
// thread where the last message was not from me, return the difference in
// seconds between the time of the last message and now. In other words,
// if we expect me to respond to these emails, this is the minimum delay it
// will when I eventually do respond to them. (And as such, the given threads
// should be threads that I am likely to respond to.)
function computeOutstandingDelays(threads) {
  var result = []
  for (var i = 0; i < threads.length; ++i) {
    var thread = threads[i]
    if (thread.getMessageCount() > 0) {
      var last_message = thread.getMessages()[thread.getMessageCount() - 1]
      if (!isFromMe(last_message)) {
        Logger.log(
          "Message from %s at time %s is awaiting a response.",
          last_message.getFrom(), 
          last_message.getDate())
        var request_date = new Date(last_message.getDate())
        var elapsed_secs = (Date.now() - request_date) / 1000
        result.push(elapsed_secs)
      }
    }
  }
  return result
}

// Given a list of email threads, looks for messages from someone else
// followed by messages from me. Returns the delay, in seconds, of my
// responses.
function computeDelays(threads) {
  var result = []
  for (var i = 0; i < threads.length; i++) {
    // get all messages in a given thread
    var messages = threads[i].getMessages();
    // iterate over each message
    var last_message = null
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j]
      // Logger.log("Found message from %s", message.getFrom())
      if (last_message != null) {
        if (isFromMe(message) && !isFromMe(last_message)) {
          Logger.log(
            "Message from %s at time %s followed message from me at time %s",
            last_message.getFrom(), 
            last_message.getDate(), 
            message.getDate())
          var request_date = new Date(last_message.getDate())
          var response_date = new Date(message.getDate())
          var elapsed_secs = (response_date - request_date) / 1000
          Logger.log("Difference is %s seconds", elapsed_secs)
          result.push(elapsed_secs)
        }
      }
      last_message = message
    }
  }
  return result
};

function computeMean(delays) {
  var sum = 0;
  for (var i = 0; i < delays.length; ++i) {
    sum += delays[i] 
  }
  return sum / delays.length  
}

function computeMedian(delays) {
  delays.sort()
  if (delays.length == 0) {
    return -1
  }
  var half = Math.floor(delays.length / 2);
  if(delays.length % 2 == 1) {
    // odd
    return delays[half];
  } else {
    // even
    return (delays[half-1] + delays[half]) / 2.0;
  }
}

function computeVariance(delays) {
  var mean = computeMean(delays)
  var squared_differences = []
  for (var i = 0; i < delays.length; ++i) {
    var diff = delays[i] - mean
    var squared_diff = Math.pow(diff, 2)
    squared_differences.push(squared_diff)
  }
  return computeMean(squared_differences)
}

function standardDeviation(delays) {
  var variance = computeVariance(delays)
  return Math.sqrt(variance) 
}

// Given a number of seconds returns a time display string
// (which may be in minutes, hours, days, etc.
function elapsedTimeString(seconds) {
  var minutes = seconds / 60
  if (minutes < 1) {
    return seconds + " seconds" 
  }
  var hours = minutes / 60
  if (hours < 1) {
    return minutes + " minutes" 
  }
  var days = hours / 24
  if (days < 1) {
    return hours + " hours" 
  }
  var months = days / 30
  if (months < 1) {
    return days + " days" 
  }
  return months + " months"
}

// Writes the current date, and median, mean and std delays to a
// new row in the spreadsheet with the given key.
function writeRowToSpreadsheet(spreadsheet_key, mean, median, stddev) {
  var spreadsheet = SpreadsheetApp.openById(spreadsheet_key)
  Logger.log("Writing to spreadsheet %s", spreadsheet.getName())
  
  var sheet = spreadsheet.getActiveSheet()
  var date = new Date()
  var date_string = (date.getMonth() + 1) + "/" + date.getDate() + "/" + date.getFullYear()
  var row_contents = [ date_string, mean, median, stddev ]
  sheet.appendRow(row_contents)
}

// Main function, ties all the above together. Gets delays on my inbox and
// oustanding delays on my priority inbox (which I tend to respond to).
// Publishes it all to a 
function processInbox() {
  delays = computeDelays(getThreads())
  Logger.log("Delays: %s", delays)
  
  outstanding_delays = computeOutstandingDelays(getExpectedResponseThreads())
  Logger.log("Outstanding delays: %s", outstanding_delays)
  
  // Combine delays.
  delays = delays.concat(outstanding_delays)
  
  Logger.log("Mean delay: %s", elapsedTimeString(computeMean(delays)))
  Logger.log("Median delay: %s", elapsedTimeString(computeMedian(delays)))
  Logger.log("Variance: %s", elapsedTimeString(computeVariance(delays)))
  Logger.log("Standard Deviation: %s", elapsedTimeString(standardDeviation(delays)))
  
  // Use your output spreadsheet here.
  // Pulled from the spreadsheet's url:
  // https://docs.google.com/spreadsheets/d/1uiY2vHH2xcWre7Z93L_IDvxpilNnQnIRHqxfWnZljJA/edit#gid=0
  var key = "1uiY2vHH2xcWre7Z93L_IDvxpilNnQnIRHqxfWnZljJA"
  writeRowToSpreadsheet(
    key, computeMean(delays), computeMedian(delays), standardDeviation(delays))
};
