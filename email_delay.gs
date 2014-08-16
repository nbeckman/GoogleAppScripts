// Email Delay:
// Email delay computes various statistics on your email
// response times (median, 95%ile, 99%ile, mean, stdev). You can use it to track 
// how busy you are (or at least how long it takes you to reply to emails!). 
// - Response time is computed from responses you've written over the past XX days
//   plus emails in priority inbox that have not yet been responded to.
// - Output is written to a spreadsheet of your choosing.
// - To use, update the email address in 'myEmailAddress' and the spreadsheet in
//   'mySpreadsheetKey'
// - You can set it up to run daily, as I have, using "Resources" -> "Curent Project's Triggers"
//   and adding a daily (or whatever) trigger.
// - Needs spreadsheet & gmail permissions.

/////////////////////////////////////////////////
///          FIELDS YOU MUST CHANGE           ///
/////////////////////////////////////////////////

// Use your email address here.
var myEmailAddress = "men@gmail.com"

// Use your output spreadsheet here.
// Pulled from the spreadsheet's url:
// https://docs.google.com/spreadsheets/d/1uiY2vHH2xcWre7Z93L_IDvxpilNnQnIRHqxfWnZljJA/edit#gid=0
var mySpreadsheetKey = "1uiY2vHH2xcWre7Z93L_IDvxpilNnQnIRHqxfWnZljJA"

/////////////////////////////////////////////////
///          FIELDS YOU CAN CHANGE            ///
/////////////////////////////////////////////////

// When considering my responses to threads, the script will look back at emails
// newer than this period of time. (So by default, it will consider all emails in
// the past 30 days.) 
var emailThreadsDateRange = "30d"

// The search query that returns all the threads that will be considered when
// looking for existing responses. It's in "normal gmail query" format, i.e.,
// whatever you allowed to enter into the Gmail search bar is game here.
var threadsSearchQuery = "from:me newer_than:" + emailThreadsDateRange

// Caps the number of threads in the past emailThreadsDateRange 
// that are considered. Cannot exceed 500 by Google rule.
var maxEmailThreads = 250

// After some period of time of an email in your priority inbox,
// it becomes unreasonable to expect a response. :-) So after this
// amount of time in seconds computeOutstandingDelays() will not
// consider a thread.
var maxExpectedResponseElapsedSecs = 6 * 30 * 24 * 60 * 60  // 6 months

/////////////////////////////////////////////////
///                 THE CODE                  ///
/////////////////////////////////////////////////

function isFromMe(message) {
  return message.getFrom().indexOf(myEmailAddress) > -1
};

function getThreads() {
  return GmailApp.search(threadsSearchQuery, 0, maxEmailThreads)
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
        var request_date = new Date(last_message.getDate())
        var elapsed_secs = (Date.now() - request_date) / 1000
        if (elapsed_secs <= maxExpectedResponseElapsedSecs) {
          result.push(elapsed_secs)
        }
      }
    }
    // To calm down warning: "Service invoked too many times in a short 
    // time: gmail rateMax. Try Utilities.sleep(1000) between calls."
    Utilities.sleep(1000)
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
          var request_date = new Date(last_message.getDate())
          var response_date = new Date(message.getDate())
          var elapsed_secs = (response_date - request_date) / 1000
          result.push(elapsed_secs)
        }
      }
      last_message = message
    }
    // To calm down warning: "Service invoked too many times in a short 
    // time: gmail rateMax. Try Utilities.sleep(1000) between calls."
    Utilities.sleep(1000)
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

// Computes the k percentile delay (e.g., 99th percentile
// delay). Percentile is expected to be between 0 and 1, so
// 0.99 for 99th percentile.
function computePercentile(percentile, delays) {
  Logger.log("Computing %s percentile of %s elements", percentile, delays.length)
  delays.sort(function (a, b) {
   return a > b ? 1 : a < b ? -1 : 0;
  });
  var length_percentile = delays.length * percentile
  var length_percentile_ceiling = Math.ceil(length_percentile)
  if (length_percentile == length_percentile_ceiling) {
    Logger.log("length_percentile == length_percentile_ceiling: %s", length_percentile)
    // Use this value and the next one, adjusted for indices.
    var index = length_percentile - 1
    var index_plus_one = length_percentile
    if (index >= delays.length || index_plus_one >= delays.length) {
      Logger.log("index >= delays.length || index_plus_one >= delays.length: %s", index)
      return delays[delays.length - 1] 
    } else {
      Logger.log("Returning average from index %s, which is %s and %s", index, delays[index], delays[index_plus_one])
      return (delays[index] + delays[index_plus_one]) / 2.0 
    }
  } else {
    // The index we computed, adjusted for offsets, is the value we want.
    var index = length_percentile_ceiling - 1
    if  (index >= delays.length) {
      return delays[delays.length - 1] 
    } else {
      return delays[index]  
    }
  }
}

function computeMedian(delays) {
  delays.sort(function (a, b) {
   return a > b ? 1 : a < b ? -1 : 0;
  });
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

// Writes the current date, and median, 95%ile, 99th%ile, mean and stddev delays 
// to a new row in the spreadsheet with the given key.
function writeRowToSpreadsheet(spreadsheet_key, median, ninety_fifth, ninety_ninth, mean, stddev) {
  var spreadsheet = SpreadsheetApp.openById(spreadsheet_key)
  Logger.log("Writing to spreadsheet %s", spreadsheet.getName())
  
  var sheet = spreadsheet.getActiveSheet()
  var date = new Date()
  var date_string = (date.getMonth() + 1) + "/" + date.getDate() + "/" + date.getFullYear()
  var row_contents = [ date_string, median, ninety_fifth, ninety_ninth, mean, stddev ]
  sheet.appendRow(row_contents)
}

// Main function, ties all the above together. Gets delays on my inbox and
// oustanding delays on my priority inbox (which I tend to respond to).
// Publishes statistics on delays to a hardcoded spreadsheet.
function processInbox() {
  delays = computeDelays(getThreads())
  Logger.log("Delays: %s", delays)
  
  outstanding_delays = computeOutstandingDelays(getExpectedResponseThreads())
  Logger.log("Outstanding delays: %s", outstanding_delays)
  
  // Combine delays.
  delays = delays.concat(outstanding_delays)
  
  Logger.log("Median delay: %s", elapsedTimeString(computeMedian(delays)))
  Logger.log("50th percentile: %s", elapsedTimeString(computePercentile(0.50, delays)))
  Logger.log("95th percentile: %s", elapsedTimeString(computePercentile(0.95, delays)))
  Logger.log("99th percentile: %s", elapsedTimeString(computePercentile(0.99, delays)))
  Logger.log("Mean delay: %s", elapsedTimeString(computeMean(delays)))
  Logger.log("Variance: %s", elapsedTimeString(computeVariance(delays)))
  Logger.log("Standard Deviation: %s", elapsedTimeString(standardDeviation(delays)))
  
  writeRowToSpreadsheet(
    mySpreadsheetKey, 
    computePercentile(0.50, delays),
    computePercentile(0.95, delays),
    computePercentile(0.99, delays),
    computeMean(delays), 
    standardDeviation(delays))
};
