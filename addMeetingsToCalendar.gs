// function to return formatted datetime in 24-hr clock format
function customMoment(dateTimeStr, format) {
  const months = [
  "Jan", "Feb", "Mar", "Apr", "May", "Jun",
  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
  ];

  const dateTimeParts = dateTimeStr.match(/(\w{3}) (\w{3} \d{1,2}, \d{4}) (\d{1,2}:\d{2})([APap][Mm])?/);

  if (!dateTimeParts) {
    return "Invalid Date";
  }

  const [, day, date, time, ampm] = dateTimeParts;
  const dateObj = new Date(`${date} ${new Date().getFullYear()} ${time} ${ampm || ''}`);

  const formats = {
    YYYY: dateObj.getFullYear(),
    YY: dateObj.getFullYear() % 100,
    MMM: months[dateObj.getMonth()],
    MM: (dateObj.getMonth() + 1).toString().padStart(2, '0'),
    M: dateObj.getMonth() + 1,
    DD: dateObj.getDate().toString().padStart(2, '0'),
    D: dateObj.getDate(),
    ddd: day,
    H: dateObj.getHours().toString().padStart(2, '0'),
    mm: dateObj.getMinutes().toString().padStart(2, '0'),
    ss: dateObj.getSeconds().toString().padStart(2, '0'),
  };

  let result = format;
  for (const key in formats) {
    result = result.replace(key, formats[key]);
  }

  return result;
}

function moveICSFilesToMeetingsFolder() {
  // Get the "meetings" label
  var meetingsLabel = GmailApp.getUserLabelByName("meetings");
  
  // Get all threads with the "has:attachment" search query to filter emails with attachments
  var threads = GmailApp.search("has:attachment", 0, 100); // Adjust the number of threads as needed
  console.log(threads)
  
  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var messages = thread.getMessages();
    
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var subject = message.getSubject();
      
      // Check if the email has an attachment with the ".ics" file extension
      var attachments = message.getAttachments({ includeInlineImages: false, includeAttachments: true });
      
      for (var k = 0; k < attachments.length; k++) {
        var attachment = attachments[k];
        if (attachment.getName().toLowerCase().endsWith(".ics")) {
          // Move the thread to the "meetings" label
          thread.addLabel(meetingsLabel);
          break; // Move to the next thread
        }
      }
    }
  }
}

// main function to add meetings to calendar
function addToCalendar() {
  var label = GmailApp.getUserLabelByName("meetings");
  var threads = label.getThreads();
  // Load the Calendar API
  var calendar = CalendarApp.getCalendarById('YOUR_CALENDAR_ID'); // Replace with your calendar ID
  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var messages = thread.getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var subject = message.getSubject();
      console.log(subject)
      var body = message.getPlainBody();
      
      // Extract date and time information from the subject
      var dateMatchSubject = subject.match(/(\w{3} \w{3} \d{1,2}, \d{4} \d{1,2}:\d{2}(?:[APap][Mm])? - \d{1,2}:\d{2}(?:[APap][Mm])?)/);
      // var dateMatchBody = body.match(/(Friday \w{3} \d{1,2}, \d{4}.*\d{1,2}:\d{2}(?:[APap][Mm])? – \d{1,2}:\d{2}(?:[APap][Mm])?)/);
      var dateMatchBody = body.match(/(.*\d{1,2}:\d{2}(?:[APap][Mm])? – \d{1,2}:\d{2}(?:[APap][Mm])?.*)/);

      
      if (dateMatchSubject && dateMatchSubject.length > 1) {
        var dateTimeStr = dateMatchSubject[1];
      } else if (dateMatchBody && dateMatchBody.length > 0) {
        var dateTimeStr = dateMatchBody[0];
      } else {
        continue; // Skip this email if no date and time information is found
      }
      
        // Parse the date and time string into JavaScript Date objects
        // Parse the date and time string into JavaScript Date objects

      // Extract start and end times using regular expressions
      var timePattern = /(\d{1,2}:\d{2}[APap][Mm])/g;
      var timeMatches = dateTimeStr.match(timePattern);
      if (timeMatches && timeMatches.length === 2) {
        // Assuming the first time match is the start time and the second is the end time
        var startTime = timeMatches[0];
        var endTime = timeMatches[1];

        // Extract the date from the date-time string
        var datePattern = /\w{3} \w{3} \d{1,2}, \d{4}/;
        var dateMatch = dateTimeStr.match(datePattern);
        if (dateMatch) {
          // Concatenate the date with the time to create the full date-time strings
          var startDateStr = dateMatch[0] + ' ' + startTime;
          var endDateStr = dateMatch[0] + ' ' + endTime;
          startDateStr= customMoment(startDateStr, "ddd MMM D, YYYY H:mm");
          endDateStr = customMoment(endDateStr, "ddd MMM D, YYYY H:mm");
          var startDate = new Date(startDateStr);
          var endDate = new Date(endDateStr);
          

          // Validate the extracted dates
          if (!isNaN(startDate) && !isNaN(endDate)) {
            // Create the event in Google Calendar
            console.log(startDate,'startDateStr',endDate)
            var event = calendar.createEvent(subject, startDate, endDate, { description: subject });

            // Optionally, remove the email label after adding it to the calendar to avoid duplication
            thread.removeLabel(label);
          }
        }
      }

    } 
  }
}
addToCalendar();
