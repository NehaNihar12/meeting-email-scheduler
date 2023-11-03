# meeting-email-scheduler
Developed a script to automate the filtering and scheduling of meeting emails using Google Apps Script.

Steps to configure it on your machine
a. Open your Gmail and click on the gear icon in the upper right corner, then select "See all settings."

b. In the settings, go to the "Forwarding and POP/IMAP" tab, and enable "Enable POP for all mail" and "Enable IMAP."

d. Now, go to Google Drive and create a new Google Apps Script project. In the script editor copy paste the file `addMeetingsToCalendar.gs`

e. set up a trigger to run this function addToCalendar and moveICSFilesToMeetingsFolder periodically.

f. Replace YOUR_CALENDAR_ID in addToCalendar function with the id of the calendar where you want to schedule the meetings.

g. You can also Call moveICSFilesToMeetingsFolder Inside addToCalendar. But I Run Timers on Both Individually as this allows independent control and scheduling of each task.


