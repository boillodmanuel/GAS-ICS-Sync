/*
*=========================================
*       INSTALLATION INSTRUCTIONS
*=========================================
*
* 1) Click in the menu "File" > "Make a copy..." and make a copy to your Google Drive
* 2) Changes lines 19-32 to be the settings that you want to use
* 3) Click in the menu "Run" > "Run function" > "Install" and authorize the program
*    (For steps to follow in authorization, see this video: https://youtu.be/_5k10maGtek?t=1m22s )
*
*
* **To stop Script from running click in the menu "Edit" > "Current Project's Triggers".  Delete the running trigger.
*
*=========================================
*               SETTINGS
*=========================================
*/

var targetCalendarName = "MAIF";           // The name of the Google Calendar you want to add events to
var sourceCalendarURL = "https://outlook.office365.com/owa/calendar/91bf8de41e824c97ab38870265f40c4a@externe.maif.fr/375b84c13ff840c79405f2940623a4e512588709785996734452/S-1-8-1656950733-2265066604-1668049755-874945561/reachcalendar.ics";            // The ics/ical url that you want to get events from

var howFrequent = 30;                  // What interval (minutes) to run this script on to check for new events
var addEventsToCalendar = true;        // If you turn this to "false", you can check the log (View > Logs) to make sure your events are being read correctly before turning this on
var modifyExistingEvents = true;       // If you turn this to "false", any event in the feed that was modified after being added to the calendar will not update
var removeEventsFromCalendar = true;   // If you turn this to "true", any event in the calendar not found in the feed will be removed.
var addAlerts = true;                  // Whether to add the ics/ical alerts as notifications on the Google Calendar events
var addOrganizerToTitle = false;       // Whether to prefix the event name with the event organiser for further clarity 
var descriptionAsTitles = false;       // Whether to use the ics/ical descriptions as titles (true) or to use the normal titles as titles (false)
var defaultDuration = 60;              // Default duration (in minutes) in case the event is missing an end specification in the ICS/ICAL file
var filterBeforeNow = true;
var maxRecurringEventSeries = 20;

var emailWhenUpdated = false;            // Will email you when an event is updated to your calendar
var emailIfBeforeXMonths = 6            // Will email only for new event occurs in the next x months
var email = "boillodmanuel@gmail.com";                        // OPTIONAL: If "emailWhenAdded" is set to true, you will need to provide your email

var debug = false;
var performAction = true;

// uncomment to debug without any others actions
//debug = true;
//performAction = false;
//emailWhenUpdated = false;

/*
*=========================================
*           ABOUT THE AUTHOR
*=========================================
*
* This program was created by Derek Antrican
*
* If you would like to see other programs Derek has made, you can check out
* his website: derekantrican.com or his github: https://github.com/derekantrican
*
*=========================================
*            BUGS/FEATURES
*=========================================
*
* Please report any issues at https://github.com/derekantrican/GAS-ICS-Sync/issues
*
*=========================================
*           $$ DONATIONS $$
*=========================================
*
* If you would like to donate and help Derek keep making awesome programs,
* you can do that here: https://bulkeditcalendarevents.wordpress.com/donate/
*
*=========================================
*             CONTRIBUTORS
*=========================================
* Andrew Brothers
* Github: https://github.com/agentd00nut/
* Twitter: @abrothers656
*
* Joel Balmer
* Github: https://github.com/JoelBalmer
*
* Blackwind
* Github: https://github.com/blackwind
*
*/


//=====================================================================================================
//!!!!!!!!!!!!!!!! DO NOT EDIT BELOW HERE UNLESS YOU REALLY KNOW WHAT YOU'RE DOING !!!!!!!!!!!!!!!!!!!!
//=====================================================================================================
function Install(){
  ScriptApp.newTrigger("main").timeBased().everyMinutes(howFrequent).create();
}

var vtimezone;
var filterFrom;
var filterTo;
var dateLimitToEmail;
var recurringEventIds;

var updates = []

function main(){
  /*
  var now = new Date()
  var hour = now.getUTCHours();
  var minutes = now.getMinutes();
  
  // Check hour to run
  if (hour < 6 || hour > 20) {
    if (minutes > 15) {
      Logger.log("Do not run because of current hour: " + hour + ":" + minutes + " UTC");
      return;
    }
  }
  Logger.log("Do RUN at this hour: " + hour + ":" + minutes + " UTC");
  */
  
  if (filterBeforeNow) {
    filterFrom = new Date();
    filterFrom.setHours(0, 0, 0, 0);
    filterTo = new Date();
    filterTo.setFullYear(filterTo.getFullYear() + 1);
    filterTo.setHours(23, 59, 59, 999);
  } else {
    filterFrom = new Date(2000,01,01);
    filterTo = new Date(2100,01,01)
  }
  
  dateLimitToEmail = new Date();
  dateLimitToEmail.setMonth(filterTo.getMonth() + emailIfBeforeXMonths);
  

  //Get URL items
  var response = UrlFetchApp.fetch(sourceCalendarURL).getContentText();
  //Logger.log("Response: \n" + response + "\n----------------\n\n")

  //Get target calendar information
  var targetCalendar = CalendarApp.getCalendarsByName(targetCalendarName)[0];


  //------------------------ Error checking ------------------------
  if(response.includes("That calendar does not exist"))
    throw "[ERROR] Incorrect ics/ical URL";

  if(targetCalendar == null){
     throw "[ERROR] Calendar" + targetCalendarName +  " does not exist";
  }

  if (emailWhenUpdated && email == "")
    throw "[ERROR] \"emailWhenUpdated\" is set to true, but no email is defined";
  //----------------------------------------------------------------

  //------------------------ Parse events --------------------------
  var feedEventIds=[];

  //Use ICAL.js to parse the data
  var jcalData = ICAL.parse(response);
  var component = new ICAL.Component(jcalData);

  vtimezone = component.getFirstSubcomponent("vtimezone");
  if (vtimezone != null)
    ICAL.TimezoneService.register(vtimezone);

  //var vtimezones = component.getAllSubcomponents("vtimezone");  
  //Logger.log("TZs" + vtimezones.length)
  //for each (var vtimezone in vtimezones){
    //ICAL.TimezoneService.register(vtimezone);
  //}
  
  //Map the vevents into custom event objects
  var vevents = component.getAllSubcomponents("vevent");    
  var events = []
  recurringEventIds = vevents.map(ToIcalEvent).filter(FilterRecurringEvent).map(GetEventId)
  if (debug) Logger.log("recurringEventIds: " + recurringEventIds.length + " - " + recurringEventIds);
  
  vevents.map(ToIcalEvent).filter(FilterEventAlreadyInRecurringEvent).filter(FilterEndedEvent).forEach(function(icalEvent) {
    var convertedEvents = ConvertToCustomEvents(icalEvent);
    events = events.concat(convertedEvents)
  });
  
  events.forEach(function(event){ 
    feedEventIds.push(event.id); 
  }); //Populate the list of feedEventIds
  //----------------------------------------------------------------
  
  //------------------------ Check results -------------------------
  if (debug) {
    Logger.log("# of events: " + events.length);
    for each (var event in events){
      Logger.log("Title: " + event.title + " - " + (event.isAllDay ? "[all day] - " : "") + formatDate(event.startTime) + " - " + formatDate(event.endTime) + " - " + event.id);
      for each (var reminder in event.reminderTimes)
        Logger.log(" - Reminder: " + reminder + " seconds before");
    }
  }
  //----------------------------------------------------------------

  if(addEventsToCalendar || removeEventsFromCalendar){
    var calendarEvents = targetCalendar.getEvents(filterFrom, filterTo);
    var calendarFids = []
    for (var i = 0; i < calendarEvents.length; i++)
      calendarFids[i] = calendarEvents[i].getTag("FID");
  }

  //------------------------ Add events to calendar ----------------
  if (addEventsToCalendar){
    //Logger.log("Checking " + events.length + " outlook ical events for creation")
    for each (var event in events){
      if (calendarFids.indexOf(event.id) == -1){
        if (performAction) {
          var resultEvent;
          if (event.isAllDay){
            resultEvent = targetCalendar.createAllDayEvent(event.title, 
                                                    event.startTime,
                                                    event.endTime,
                                                    {
                                                      location : event.location, 
                                                      description : event.description
                                                    });
          } else {
            resultEvent = targetCalendar.createEvent(event.title, 
                                                    event.startTime,
                                                    event.endTime,
                                                    {
                                                      location : event.location, 
                                                      description : event.description
                                                    });
          }
          
          resultEvent.setTag("FID", event.id);
          
          for each (var reminder in event.reminderTimes) {
            resultEvent.addPopupReminder(reminder / 60);
          }
        }
        
        Logger.log("    Adding " + event.title + " - " + formatDate(event.startTime) + " - " + formatDate(event.endTime) + " - id=" + event.id);   
        if (event.startTime < dateLimitToEmail.getTime()) {
          updates.push("Adding " + event.title + " - " + formatDate(event.startTime) + " - " + formatDate(event.endTime) + " - id=" + event.id);   
        }
      }
    }
  }
  //----------------------------------------------------------------



  //-------------- Remove Or modify events from calendar -----------  
  for (var i = 0; i < calendarEvents.length; i++){
    //Logger.log("Checking " + calendarEvents.length + " google cal events for removal or modification");
    var c = calendarEvents[i]
    var tagValue = calendarEvents[i].getTag("FID");
    if (debug) Logger.log("GoogleCal " + c.getTitle() + " - " + formatDate(c.getStartTime()) + " - " + formatDate(c.getEndTime()) + " - " + tagValue);
    var feedIndex = feedEventIds.indexOf(tagValue);
    
    if(removeEventsFromCalendar){
      if(feedIndex  == -1 && tagValue != null){
        Logger.log("    Deleting " + calendarEvents[i].getTitle() + " - " + formatDate(calendarEvents[i].getStartTime()) + " - " + formatDate(calendarEvents[i].getEndTime()));        
        updates.push("Deleting " + calendarEvents[i].getTitle() + " - " + formatDate(calendarEvents[i].getStartTime()) + " - " + formatDate(calendarEvents[i].getEndTime()));
        if (performAction) {
          calendarEvents[i].deleteEvent();
        }
      }
    }

    if(modifyExistingEvents){
      if(feedIndex != -1){
        var e = calendarEvents[i];
        var fes = events.filter(sameEvent, tagValue);
        
        if(fes.length > 0){
          var fe = fes[0];
          var updated = false
          var eventUpdates = ""

          if (e.isAllDayEvent() != fe.isAllDay) {
            if (fe.isAllDay) {
              eventUpdates += " becomes all day event: " + formatDate(e.getStartTime()) + " => " + formatDate(fe.startTime)
              if (performAction) e.setAllDayDates(fe.startTime, fe.endTime);
            } else {
              eventUpdates += " stop being an all day event: " + formatDate(e.getStartTime()) + " => " + formatDate(fe.startTime)
              if (performAction) e.setTime(fe.startTime, fe.endTime);
            }
            updated = true;
          } else if (!fe.isAllDay) {
            if(e.getStartTime().getTime() != fe.startTime.getTime() || e.getEndTime().getTime() != fe.endTime.getTime()) {
              if(e.getStartTime().getTime() != fe.startTime.getTime()) {
                eventUpdates += " - startTime: " + formatDate(e.getStartTime()) + " => " + formatDate(fe.startTime)
              }
              if (e.getEndTime().getTime() != fe.endTime.getTime()) {
                eventUpdates += " - endTime: " + formatDate(e.getEndTime()) + " => " + formatDate(fe.endTime)
              }
              if (performAction) e.setTime(fe.startTime, fe.endTime);
              updated = true;
            }
          } else { // fe.isAllDay
            if(e.getAllDayStartDate().getTime() != fe.startTime.getTime() || e.getAllDayEndDate().getTime() != fe.endTime.getTime()) {
              if(e.getAllDayStartDate().getTime() != fe.startTime.getTime()) {
                eventUpdates += " - [allday] startTime: " + formatDate(e.getAllDayStartDate()) + " => " + formatDate(fe.startTime)
              }
              if (e.getAllDayEndDate().getTime() != fe.endTime.getTime()) {
                eventUpdates += " - [allday] endTime: " + formatDate(e.getAllDayEndDate()) + " => " + formatDate(fe.endTime)
              }
              if (performAction) e.setAllDayDates(fe.startTime, fe.endTime);
              updated = true;
            }
          }
          
          if(e.getTitle() != fe.title) {
            eventUpdates += " - title: " + e.getTitle() + " => " + fe.title
            if (performAction) e.setTitle(fe.title);
            updated = true;
          }
          if(e.getLocation() != fe.location) {
            eventUpdates += " - location"
            if (performAction) e.setLocation(fe.location)
            updated = true;
          }
          if(e.getDescription() != fe.description) {
            eventUpdates += " - description: " + e.getDescription() + " => " + fe.description
            if (performAction) e.setDescription(fe.description)
            updated = true;
          }
          if(updated){
            Logger.log("    Updating " + e.getTitle() + " - " + formatDate(e.getStartTime()) + " - " + formatDate(e.getEndTime()) + eventUpdates);        
            updates.push("Updating " + e.getTitle() + " - " + formatDate(e.getStartTime()) + " - " + formatDate(e.getEndTime()) + eventUpdates);              
          }
        }
      }
    }
  }
  
  if (emailWhenUpdated && updates.length > 0) {
    if (debug) Logger.log("\n\nSend email : \n" + "- Title: Calendar " + targetCalendarName + " updated\n" + "- Content: " + updates.join("\n"));   
    GmailApp.sendEmail(email, "Calendar " + targetCalendarName + " updated", updates.join("\n"));
  }
  
  if (updates.length == 0) {
    Logger.log("Done: No updates")
  } else {
    Logger.log("Done: " + updates.length + " updates")
  }
}

function ToIcalEvent(vevent){
  return new ICAL.Event(vevent); 
}

function FilterRecurringEvent(icalEvent){
  return icalEvent.isRecurring()
}

function GetEventId(icalEvent){
  return icalEvent.uid
}

function FilterEventAlreadyInRecurringEvent(icalEvent){
  var keep = icalEvent.isRecurring() || recurringEventIds.indexOf(icalEvent.uid) == -1
  if (!keep && debug) Logger.log("Filtered outlook ical event [FilterEventAlreadyInRecurringEvent]: " + icalEvent.summary + " - " + icalEvent.startDate.toJSDate() + " - " + icalEvent.endDate.toJSDate())
  return keep ;
}

function FilterEndedEvent(icalEvent){
  var keep;
  if (icalEvent.isRecurring()) {
    // Keep if there is an avent after filterFrom date?
    var next = icalEvent.iterator(ICAL.Time.fromJSDate(filterFrom)).next()
    keep = typeof next !== 'undefined'
  } else {  
    var startTime = icalEvent.startDate.toJSDate();
    var endTime = icalEvent.endDate.toJSDate();
    keep = endTime > filterFrom && startTime < filterTo;
  }
  if (!keep && debug) Logger.log("Filtered outlook ical event [FilterEndedEvent]: " + icalEvent.summary + " - " + startTime + " - " + endTime)
  return keep ;
}

function ConvertToCustomEvents(icalEvent){
  var event = new Event();
  event.event = event;
  event.id = icalEvent.uid;
  event.title = icalEvent.summary;
  event.description = (icalEvent.description != null) ? icalEvent.description.trim() : "";
  event.location = (icalEvent.location != null) ? icalEvent.location.trim() : "";;
  if (icalEvent.startDate.isDate && icalEvent.endDate != null && icalEvent.endDate.isDate) {
    event.isAllDay = true;
  }
  event.startTime = icalEvent.startDate.toJSDate();
    
  if (icalEvent.endDate == null)
    event.endTime = new Date(event.startTime.getTime() + defaultDuration * 60 * 1000);
  else{
    event.endTime = icalEvent.endDate.toJSDate();
  }
  
  if (addAlerts){
    var valarms = icalEvent.component.getAllSubcomponents('valarm');
    for each (var valarm in valarms){
      var trigger = valarm.getFirstPropertyValue('trigger').toString();
      event.reminderTimes[event.reminderTimes.length++] = ParseNotificationTime(trigger);
    }
  }
  
  if (icalEvent.isRecurring()) {
    if (debug) {
      Logger.log("Expand recurring event " + event.title + " - " + formatDate(event.startTime) + " - " + formatDate(event.endTime))
    }
    
    // expand events
    var revents = []
    var iterator = icalEvent.iterator(ICAL.Time.fromJSDate(filterFrom));
    var latestStartTime = filterFrom
    var i = 0
    while (i < maxRecurringEventSeries && // stop after max iteration
           latestStartTime.getTime() < filterTo.getTime() && // stop if recurring event startTime is > filterTo
           (next = iterator.next())) { // stop if no more event ;)
      var revent = cloneForRecurringEvent(event)
      var detail = icalEvent.getOccurrenceDetails(next);
      revent.startTime = detail.startDate.toJSDate();
      revent.endTime = detail.endDate.toJSDate();

      // fix hour due to a bug in ical library
      var offsetInMin = getTimezoneOffset(revent.startTime, 'Europe/Paris') - getTimezoneOffset(event.startTime, 'Europe/Paris') // Daylight Saving Time
      //Logger.log("Expanded event " + i + ": " + offsetInMin+ " " + getTimezoneOffset(event.startTime, 'Europe/Paris') + " - " + getTimezoneOffset(revent.startTime, 'Europe/Paris'))
      revent.startTime.setUTCHours(event.startTime.getUTCHours(), event.startTime.getMinutes() + offsetInMin, event.startTime.getSeconds(), event.startTime.getMilliseconds());
      revent.endTime.setUTCHours(event.endTime.getUTCHours(), event.endTime.getMinutes() + offsetInMin, event.endTime.getSeconds(), event.endTime.getMilliseconds());
      // end fix bug
      
      revent.id = event.id + "-" + formatDateOnly(revent.startTime)
      if (debug) {
        Logger.log("Expanded event " + i + ": " + formatDate(revent.startTime) + " - " + formatDate(revent.endTime))
      }
      if (revent.startTime.getTime() <= filterTo.getTime()) { // add only if startTime < filterTo
        revents.push(revent);
      }
      latestStartTime = revent.startTime;
      i++;
    }
    if (debug) {
      Logger.log("Expand recurring event " + event.title + " into " + revents.length + " events")
    }
    return revents;
  } else {
    return [event];
  }
}

function cloneForRecurringEvent(event) {
    var e = new Event();
    e.title = event.title;
    e.description = event.description;
    e.location = event.location;
    e.event = event.event;
    // do not do a full clone of reminderTimes
    e.reminderTimes = event.reminderTimes;
    // do not clone startTime, endTime, id
    return e;
}
function ParseOrganizerName(veventString){
  /*A regex match is necessary here because ICAL.js doesn't let us directly
  * get the "CN" part of an ORGANIZER property. With something like
  * ORGANIZER;CN="Sally Example":mailto:sally@example.com
  * VEVENT.getFirstPropertyValue('organizer') returns "mailto:sally@example.com".
  * Therefore we have to use a regex match on the VEVENT string instead
  */

  var nameMatch = RegExp("ORGANIZER(?:;|:)CN=(.*?):", "g").exec(veventString);
  if (nameMatch.length > 1)
    return nameMatch[1];
  else
    return null;
}

function ParseNotificationTime(notificationString){
  //https://www.kanzaki.com/docs/ical/duration-t.html
  var reminderTime = 0;

  //We will assume all notifications are BEFORE the event
  if (notificationString[0] == "+" || notificationString[0] == "-")
    notificationString = notificationString.substr(1);

  notificationString = notificationString.substr(1); //Remove "P" character

  var secondMatch = RegExp("\\d+S", "g").exec(notificationString);
  var minuteMatch = RegExp("\\d+M", "g").exec(notificationString);
  var hourMatch = RegExp("\\d+H", "g").exec(notificationString);
  var dayMatch = RegExp("\\d+D", "g").exec(notificationString);
  var weekMatch = RegExp("\\d+W", "g").exec(notificationString);

  if (weekMatch != null){
    reminderTime += parseInt(weekMatch[0].slice(0, -1)) & 7 * 24 * 60 * 60; //Remove the "W" off the end

    return reminderTime; //Return the notification time in seconds
  }
  else{
    if (secondMatch != null)
      reminderTime += parseInt(secondMatch[0].slice(0, -1)); //Remove the "S" off the end

    if (minuteMatch != null)
      reminderTime += parseInt(minuteMatch[0].slice(0, -1)) * 60; //Remove the "M" off the end

    if (hourMatch != null)
      reminderTime += parseInt(hourMatch[0].slice(0, -1)) * 60 * 60; //Remove the "H" off the end

    if (dayMatch != null)
      reminderTime += parseInt(dayMatch[0].slice(0, -1)) * 24 * 60 * 60; //Remove the "D" off the end

    return reminderTime; //Return the notification time in seconds
  }
}


function sameEvent(x){
  return x.id == this;
}

function formatDate(date) {
  if (date == null) return "(null)";
  return Utilities.formatDate(date, "Europe/Paris", "yyyy-MM-dd HH:mm:ss");
}

function formatDateOnly(date) {
  if (date == null) return "(null)";
  return Utilities.formatDate(date, "Europe/Paris", "yyyy-MM-dd");
}

function cloneDateInTz(d, tz) {
  var ls = Utilities.formatDate(d, tz, "yyyy/MM/dd HH:mm:ss");
  var a = ls.split(/[\/\s:]/);
  //Logger.log("getTimezoneOffset:" + tz + ' = ls = ' + ls + ' / a = ' + a)
  a[1]--;
  var t1 = Date.UTC.apply(null, a);
  var t2 = new Date(d).setMilliseconds(0);
}

/** 
  get the timezoneoffset of a date for a given timezone 
  Default method getTimeZoneOffset() return for the current time which is in USA for google script and which have different daylight saving time
*/
function getTimezoneOffset(d, tz) {
  var ls = Utilities.formatDate(d, tz, "yyyy/MM/dd HH:mm:ss");
  var a = ls.split(/[\/\s:]/);
  //Logger.log("getTimezoneOffset:" + tz + ' = ls = ' + ls + ' / a = ' + a)
  a[1]--;
  var t1 = Date.UTC.apply(null, a);
  var t2 = new Date(d).setMilliseconds(0);
  return (t2 - t1) / 60 / 1000;
}

function clearCalendar(){
  var targetCalendar = CalendarApp.getCalendarsByName(targetCalendarName)[0];
  var calendarEvents = targetCalendar.getEvents(new Date(2000,01,01), new Date(2100,01,01))
  for (var i = 0; i < calendarEvents.length; i++) {
    Logger.log("Deleting " + calendarEvents[i].getTitle() + " - " + formatDate(calendarEvents[i].getStartTime()) + " - " + formatDate(calendarEvents[i].getEndTime()));        
    calendarEvents[i].deleteEvent();
  }
}
