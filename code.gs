// Author: John Rush (https://github.com/JohnRush)
// License: MIT
// Link: https://github.com/JohnRush/google-calendar-copy
//
// --Overview--
// This script imports events from one calander (external) into another (primary).
// Update the "settings" found below to make any changes you need.
// This does not copy all-day events, but it does work with recurring events.create.
// It will also update events that have changed and delete events that no longer exist.
//
// It uses to Google Calendar GSuite Service found here:
//   https://developers.google.com/apps-script/reference/calendar/
//
// --Installation and Setup --
// Here is a good description of how to get a script like this working.
// I think this script has some good improvements over the one at this link, but the process is the same.
//   https://medium.com/@willroman/auto-block-time-on-your-work-google-calendar-for-your-personal-events-2a752ae91dab
//
// --Special Thanks--
// The following people/posts inspired this effort:
//   Debsankha at http://blog.debsankha.net/2011/01/merging-two-google-calendars.html
//   Will Roman at https://medium.com/@willroman/auto-block-time-on-your-work-google-calendar-for-your-personal-events-2a752ae91dab
//
function syncEmail() {
  const settings = {
    externalCalendarId: 'your_google_calendar_id_goes_here',
    daysToLookAhead: 60, // how many days in advance to monitor and block off time
    markPrivate: false, // Mark imported events as private?
    appointmentTitle: 'Booked', // The name of events once they are imported (if not private)
    includeWeekends: false, // Include events that are on a weekend?
    keepReminders: false, // Do you want reminders to be added on your new imported events?
    color: 0, // 0 gives the default color, or a pick number from 1 to 11 that corresponds to a color (https://developers.google.com/apps-script/reference/calendar/event-color)
    removeAllImportedEvents: false, // Delete all existing imported events (in case there is a problem)
    slowDownPlease: 200 // If you get "You have been creating or deleting too many calendars or calendar events in a short time." errors, you can increase this number of ms to wait
  };
  
  const ID_TAG = 'externalId';
  const today = new Date();
  var enddate = new Date();
  enddate.setDate(today.getDate() + settings.daysToLookAhead);
 
  // Get all of the events from the primary (target) calendar
  const primaryCalendar = CalendarApp.getDefaultCalendar();
  const primaryEvents = primaryCalendar.getEvents(today, enddate)
    .filter(function(x) { return x.getTag(ID_TAG) }); // only process events with an external id
  Logger.log('Found ' + primaryEvents.length + ' imported event' + (primaryEvents.length === 1 ? '' : 's' ) + '.');
  
  // If requested, delete all of the external events and exit
  if (settings.removeAllImportedEvents) {
    primaryEvents.forEach(function(event) {
      event.deleteEvent();
    });

    Logger.log('Deleted ' + primaryEvents.length + ' imported event' + (primaryEvents.length === 1 ? '' : 's' ) + '.');
    return;
  }

  // Get all of the events from the external (source) calendar
  const externalEvents = CalendarApp.getCalendarById(settings.externalCalendarId).getEvents(today, enddate)
  // Don't include all day events
  .filter(function(x) { return !x.isAllDayEvent(); })
  // Don't include weekend events unless the settings say otherwise
  .filter(function(x) { return settings.includeWeekends || (x.getStartTime().getDay() >= 1 && x.getStartTime().getDay() <= 5); });
    Logger.log('Found ' + externalEvents.length + ' external event' + (primaryEvents.length === 1 ? '' : 's' ) + ' to process.');

  // Go through each external event and add or update it, as needed
  externalEvents.forEach(function(externalEvent) {
    var externalId = externalEvent.getId();
    var existingEvents = primaryEvents.filter(function(x) {
      return x.getTag(ID_TAG) === externalId
    }); // Find all matches for this event id (repeating events have the same id)

    var desiredTitle = settings.markPrivate ? externalEvent.getTitle() : settings.appointmentTitle;
    var desiredVisbility = settings.markPrivate ? CalendarApp.Visibility.PRIVATE : CalendarApp.Visibility.PUBLIC;
    var desiredDescription = settings.markPrivate ? externalEvent.getDescription() : '';
    var desiredColor = settings.color;
    var desiredStartTime = externalEvent.getStartTime();
    var desiredEndTime = externalEvent.getEndTime();

    if (existingEvents.length > 1) {
      // We can reuse (update) a repeating sequence, but only if we find an exact match
      // for the time slot. If the time changes on the external event we will just create
      // new events and delete the old events, instead trying to match them up.
      var perfectMatch = existingEvents.filter(function(x) {
        return x.getStartTime().getTime() === desiredStartTime.getTime()
        && x.getEndTime().getTime() === desiredEndTime.getTime()
      });

      if (perfectMatch.length === 1) {
        existingEvents = perfectMatch;
      } else {
        existingEvents = [];
      }
    }

    var event = null;
    var isNew = existingEvents.length === 0;

    if (isNew) {
      // Create a new event
      var dayOfWeek = externalEvent.getStartTime().getDay();
      event = primaryCalendar.createEvent(desiredTitle, desiredStartTime, desiredEndTime);
      if (settings.slowDownPlease) Utilities.sleep(settings.slowDownPlease);
      event.setTag(ID_TAG, externalId);

      // New events default to having reminders. Since we have reminders in the original
      // calendar, we probably don't want them here as well
      if(!settings.keepReminders) {
        event.removeAllReminders();
      }
      
      Logger.log('ADD ' + externalEvent.getTitle() + ' at ' + externalEvent.getStartTime());
    } else {
      // Remove this from the list of "known" events, so we know not to delete it later
      event = existingEvents[0];
      var i = primaryEvents.indexOf(event);
      primaryEvents.splice(i, 1);
    }

    var changes = [];

    var setTime = false;
    if (event.getStartTime().getTime() !== desiredStartTime.getTime()) {
      changes.push('Changed start time to ' + desiredStartTime);
      setTime = true;
    }

    if (event.getEndTime().getTime() !== desiredEndTime.getTime()) {
      changes.push('Changed end time to ' + desiredStartTime);
      setTime = true;
    }

    if (setTime) {
      event.setTime(desiredStartTime, desiredEndTime);
    }

    if (event.getVisibility() !== desiredVisbility) {
      event.setVisibility(desiredVisbility);
      changes.push('Changed visiblity to ' + (settings.markPrivate ? 'private' : 'public'));
    }

    if (event.getTitle() !== desiredTitle) {
      event.setTitle(desiredTitle);
      changes.push('Changed title to "' + desiredTitle + '"');
    }

    if (event.getDescription() !== desiredDescription) {
      event.setDescription(desiredDescription);
      changes.push('Changed description to "' + desiredDescription + '"');
    }

    if (event.getColor() != desiredColor) {
      event.setColor(desiredColor);
      changes.push('Changed color to "' + desiredColor + '"');
    }

    if (!isNew) {
      if (changes.length) {
        Logger.log('MOD ' + externalEvent.getTitle() + ' at ' + externalEvent.getStartTime());
        changes.forEach(function(change) {
          Logger.log('    ' + change);
        });
      } else {
        Logger.log('NOP ' + externalEvent.getTitle() + ' at ' + externalEvent.getStartTime());
      }
    }
  });

  // Anything left in primaryEvents will be an event that is no longer in the external calendar
  // This can happen when an item is moved to a different time or is just deleted
  primaryEvents.forEach(function(event) {
      event.deleteEvent();
      if (settings.slowDownPlease) Utilities.sleep(settings.slowDownPlease);
      Logger.log('DEL ' + event.getTitle() + ' at ' + event.getStartTime());
  });
}
