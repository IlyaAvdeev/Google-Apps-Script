var targetCalendarName = 'Autogen birthdays';//hand-made(manually pre-configured) special calendar

var targetCalendar;
var today;
var contactList;
var thisYear;
var targetCalendarEvents;
var journal = '';

function agetostr(age) {//inspired by https://gist.github.com/paulvales/113b08bdf3d4fc3beb2a5e0045d9729d
	var txt;
	count = age % 100;
	if (count >= 5 && count <= 20) {
		txt = 'лет';
	} else {
		count = count % 10;
		if (count == 1) {
			txt = 'год';
		} else if (count >= 2 && count <= 4) {
			txt = 'года';
		} else {
			txt = 'лет';
		}
	}
	return age+" "+txt;
}

(function() {
  //var contactsCalendar;
  //https://developers.google.com/apps-script/reference/calendar/calendar-app
  //contactsCalendar = CalendarApp.getCalendarById('addressbook#contacts@group.v.calendar.google.com');
  var calendars = CalendarApp.getCalendarsByName(targetCalendarName);//https://developers.google.com/apps-script/reference/calendar/calendar-app#getCalendarsByName(String)
  if (calendars.length == 0) {
    Logger.log('ERROR: No target calendar found.');
    return;
  }
  if (calendars.length > 1) {
    Logger.log('WARNING: Ambigious target calendar!');
    return;
  } else {
    targetCalendar = calendars[0];
  }
  
  today = new Date();
  thisYear = today.getYear();
  const dateTimeFormat = 'MMMM dd, yyyy HH:mm:ss Z';
  const targetTimeZone = CalendarApp.getTimeZone();
  Logger.log('INFO: today is ' + Utilities.formatDate(today,  targetTimeZone, dateTimeFormat));
  contactList = ContactsApp.getContacts();
  if (contactList == null || contactList.length == 0) {
    Logger.log('ERROR: No contacts in address book.');
    return;
  }
  Logger.log('INFO: total contacts ' + contactList.length);
  targetCalendarEvents = targetCalendar.getEvents(today, new Date(thisYear + 1, 12, 31));
  Logger.log('INFO: total events in calendar ' + targetCalendarEvents.length);
})();

function month2I(month) {
  if (month == CalendarApp.Month.JANUARY) return 1;
  else if (month == CalendarApp.Month.FEBRUARY) return 2;
  else if (month == CalendarApp.Month.MARCH) return 3;
  else if (month == CalendarApp.Month.APRIL) return 4;
  else if (month == CalendarApp.Month.MAY) return 5;
  else if (month == CalendarApp.Month.JUNE) return 6;
  else if (month == CalendarApp.Month.JULY) return 7;
  else if (month == CalendarApp.Month.AUGUST) return 8;
  else if (month == CalendarApp.Month.SEPTEMBER) return 9;
  else if (month == CalendarApp.Month.OCTOBER) return 10;
  else if (month == CalendarApp.Month.NOVEMBER) return 11;
  else if (month == CalendarApp.Month.DECEMBER) return 12;
  else return -1;
}

function contactAge2Calendar() {
  var today_month = today.getMonth();//0-11
  var today_day = today.getDate();//1-31 
  for (var i = 0; i < contactList.length; i++) {//https://developers.google.com/apps-script/reference/contacts/contact
    var contact = contactList[i];
    var contactId = contact.getId();
    var birthdayDates = contact.getDates(ContactsApp.Field.BIRTHDAY);
    var name = contact.getFullName();
    Logger.log('INFO: name ' + name.toString());
    Logger.log('INFO: id ' + contactId);
    if (birthdayDates.length == 0) {
      continue;
    }

    var birthday = birthdayDates[0];
    var month, year = 0, day;
    try {
      year = birthday.getYear();
      month = month2I(birthday.getMonth());//1-12
      day = birthday.getDay();//1-31
    } 
    catch (error) {
      Logger.log('ERROR: ' + error);
      continue;
    }
    Logger.log('INFO: birthday is ' + year + '/' + month + '/' + day);
    
    if (month == -1) {
      continue;
    }
    
    if (year == undefined || year == null || year == 0) {
      Logger.log('WARN: information about year of birthday is not available.');
      continue;
    }
    
    if ((month - 1) < today_month) {
      // too late, not interested in this.      
      continue;
    } else {
      if ((month - 1) == today_month) {
        if (day < today_day) {
          // too late, not interested in this.
          continue;
        }
      }
    }
          
    var years = thisYear - year;
    try {//https://developers.google.com/apps-script/reference/calendar/calendar#createAllDayEvent(String,Date)
      var title = name + " – день рождения, " + agetostr(years);
      var event = getEventByTag('CID', contactId);
      if (event == null) {
        var event2 = targetCalendar.createAllDayEvent(title,new Date(thisYear, month - 1, day));
        event2.setTag('CID', contactId);
        Logger.log('INFO: EventID is ' + event2.getId() + ' for contact ' + name);
        journal += title + ' on ' + event2.getAllDayEndDate().toString() + '\n';
      } else {
        Logger.log('INFO: Event with CID found.');
        event.setTitle(title);
      }
    } catch (error) {
      Logger.log('ERROR: ' + error);
    }
  }

  if (journal.length > 0) {
    //send email to self https://developers.google.com/apps-script/reference/base/session?hl=ru
    var gmailDraft = GmailApp.createDraft(Session.getActiveUser().getEmail(), 'Ages2Calendar Report ' + today.toString(), journal);//https://developers.google.com/apps-script/reference/gmail/gmail-app
    gmailDraft.send();
  }
}

function removeAllCalendarEvents() {
  for (var i = 0; i < targetCalendarEvents.length; i++) {
    var event = targetCalendarEvents[i];
    event.deleteEvent();
  }
}

function getEventByTag(key, value) {
  for (var i = 0; i < targetCalendarEvents.length; i++) {
    var event = targetCalendarEvents[i];
    if(typeof event.getTag === 'function') {
      if (event.getTag(key) === value) {
        return event;
      }
    }
  }
  return null;
}

function scheduleNextAppLaunch() {
    // Deletes all triggers in the current project.
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        ScriptApp.deleteTrigger(triggers[i]);
    }
    
    ScriptApp.newTrigger("contactAge2Calendar")//https://developers.google.com/apps-script/reference/script/script-app#newTrigger(String)
    .timeBased()//https://developers.google.com/apps-script/reference/script/trigger-builder#timeBased()
    .onWeekDay(CalendarApp.Weekday.SUNDAY)
    .atHour(1)
    .inTimezone("Europe/Moscow")//http://joda-time.sourceforge.net/timezones.html
    .create();
}
