var targetCalendarName = 'Autogen birthdays';//hand-made(manually pre-configured) special calendar
var daysCount = 5;

var contactsCalendar;
var targetCalendar;
var today;
var startDate;
var endDate;
var eventList;

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
  contactsCalendar = CalendarApp.getCalendarById('addressbook#contacts@group.v.calendar.google.com');//https://developers.google.com/apps-script/reference/calendar/calendar-app
  var calendars = CalendarApp.getCalendarsByName(targetCalendarName);
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
  startDate = new Date(today.getTime());
  endDate = new Date(today.getTime() + daysCount * (1000 * 60 * 60 * 24));
  const dateTimeFormat = 'MMMM dd, yyyy HH:mm:ss Z';
  const targetTimeZone = CalendarApp.getTimeZone();
  Logger.log('INFO: start date: ' + Utilities.formatDate(startDate,  targetTimeZone, dateTimeFormat));
  Logger.log('INFO: end date' + Utilities.formatDate(endDate, targetTimeZone, dateTimeFormat));
  eventList = contactsCalendar.getEvents(startDate, endDate);
  Logger.log('INFO: total contacts found' + eventList.length);
})();

function birthdayAgeToCalendar() {
    for (var i in eventList) {
        Logger.log('birthdayAgeToCalendar. дни рождения. Найдено: ' + eventList[i].getTitle());
        var name = eventList[i].getTitle().split(" – день рождения")[0];
        var contacts = ContactsApp.getContactsByName(name);
        Logger.log('birthdayAgeToCalendar. дни рождения. Name: ' + name);

        for (var c in contacts) {
            var bday = contacts[c].getDates(ContactsApp.Field.BIRTHDAY);
            var bdayMonthName, bdayYear, bdayDate;
            try {
                bdayMonthName = bday[0].getMonth();
                bdayYear = bday[0].getYear();
                bdayDate = new Date(bdayMonthName + ' ' + bday[0].getDay() + ', ' + bdayYear);
                Logger.log('birthdayAgeToCalendar. bdayDate: ' + bdayDate);
            } catch (error) {
              Logger.log('ERROR: ' + error);
              continue;
            }

            var years = parseInt(today.getFullYear()) - bdayYear;
            try {
              var event = targetCalendar.createAllDayEvent(name + " – день рождения, " + agetostr(years),
                  new Date(bdayMonthName + ' ' + bday[0].getDay() + ', ' + new Date().getFullYear()));
              Logger.log('INFO: EventID is ' + event.getId());
            } catch (error) {
              Logger.log('ERROR: ' + error);
            }
        }
    }
}

function anniversaryAgeToCalendar() { //юбилеи
    for (var i in eventList) {
        Logger.log('anniversaryAgeToCalendar. Юбилеи. Найдено: ' + eventList[i].getTitle());
        var name = eventList[i].getTitle().split("Юбилей у пользователя ")[1];
        var contacts = ContactsApp.getContactsByName(name);
        Logger.log('anniversaryAgeToCalendar. юбилеи. Name: ' + name);

        for (var c in contacts) {
            var bday = contacts[c].getDates(ContactsApp.Field.ANNIVERSARY); //существующие типы данных https://developers.google.com/apps-script/reference/contacts/field
            var bdayMonthName, bdayYear, bdayDate;
            try {
                bdayMonthName = bday[0].getMonth();
                bdayYear = bday[0].getYear();
                bdayDate = new Date(bdayMonthName + ' ' + bday[0].getDay() + ', ' + bdayYear);
                Logger.log('anniversaryAgeToCalendar. bdayDate: ' + bdayDate);
            } catch (error) {}

            var years = parseInt(new Date().getFullYear()) - bdayYear;
            try {
                targetCalendar.createAllDayEvent("Юбилей у пользователя " + name + ", " + years + " лет (года)",
                    new Date(bdayMonthName + ' ' + bday[0].getDay() + ', ' + new Date().getFullYear()));
                Logger.log("Создано: " + "Юбилей у пользователя " + name + ", " + years + " лет (года)");
            } catch (error) {}
        }
    }
}

function TriggersCreateTimeDriven() { //автоматически создаем новые триггеры для запуска
    // Deletes all triggers in the current project.
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        ScriptApp.deleteTrigger(triggers[i]);
    }
    
    ScriptApp.newTrigger("birthdayAgeToCalendar") //дни рождения
        .timeBased()
        .onMonthDay(1) //день месяца
        .atHour(1)
        .create();
  /*
    ScriptApp.newTrigger("anniversaryAgeToCalendar") //юбилеи
        .timeBased()
        .onMonthDay(1)
        .atHour(2)
        .create();
        */
}
