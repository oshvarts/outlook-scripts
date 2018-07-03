// This tool is written in JSCRIPT.

// Parts Copyright 2006 - Ryan Watkins (ryan@ryanwatkins.net)
//
// 
// much code converted to javascript from outlook2ical v1.03
// by Norman L. Jones, Provo Utah (njones61@gmail.com)
// additional fixes by
// Andrew Johnson, Alastair Rankine, Dane Walther, Zan Hecht, Markus Untera

// // version 2.0 (2018)

// Lots of changes by Oleg Shvartsman.

// configuration --------------------------------------------------------

//var categories = new Array("Private","Holiday"); // calendar categories not/only to export
var categories = new Array("Holiday","DoNotExport"); // calendar categories not/only to export
var exportMode = "not";                          // "all"  - to export ALL entries (regardless of categories above)
// "not"  - to export all BUT categories listed; 
// "only" - to export ONLY those categories listed; 

var includeHistory = 5;          // how many days back to include old events
var icsFilename = "C:\\Public\\calendar.ics";  // where to store the file


var linebreak = "\r\n";

// Outlook Redemption works around limitations imposed by the Outlook
// Security Patch and Service Pack 2 of MS Office 98/2000 and Office
// 2002 and 2003 (which include Security Patch).
// 
// In our case, the 'body' object for the calendar items - the
// 'description' in the ics files is protected by default and Outlook
// will prompt you to allow access when the script is run.  If you
// install Redemption, available at
//
//   http://www.dimastr.com/redemption/
//
// then you can avoid this issue and automate running this script.
//
// If you cant install Redemption and want to include the
// body/description for the calendar items, despite the promoting from
// Outlook, set includeBody to true.  If this script finds Redemption
// is installed, it will set includeBody to true. 

// Note -- September 2016, Redemption seems to have no effect.  Outlook still prompts for 
// confirmation.

var includeBody = true;

// if you wish to include reminders with your ical items, set this to 'true'

var includeAlarm = true;

// ----------------------------------------------------------------------

// Outlook constants - from http://www.winscripter.com/
// OlDaysOfWeek
var olSunday    =  1;
var olMonday    =  2;
var olTuesday   =  4;
var olWednesday =  8;
var olThursday  = 16;
var olFriday    = 32;
var olSaturday  = 64;

// OlDefaultFolders Constants 
var olFolderCalendar = 9;

// OlRecurrenceType Constants 
var olRecursDaily    = 0;
var olRecursWeekly   = 1;
var olRecursMonthly  = 2;
var olRecursMonthNth = 3;
var olRecursYearly   = 5;
var olRecursYearNth  = 6;

// OlSensitivity Constants
var olNormal       = 0;
var olPersonal     = 1;
var olPrivate      = 2;
var olConfidential = 3;

var ics = "BEGIN:VCALENDAR" + linebreak +
	"VERSION:2.0" + linebreak +
	"PRODID:-//Google Inc//Google Calendar 70.9054//EN" + linebreak +
	"METHOD:PUBLISH" + linebreak +
	"CALSCALE:GREGORIAN" + linebreak + linebreak +
	"BEGIN:VTIMEZONE" + 
	"TZID:America/New_York" + linebreak +
	"LAST-MODIFIED:20050809T050000Z" + linebreak +
	"BEGIN:STANDARD" + linebreak +
	"DTSTART:20071104T020000" + linebreak +
	"TZOFFSETFROM:-0400" + linebreak +
	"TZOFFSETTO:-0500" + linebreak +
	"TZNAME:EST" + linebreak +
	"END:STANDARD" + linebreak +
	"BEGIN:DAYLIGHT" + linebreak +
	"DTSTART:20070311T020000" + linebreak +
	"TZOFFSETFROM:-0500" + linebreak +
	"TZOFFSETTO:-0400" + linebreak +
	"TZNAME:EDT" + linebreak +
	"END:DAYLIGHT" + linebreak +
	"END:VTIMEZONE" + linebreak + linebreak;

var ol = new ActiveXObject("outlook.application");
var calendar = ol.getnamespace("mapi").getdefaultfolder(olFolderCalendar).items;

var today = new Date();
var total = calendar.Count;  
var exportItem = true;
//alert ("Total " + total + " items."); 
// if (total > 275) {total = 270);

for (var i=1; i<=total; i++) {
	
    try {
        var rItem = calendar(i);
		var item = new ActiveXObject("Redemption.SafeMeetingItem");
        item.item = rItem;
        includeBody = true;
		} catch(e) {
		var item = calendar(i);
	}
	
    try {
        var m = item.categories;
		} catch(e) {
		//     ignore
	}
	
	
	
    if ((exportMode == "not") || (exportMode == "only")) {
        
        exportItem = (exportMode=="not") ? true : false; // setup default
        
        for (var j=0; j < categories.length; j++) {
            if ((null != item.categories) && (item.categories.indexOf(categories[j]) != -1)) {  // category found on item
                exportItem = (exportMode == "only") ? true : false;
			}
            continue; // no need to go through the rest of the array... 
		}
	} 
	
    if (exportItem) {
        if (!item.isrecurring) {
            var endDate = new Date(item.end);
            if (Math.round(((today - endDate) / (86400000))) > includeHistory) { continue; }
		}
        if (item.isrecurring) {
            var endDate = new Date(item.getrecurrencepattern.patternenddate); // abracadabra
            if (Math.round(((today - endDate) / (86400000))) > includeHistory) { continue; }
		}
		ics += createEvent(item, false);
	}
	item.close(false);
	item = null; 
}

ics += "END:VCALENDAR" + linebreak + linebreak;

var fso = new ActiveXObject("Scripting.FileSystemObject");
var icsFH = fso.CreateTextFile(icsFilename, true, false);
icsFH.Write(ics);
icsFH.Close();

WScript.Quit();

///////  END FILE.


function createEvent(item, notRecurring) {
	
    var event = "BEGIN:VEVENT" + linebreak;
	
    if (item.alldayevent == true) {
	    event += "DTSTAMP;VALUE=DATE:" + formatDate(item.start) + "T000000" + linebreak;
        event += "DTSTART;VALUE=DATE:" + formatDate(item.start) + linebreak;
        if (item.isrecurring == false) {
            event += "DTEND;VALUE=DATE:" + formatDate(item.end) + linebreak;
		}
	} 
	else {
            //event += "// " + item.StartInStartTimeZone + "\n";
        event += "DTSTART:" + formatDateTime(item.start) + linebreak;
		event += "DTSTAMP:" + formatDateTime(item.start) + linebreak;
        event += "DTEND:" + formatDateTime(item.end) + linebreak;
	}
	
    if (item.isrecurring == true && notRecurring == false) {
        event += createReoccuringEvent(item);
	}
	
	try {
    event += "LOCATION:" + cleanLineEndings(item.location) + optionalOrganizer (item.Organizer) + linebreak;
	} catch(e) {
		WScript.Quit();
	}
	
    event += "SUMMARY:" + cleanLineEndings(item.subject) + linebreak;
	event += "UID:" + item.entryid;
	
	if (notRecurring == true)
	{  event += "-" + randInt(0, 100000); }
	event += linebreak;
    
    if (item.categories.length < 1) {
        event += "CATEGORIES:(none)" + linebreak;
		} else {
        event += "CATEGORIES:" + item.categories + linebreak;
	}
	
    if (item.sensitivity == olNormal) {  
        event += "CLASS:PUBLIC" + linebreak;
		} else if (item.sensitivity == olPersonal) {
        event += "CLASS:CONFIDENTIAL" + linebreak;
		} else {
        event += "CLASS:PRIVATE" + linebreak;
	}
	
    if (includeBody) {
        event += "DESCRIPTION:" + getOptionalBody(item) + linebreak;
	}
    if (includeAlarm) {
		if ((item.reminderminutesbeforestart > 0) && (item.reminderset)){
	        event += "BEGIN:VALARM" + linebreak;
	        event += "TRIGGER:-PT" + item.reminderminutesbeforestart + "M" + linebreak;
	        event += "ACTION:DISPLAY" + linebreak + "DESCRIPTION:Reminder" + linebreak + "END:VALARM" + linebreak;
		}
		else {
	        // event += "BEGIN:VALARM" + linebreak;
			// event += "TRIGGER:+PT60M" + linebreak;
	        // event += "ACTION:DISPLAY" + linebreak + "DESCRIPTION:Reminder" + linebreak + "END:VALARM" + linebreak;
		}
		
	}
    event += "END:VEVENT" + linebreak + linebreak;
	
	
	
	if (!notRecurring && item.IsRecurring)
	{
		var pattern = item.getrecurrencepattern;
		for (var i = 1; i <= pattern.Exceptions.Count; i++)
		
		{
			except = pattern.exceptions(i);
			if (!except.Deleted)
			{
				
				try {
					event += createEvent(except.AppointmentItem, true);  
					} catch(e) {
					//alert(item.subject)
					//alert (except);
					//alert(e.description);
				}
				
			}
		}
	}
	
	
	
    return event;
}

function createReoccuringEvent(item) {
	
    var recurEvent = "RRULE:";
	
    var pattern = item.getrecurrencepattern;
    var patternType = pattern.recurrencetype;
	
    if (patternType == olRecursDaily) {
		
        recurEvent += "FREQ=DAILY";
        if (pattern.noenddate != true) {
            recurEvent += ";UNTIL=" + formatDateTime(pattern.patternenddate);
            // The end date/time is marked as 12:00am on the last day.
            // When this is parsed by php-ical, the last day of the
            // sequence is missed. The MS Outlook code has the same
            // bug/issue.  To fix this, change the end time from 12:00 am
            // to 11:59:59 pm.
            recurEvent = recurEvent.replace(/T000000/g,"T235959Z");
		}
        recurEvent += getInterval(pattern.interval);
		
		} else if (patternType == olRecursMonthly) {
		
        recurEvent += "FREQ=MONTHLY";
        if (pattern.noenddate != true) {
            recurEvent += ";UNTIL=" + formatDateTime(pattern.patternenddate);
		}
        recurEvent += getInterval(pattern.interval);
        recurEvent += ";BYMONTHDAY=" + pattern.dayofmonth;
		
		} else if (patternType == olRecursMonthNth) {
		
        recurEvent += "FREQ=MONTHLY";
        if (pattern.noenddate != true) {
            recurEvent += ";UNTIL=" + formatDateTime(pattern.patternenddate);
		}
		
        recurEvent += getInterval(pattern.interval);
        // php-icalendar has a bug for monthly recurring events.  If
        // it is the last day of the month, you can't use the
        // BYDAY=-1SU option, unless you also do the BYMONTH option
        // (which only is useful for yearly events).  However, the
        // BYWEEK option seems to work for the last week of the month
        // (but not for the first week of the month).  Anyway, this
        // exeception seems to work.
        if (pattern.instance == 5) {
            recurEvent += ";BYWEEK=-1;BYDAY=" & daysOfWeek("", pattern);
            } else {
            recurEvent += ";BYDAY=" + daysOfWeek(weekNum(pattern.instance), pattern);
		}
		
		} else if (patternType == olRecursWeekly) {
		
        recurEvent += "FREQ=WEEKLY";
        if (pattern.noenddate != true) {
            recurEvent += ";UNTIL=" + formatDateTime(pattern.patternenddate);
		}
        recurEvent += getInterval(pattern.interval);
        recurEvent += ";BYDAY=" + daysOfWeek("", pattern);
		
		} else if (patternType == olRecursYearly) {
		
        recurEvent += "FREQ=YEARLY";
        if (pattern.noenddate != true) {
            recurEvent += ";UNTIL=" + formatDateTime(pattern.patternenddate);
		}
        recurEvent += ";INTERVAL=1";
		//        recurEvent += ";BYDAY=" + daysOfWeek("", pattern);
		
		} else if (patternType == olRecursYearNth) {
		
        recurEvent += "FREQ=YEARLY";
        if (pattern.noenddate != true) {
            recurEvent += ";UNTIL=" + formatDateTime(pattern.patternenddate);
		}
        recurEvent += ";BYMONTH=" + monthNum(pattern.monthofyear);
        recurEvent += ";BYDAY=" + daysOfWeek(weekNum(pattern.instance), pattern);
		
	}
	
    recurEvent += "" + linebreak;
	
	if (pattern.Exceptions.Count > 0)
	{
		recurEvent += "EXDATE:";
		//NOTE: I need to think of a better way to do this, but this works for now.
		var firstExcept = true;
		for (var i = 1; i <= pattern.Exceptions.Count; i++)
		
		{
			except = pattern.exceptions(i);
			if (!firstExcept)
			{
				recurEvent += ",";
			}
			var exceptiondate = new Date(except.OriginalDate);
			
			var secondsSinceDayStart = exceptiondate.getHours()*3600 + exceptiondate.getMinutes()*60+exceptiondate.getSeconds();
			if (secondsSinceDayStart == 0)
			{  
			    var start = new Date(item.start);
				exceptiondate.setHours(start.getHours());
				exceptiondate.setMinutes(start.getMinutes());
				exceptiondate.setSeconds(start.getSeconds());
			}
			recurEvent += formatDateTime(exceptiondate); 
			
			firstExcept = false;
		}
		recurEvent += linebreak;
		
	}
	
	
	return recurEvent;
	
}

function alert(string)
{
	
	var shell = new ActiveXObject( 'WScript.Shell' );
	shell.Popup( string );
	
}

function formatDate(date) {
    var oDate = new Date(date);
    icaldate = "" + oDate.getFullYear() + padzero((oDate.getMonth() + 1)) + padzero((oDate.getDate()));
    return icaldate;
}

function formatDateTime(date) {
    var oDate = new Date(date);
    icaldate = "" + oDate.getFullYear() + padzero((oDate.getMonth() + 1)) + padzero((oDate.getDate())) +
	"T" + padzero(oDate.getHours()) + padzero(oDate.getMinutes()) + padzero(oDate.getSeconds());
    return icaldate;
}

function daysOfWeek(week, pattern) {
    var mask = pattern.dayofweekmask;
    var daysOfWeek = "";
	
    if (mask & olMonday) {
        daysOfWeek = week + "MO";
	}
    if (mask & olTuesday) {
        if (daysOfWeek != "") { daysOfWeek += ","; }
        daysOfWeek += week + "TU";
	}
    if (mask & olWednesday) {
        if (daysOfWeek != "") { daysOfWeek += ","; }
        daysOfWeek += week + "WE";
	}
    if (mask & olThursday) {
        if (daysOfWeek != "") { daysOfWeek += ","; }
        daysOfWeek += week + "TH";
	}
    if (mask & olFriday) {
        if (daysOfWeek != "") { daysOfWeek += ","; }
        daysOfWeek += week + "FR";
	}
    if (mask & olSaturday) {
        if (daysOfWeek != "") { daysOfWeek += ","; }
        daysOfWeek += week + "SA";
	}
    if (mask & olSunday) {
        if (daysOfWeek != "") { daysOfWeek += ","; }
        daysOfWeek += week + "SU";
	}
	
    return daysOfWeek;
}

function weekNum(week) {
    if (week == 5) { 
        week = "-1"; 
		} else {
        padzero(week);
	}
    return week;
}

function monthNum(month) {
    var month = month + "";  // incase month comes in as a num
    month = month.toLowerCase().substr(0,3);
	
    var monthNum = 0;
	
    if (month == "jan") {
        monthNum = 1;
		} else if (month == "feb") {
        monthNum = 2;
		} else if (month == "mar") {
        monthNum = 3;
		} else if (month == "apr") {
        monthNum = 4;
		} else if (month == "may") {
        monthNum = 5;
		} else if (month == "jun") {
        monthNum = 6;
		} else if (month == "jul") {
        monthNum = 7;
		} else if (month == "aug") {
        monthNum = 8;
		} else if (month == "sep") {
        monthNum = 9;
		} else if (month == "oct") {
        monthNum = 10;
		} else if (month == "nov") {
        monthNum = 11;
		} else if (month == "dec") {
        monthNum = 12;
		} else {
        monthNum = month;
	}
	
    return monthNum;
}

function padzero(string) {
    if (String(string).length < 2) {
        string = "0" + string;
	}
    return string;
}

function cleanLineEndings(string) {
	// replace all ":"s with semicolons (why? not sure) but preserve tel:// URLs for formatted mobile friendly dial-in numbers.
    string = string.replace(/:\/\//g,'&#058;');
	string = string.replace(/:/g,'\;');
	string = string.replace(/&#058;/g, '://');
	
    string = string.replace(/\r/g,'\n');
    string = string.replace(/\n\n/g,'\n');
    string = string.replace(/\n/g,'\\n');
    string = string.replace(/,/g,'\,');
    return string;
}

function optionalOrganizer (organizer) {
	if (organizer == "" ) // "me" if not needed on every event
	{
		return "";
	}
	else
	{
		return " (" + cleanCommas (organizer) + ")";
	}
	
}

function getOptionalBody(item)
{
	if (item.body.length < 5)
		return "";
	
	if (null == item.body.match(/webex/ig))
		return "(removed)";
	
	return cleanLineEndings(item.body);
	
}



function cleanCommas(string) {
	// replace all ":"s with semicolons (why? not sure) but preserve tel:// URLs for formatted mobile friendly dial-in numbers.
    
	//	string = string.replace(/,/g,'');
    string = string.replace(/,\s+/g,'\\,');
    return string;
}


function randInt(min, max) 
{ 
	return Math.round(Math.random()*(max-min)+min) 
} 

function getInterval(interval)
{
	if (0 == interval)
	{ return ""; }
	return ";INTERVAL=" + interval;
}
