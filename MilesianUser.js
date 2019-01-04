/* MilesianUserFunctions : Enter and display in Google Sheets dates following Milesian calendar conventions
For use as a Google Application Script associated with a Google sheet.
This package is consistent with the similar Basic and VBA modules for OpenOffice and MS Excel.
Functions are aimed at extending Date & Time functions, and use similar parameters syntax in English
Versions GAS: 
  M2018-02-19: first release
  M2018-06-12: update comments and adapt to MilesianPrimitives
  M2019-01-15: solar intercalation rule is as of Gregorian
functions:
  MILESIAN_IS_LONG_YEAR
  MILESIAN_DATE
  MILESIAN_YEAR, MILESIAN_MONTH, MILESIAN_DAY, MILESIAN_DISPLAY
  MILESIAN_UTCYEAR, MILESIAN_UTCMONTH, MILESIAN_UTCDAY, MILESIAN_UTCDISPLAY
  TIMEZONE_OFFSET
  MILESIAN_MONTH_SHIFT,MILESIAN_MONTH_END
  JULIAN_EPOCH_COUNT,JULIAN_EPOCH_DATE
*/
/* Copyright Miletus SARL 2018. www.calendriermilesien.org
No warranty.
May be used for personal or professional purposes.
If transmitted or integrated, even with changes, present header shall be maintained in full.
*/
/* Implementation notes: 
When translated to GAS, the Google Sheets dates are considered local time. This gives generally non expected effects.
*/
var HighYear = 99999,  // Upper limit that yields a Google Sheet date value is 99999 as Gregorian year.
LowYear = -2,       // Bottom limit. We add the real day limit
LowCount = -694324; // The lowest MS Count that represents a valid date under Google Sheet.

//#Part 1: internal procedures -> brand other functions, to be added to CBCCE.
/**
* Positive modulo. Divisor must be positive. 
* @param {number} Dividend. If negative, the result shall be positive.
* @param {number} Divisor, must be positive, non zero, else return "undefined"
* @return {number} The modulo, positive or zero.
*/
function positiveModulo (dividend, divisor) {
	if (divisor <= 0) return ;	// Stop execution and return "Undefined"
	while (dividend < 0) dividend += divisor;
	while (dividend >= divisor) dividend -= divisor;
	return dividend
}
function MSCount_ (theDate) { // Translate a GS Date object into an Google Sheet counter.
const MSDaystoPosix = 25569; 
	return (theDate.valueOf() / Chronos.DAY_UNIT) + MSDaystoPosix;	
}
function pad(number) {	// utility function, pad 2-digit integer numbers. No control.
	return ( number < 10 ) ? ('0' + number) : number;
}

//#Part 2: a function not more used internally, but available to users
/**
* Whether the year is a milesian long year (366 days)
* @param {integer} the year in question; may be positive, 0 or negative.
* @return {boolean} true if long year, false if not, error if year is not integer
*/
function MILESIAN_IS_LONG_YEAR(Year) {
//Is year Year a 366 days year, i.e. a year just before a bissextile year following the Milesian rule.
if (Year !== Math.round(Year) || Year < LowYear || Year > HighYear) throw "MILESIAN_IS_LONG_YEAR: Invalid argument: " + Year; //Check that we have an integer numeric value
  Year += 1;
  return (positiveModulo (Year,4) == 0 && (positiveModulo (Year,100) !== 0 || positiveModulo(Year, 400) == 0 ));
} 

//#Part 3: Compute date (local time) from milesian parameters
/**
* Date computed from Milesian components. Error raised if components are irrelevant or out of range.
* @param {integer} year, positive, 0 or negative (from -2 to 100000)
* @param {integer} month, 1 to 12.
* @param {integer} day in month, 1 to 31. Date validity control is performed.
* @return {date} a Google Sheets day number (0 for 30 december 1899)
*/
function MILESIAN_DATE(Year, Month, DayInMonth) {
//Date number at 00:00 from a Milesian date given as year (positive or negative), month, daynumber in month
//Check that Milesian date is OK
	if (Year !== Math.round(Year) || Month !== Math.round(Month) || DayInMonth !== Math.round(DayInMonth)) 
		throw "MILESIAN_DATE: Invalid argument: " + Year +" "+ Month +" "+ DayInMonth;
	if (Year >= LowYear && Year <= HighYear && Month > 0 && Month < 13 && DayInMonth > 0 && DayInMonth < 32) { //Basic filter
		if (DayInMonth < 31 || (Month % 2 == 0) && (Month < 12 || MILESIAN_IS_LONG_YEAR(Year))) { 
			var theDate = setUTCDateFromMilesian (Year, Month-1, DayInMonth);
            var MyCount = MSCount_ (theDate);
            if (MyCount < LowCount) throw "MILESIAN_DATE: Out-of-range argument: " + Year +" "+ Month +" "+ DayInMonth
              else return MyCount; 
			}
		else	// Case where date elements do not build a correct milesian date
			throw "MILESIAN_DATE: Invalid date: " + Year +" "+ Month +" "+ DayInMonth;	
		}
	else		// Case where the date elements are outside basic values
		throw "MILESIAN_DATE: Out-of-range argument: " + Year +" "+ Month +" "+ DayInMonth;
}
/**
* the last day at 00:00 UTC before a Milesian year. Usefull for doomsday and epact. 
* @param {integer} the year in question; may be positive, 0 or negative.
* @return {boolean} date of the day before, as a Google Sheets count.
*/
function MILESIAN_YEAR_BASE(Year) { 
	if (Year !== Math.round(Year) || Year < LowYear || Year > HighYear) throw "MILESIAN_YEAR_BASE: Invalid argument: " + Year;
	var theDate = setUTCDateFromMilesian (Year, 0, 0) ; 
	return MSCount_(theDate);
}

//#Part 4: Extract Milesian elements from Date element, using getMilesianDate (theDate) that gives local time.
/**
* The Milesian year (common era, relative notation) for a given date.
* @param {date} the date being converted.
* @return {integer} the Milesian year, may be positive, 0 or negative.
*/
function MILESIAN_YEAR(TheDate) {
	var R = getMilesianDate (TheDate); 
	return R.year;
}
/**
* The Milesian month for a given date.
* @param {date} the date being converted.
* @return {integer} the Milesian month number, 1 to 12.
*/
function MILESIAN_MONTH(TheDate) { 
	var R = getMilesianDate (TheDate); 
	return R.month+1;	// under JS month begin with 0.
}
/**
* The Milesian day in month for a given date.
* @param {date} the date being converted.
* @return {integer} the Milesian day in month, 1 to 31.
*/
function MILESIAN_DAY(TheDate) { 
	var R = getMilesianDate (TheDate); 
	return R.date;
}
/**
* The time element of a given date, shows the conversion made between Google Sheets and GS
* @param {date} the date being converted.
* @return {number} a number greater or equal to 0 and lower than 1.
*/
function MILESIAN_TIME(TheDate) {
// In a first attempt, made a decomposition in { Days,  Milliseconds }, which is too long an operation.
var R = getMilesianDate (TheDate)
	return (R.hours * Chronos.HOUR_UNIT + R.minutes * Chronos.MINUTE_UNIT + R.seconds * Chronos.SECOND_UNIT + R.milliseconds) / Chronos.DAY_UNIT;
}
/**
* A string, displays the date in Milesian calendar.
* @param {date} the date being converted.
* @param {number} optional, if anything bt 0, the time is also displayed.
* @return {string} the milesian date and optionally the time.
*/
function MILESIAN_DISPLAY(TheDate, Wtime) {
	var R = getMilesianDate (TheDate); 
	var S = R.date + " " + (R.month+1) + "m " + R.year + (Wtime ? " " + pad(R.hours)+":"+pad(R.minutes)+":"+pad(R.seconds) : "");
	return S;
}
/**
* The Milesian year for a given date, UTC.
* @param {date} the date being converted.
* @return {integer} the Milesian year, may be positive, 0 or negative.
*/
function MILESIAN_UTCYEAR(TheDate) {
var R = getUTCMilesianDate (TheDate); 
	return R.year;
}
/**
* The Milesian month for a given date, UTC.
* @param {date} the date being converted.
* @return {integer} the Milesian month number, 1 to 12.
*/
function MILESIAN_UTCMONTH(TheDate) {
	var R = getUTCMilesianDate (TheDate); 
	return R.month+1;	// under JS month begin with 0.
}
/**
* The Milesian day in month for a given date, UTC.
* @param {date} the date being converted.
* @return {integer} the Milesian day in month, 1 to 31.
*/
function MILESIAN_UTCDAY(TheDate) { 
	var R = getUTCMilesianDate (TheDate); 
	return R.date;
}
/**
* The time element of a given date, UTC
* @param {date} the date being converted.
* @return {number} a number greater or equal to 0 and lower than 1.
*/
function MILESIAN_UTCTIME(TheDate) {
	var R = getUTCMilesianDate (TheDate);
	return (R.hours * Chronos.HOUR_UNIT + R.minutes * Chronos.MINUTE_UNIT + R.seconds * Chronos.SECOND_UNIT + R.milliseconds) / Chronos.DAY_UNIT;
}
/**
* A string, displays the date in Milesian calendar, UTC.
* @param {date} the date being converted.
* @param {number} optional, if anything bt 0, the time is also displayed.
* @return {string} the milesian date and optionally the time, UTC.
*/
function MILESIAN_UTCDISPLAY(TheDate, Wtime) { 
	var R = getUTCMilesianDate (TheDate); 
	var S = R.date + " " + (R.month+1) + "m " + R.year + (Wtime ? " " + pad(R.hours)+":"+pad(R.minutes)+":"+pad(R.seconds) : "");
	return S;
}
/**
* The timezone offset of a date, in minutes.
* @param {date} the date in question
* @return {integer} the number of minutes added to local time in order to get UTC
*/
function TIMEZONE_OFFSET(TheDate) { // The timezone offset of a (JS) Date object, at a given date.
  return TheDate.getTimezoneOffset();
}

//#Part 5: Computations on milesian months

var
Year_Month_Params = { // to be used in order to shift months, or change lunar calendar epoch without changing lunar age.
	timeepoch : 0, // put the timeepoch in the parameter call.
	coeff : [
		{cyclelength : 12, ceiling : Infinity, subCycleShift : 0, multiplier : 1, target : "year"},
		{cyclelength : 1,  ceiling : Infinity, subCycleShift : 0, multiplier : 1, target : "month"}
	],
	canvas : [
		{name : "year", init :0},
		{name : "month", init : 0}		
	]
};
/**
* A date several Milesian month later of earlier.
* @param {date} start date.
* @param {number} number of month, positive if later, negative if earlier. 
* @return {date} the Milesian date (and time). If start day is 31, target date may be 30.
*/
function MILESIAN_MONTH_SHIFT (TheDate, MonthShift) { //Same date several (milesian) months later of earlier
	if (MonthShift !== Math.round(MonthShift)) throw "MILESIAN_MONTH_SHIFT: Invalid argument: " + Monthshift;
	var MyMil = getMilesianDate (TheDate) ; 	// Construct start date	
	var Shift = cbcceDecompose (MonthShift + MyMil.month,Year_Month_Params) ;	// compute new month number and year shift;
	var year = MyMil.year + Shift.year; 	// year of target date
	if (year < LowYear || year > HighYear) throw "MILESIAN_MONTH_SHIFT: Out-of-range : " + year + " " + Shift.month
	else {	// Construct result
		if (MyMil.date == 31) // In this case maybe we should change this figure to 30
			MyMil.date = ((Shift.month % 2 == 1) && (Shift.month < 11 || MILESIAN_IS_LONG_YEAR(year)) ? 31 : 30);
		MyMil.year = year;
		MyMil.month = Shift.month;
		return MSCount_ (cbcceCompose (MyMil, Milesian_time_params)); 
		}
	}	
/**
* Date of last day in month several Milesian month later of earlier.
* @param {date} start date.
* @param {number} number of month, positive if later, negative if earlier. 
* @return {date} the Milesian date (and time) at end of Milesian month.
*/
function MILESIAN_MONTH_END (TheDate, MonthShift) { //End of month, several (milesian) months later of earlier
	if (MonthShift !== Math.round(MonthShift)) throw "MILESIAN_MONTH_END: Invalid argument: " + Monthshift;
	var MyMil = getMilesianDate (TheDate) ; 	// Construct start date	
	var Shift = cbcceDecompose (MonthShift + MyMil.month, Year_Month_Params) ;	// compute new month number and year shift;
	var year = MyMil.year + Shift.year; 	// year of target date
	if (year < LowYear || year > HighYear) throw "MILESIAN_MONTH_SHIFT: Out-of-range : " + year + " " + Shift.month
	else {	// Construct result
		MyMil.date = ((Shift.month % 2 == 1) && (Shift.month < 11 || MILESIAN_IS_LONG_YEAR(year)) ? 31 : 30);
		MyMil.year = year;
		MyMil.month = Shift.month;
		return MSCount_ (cbcceCompose (MyMil, Milesian_time_params)); 
		}
	}

//#Part 6: Julian Epoch Day conversion functions

var  JULIAN_DAY_UTC0_EPOCH_OFFSET = 210866803200000; // Julian Day 0 at 0h00 UTC.
/**
* The number of the Julian day corresponding to a date with time. Always UTC.
* @param {date} the date being converted
* @return {number} the Julian day number
*/
function JULIAN_EPOCH_COUNT(TheDate) {
	return (TheDate.getTime() + JULIAN_DAY_UTC0_EPOCH_OFFSET - (12 * Chronos.HOUR_UNIT)) / Chronos.DAY_UNIT;
	}	
/**
* Compute a date corresponding to a Julian day number. 
* @param {number} the Julian day to convert
* @return {date} the corresponding date
*/
function JULIAN_EPOCH_DATE(Julian_Count) {
	var MyDate = new Date (0);
	MyDate.setTime (Math.round(Julian_Count*Chronos.DAY_UNIT) - JULIAN_DAY_UTC0_EPOCH_OFFSET + 12 * Chronos.HOUR_UNIT);
	return MyDate;
	}	
