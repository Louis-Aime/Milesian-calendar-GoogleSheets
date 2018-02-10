//Milesian Calendar: Enter and display in Google Sheets dates following Milesian calendar conventions
//Copyright Miletus SARL 2018. www.calendriermilesien.org
//For use as a Google Application Script associated with a Google sheet.
//This package is consistent with the similar Basic and VBA modules for OpenOffice and MS Excel.
//No warranty.
//May be used for personal or professional purposes.
//If transmitted or integrated, even with changes, present header shall be maintained in full.
//Functions are aimed at extending Date & Time functions, and use similar parameters syntax in English
//Version GAS M2018-02-21.
//
//functions:
//MILESIAN_IS_LONG_YEAR
//MILESIAN_DATE
//MILESIAN_YEAR, MILESIAN_MONTH, MILESIAN_DAY, MILESIAN_DISPLAY
//MILESIAN_UTCYEAR, MILESIAN_UTCMONTH, MILESIAN_UTCDAY, MILESIAN_UTCDISPLAY
//TIMEZONE_OFFSET

/* Implementation notes: 
When translated to GAS, the Google Sheets dates are considered local time. This gives generally non expected effects.
*/
var HighYear = 99999,  // Upper limit that yields a Google Sheet date value is 99999 as Gregorian year.
LowYear = -2,       // Bottom limit. We add the real day limit
LowCount = -694324; // The lowest MS Count that represents a valid date under Google Sheet.

//#Part 1: internal procedures -> brand other functions, to be added to CBCCE.

function positiveModulo (dividend, divisor) {	// Positive modulo, with only positive divisor
	if (divisor <= 0) return ;	// Stop execution and return "Undefined"
	while (dividend < 0) dividend += divisor;
	while (dividend >= divisor) dividend -= divisor;
	return dividend
}

function MSCount (theDate) { // Translate a JS Date object into a MS counter.
const MSDaystoPosix = 25569; 
	return (theDate.valueOf() / Chronos.DAY_UNIT) + MSDaystoPosix;	
}

function pad(number) {	// utility function, pad 2-digit integer numbers. No control.
	return ( number < 10 ) ? ('0' + number) : number;
}

//#Part 2: a function not more used internally, but available to users

function MILESIAN_IS_LONG_YEAR(Year) { // As Boolean
//Is year Year a 366 days year, i.e. a year just before a bissextile year following the Milesian rule.
if (Year !== Math.round(Year) || Year < LowYear || Year > HighYear) throw "MILESIAN_IS_LONG_YEAR: Invalid argument: " + Year; //Check that we have an integer numeric value
  Year += 1;
  return (positiveModulo (Year,4) == 0 && (positiveModulo (Year,100) !== 0 || (positiveModulo(Year, 400) == 0 && positiveModulo(Year+800, 3200) !== 0)));
} //End Function

//#Part 3: Compute date (local time) from milesian parameters

function MILESIAN_DATE(Year, Month, DayInMonth) { //First compute a Posix UTC date, then convert into a MS count.
//Date number from a Milesian date given as year (positive or negative), month, daynumber in month
//Check that Milesian date is OK
	if (Year !== Math.round(Year) || Month !== Math.round(Month) || DayInMonth !== Math.round(DayInMonth)) 
		throw "MILESIAN_DATE: Invalid argument: " + Year +" "+ Month +" "+ DayInMonth;
	if (Year >= LowYear && Year <= HighYear && Month > 0 && Month < 13 && DayInMonth > 0 && DayInMonth < 32) { //Basic filter
		if (DayInMonth < 31 || (Month % 2 == 0) && (Month < 12 || MILESIAN_IS_LONG_YEAR(Year))) { 
			var theDate = setUTCDateFromMilesian (Year, Month-1, DayInMonth);
            var MyCount = MSCount (theDate);
            if (MyCount < LowCount) throw "MILESIAN_DATE: Out-of-range argument: " + Year +" "+ Month +" "+ DayInMonth
              else return MyCount; 
			}
		else	// Case where date elements do not build a correct milesian date
			throw "MILESIAN_DATE: Invalid date: " + Year +" "+ Month +" "+ DayInMonth;	
		}
	else		// Case where the date elements are outside basic values
		throw "MILESIAN_DATE: Out-of-range argument: " + Year +" "+ Month +" "+ DayInMonth;
}

function MILESIAN_YEAR_BASE(Year) { //The UTC Year base or Doomsday of a year i.e. the date just before the 1 1m of the year
	if (Year !== Math.round(Year) || Year < LowYear || Year > HighYear) throw "MILESIAN_YEAR_BASE: Invalid argument: " + Year;
	var theDate = setUTCDateFromMilesian (Year, 0, 0) ; 
	return MSCount(theDate);
}

//#Part 4: Extract Milesian elements from Date element, using getMilesianDate (theDate) that gives local time.

function MILESIAN_YEAR(TheDate) { //The milesian year (common era) for a Date argument (a series number or a string)
	var R = getMilesianDate (TheDate); 
	return R.year;
}

function MILESIAN_MONTH(TheDate) { //The milesian month number (1-12) for a Date argument
	var R = getMilesianDate (TheDate); 
	return R.month+1;	// under JS month begin with 0.
}

function MILESIAN_DAY(TheDate) { //The day number in the milesian month for a Date argument
	var R = getMilesianDate (TheDate); 
	return R.date;
}

function MILESIAN_TIME(TheDate) { //The local time element in the Date argument, as a decimal part of 1.
	var R = cbcceDecompose (TheDate - (TheDate.getTimezoneOffset() * Chronos.MINUTE_UNIT), Day_milliseconds); //Use the simpliest decomposition canvas
	return R.milliseconds_in_day / Chronos.DAY_UNIT;
}

function MILESIAN_DISPLAY(TheDate, Wtime) {
//Milesian date as a string, for a Date argument
// If Wtime is anything but 0 or undefined, the time is also displayed.
	var R = getMilesianDate (TheDate); 
	var S = R.date + " " + (R.month+1) + "m " + R.year + (Wtime ? " " + pad(R.hours)+":"+pad(R.minutes)+":"+pad(R.seconds) : "");
	return S;
}
function MILESIAN_UTCYEAR(TheDate) { //The milesian year (common era) for a Date argument (a series number or a string)
	var R = getMilesianUTCDate (TheDate); 
	return R.year;
}

function MILESIAN_UTCMONTH(TheDate) { //The milesian month number (1-12) for a Date argument
	var R = getMilesianUTCDate (TheDate); 
	return R.month+1;	// under JS month begin with 0.
}

function MILESIAN_UTCDAY(TheDate) { //The day number in the milesian month for a Date argument
	var R = getMilesianUTCDate (TheDate); 
	return R.date;
}

function MILESIAN_UTCTIME(TheDate) { //The local time element in the Date argument, as a decimal part of 1.
	var R = cbcceDecompose (TheDate ); //Use the simpliest decomposition canvas
	return R.milliseconds_in_day / Chronos.DAY_UNIT;
}

function MILESIAN_UTCDISPLAY(TheDate, Wtime) { 
//Milesian date as a string, for a Date argument
// If Wtime is anything but 0 or undefined, the time is also displayed.
	var R = getMilesianUTCDate (TheDate); 
	var S = R.date + " " + (R.month+1) + "m " + R.year + (Wtime ? " " + pad(R.hours)+":"+pad(R.minutes)+":"+pad(R.seconds) : "");
	return S;
}
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
		return MSCount (cbcceCompose (MyMil, Milesian_time_params)); 
		}
	}
	
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
		return MSCount (cbcceCompose (MyMil, Milesian_time_params)); 
		}
	}

//#Part 6: Julian Epoch Day conversion functions

var  JULIAN_DAY_UTC0_EPOCH_OFFSET = 210866803200000; // Julian Day 0 at 0h00 UTC.

function JULIAN_EPOCH_COUNT(TheDate) {
	return (TheDate.getTime() + JULIAN_DAY_UTC0_EPOCH_OFFSET - (12 * Chronos.HOUR_UNIT)) / Chronos.DAY_UNIT;
	}	

function JULIAN_EPOCH_DATE(Julian_Count) {
	var MyDate = new Date (0);
	MyDate.setTime (Math.round(Julian_Count*Chronos.DAY_UNIT) - JULIAN_DAY_UTC0_EPOCH_OFFSET + 12 * Chronos.HOUR_UNIT);
	return MyDate;
	}	
