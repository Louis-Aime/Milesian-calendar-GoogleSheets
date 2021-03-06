/* Milesian date computations functions for Google Apps Script
Character set is UTF-8
Custom functions for Google sheets in order to compute and display Milesian date elements
using the CBCCE, since a Date object uses Posix conventions here.
Versions: 
  M2018-02-21 : adapt to GAS
  M2018-06-12 : Change getMilesianUTCDate to getUTCMilesianDate (compatibility with JS version)
  M2019-01-15 : Solar intercalation rule is same as Gregorian
Package CBCCE is used. 
Milesian month names are not used. Here the only output uses the simple international code.
 getMilesianDate : the day date as a three elements object: .year, .month, .date; .month is 0 to 11. Conversion is in local time.
 getUTCMilesianDate : same as above, in UTC time.
 setTimeFromMilesian (year, month, date, hours) : set Time from milesian date at 00h local hour.
 setUTCTimeFromMilesian (year, month, date, hours) : same but at 00h UTC.
 toIntlMilesianDateString : return a string with the date elements in Milesian: (day) (month)"m" (year), month 1 to 12.
 toUTCIntlMilesianDateString : same as above, in UTC time zone.
*/
/* Copyright Miletus 2016-2019 - Louis A. de Fouqui�res
Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:
  1. The above copyright notice and this permission notice shall be included
  in all copies or substantial portions of the Software.
  2. Changes with respect to any former version shall be documented.

The software is provided "as is", without warranty of any kind,
express of implied, including but not limited to the warranties of
merchantability, fitness for a particular purpose and noninfringement.
In no event shall the authors of copyright holders be liable for any
claim, damages or other liability, whether in an action of contract,
tort or otherwise, arising from, out of or in connection with the software
or the use or other dealings in the software.
Inquiries: www.calendriermilesien.org
*/
//
// 1. Basic tools of this package
//
/* Import or make visible: 
	CBCCE.js
*/
var Milesian_time_params = { // To be used with a Unix timestamp in ms. Decompose into Milesian years, months, date, hours, minutes, seconds, ms
	timeepoch : -62168083200000, // Unix timestamp of 1 1m 0 00h00 UTC in ms
	coeff : [ 
	  {cyclelength : 12622780800000, ceiling : Infinity, subCycleShift : 0, multiplier : 400, target : "year"},
	  {cyclelength : 3155673600000, ceiling :  3, subCycleShift : 0, multiplier : 100, target : "year"},
	  {cyclelength : 126230400000, ceiling : Infinity, subCycleShift : 0, multiplier : 4, target : "year"},
	  {cyclelength : 31536000000, ceiling : 3, subCycleShift : 0, multiplier : 1, target : "year"},
	  {cyclelength : 5270400000, ceiling : Infinity, subCycleShift : 0, multiplier : 2, target : "month"},
	  {cyclelength : 2592000000, ceiling : 1, subCycleShift : 0, multiplier : 1, target : "month"}, 
	  {cyclelength : 86400000, ceiling : Infinity, subCycleShift : 0, multiplier : 1, target : "date"},
	  {cyclelength : 3600000, ceiling : Infinity, subCycleShift : 0, multiplier : 1, target : "hours"},
	  {cyclelength : 60000, ceiling : Infinity, subCycleShift : 0, multiplier : 1, target : "minutes"},
	  {cyclelength : 1000, ceiling : Infinity, subCycleShift : 0, multiplier : 1, target : "seconds"},
	  {cyclelength : 1, ceiling : Infinity, subCycleShift : 0, multiplier : 1, target : "milliseconds"}
	],
	canvas : [ 
		{name : "year", init : 0},
		{name : "month", init : 0},
		{name : "date", init : 1},
		{name : "hours", init : 0},
		{name : "minutes", init : 0},
		{name : "seconds", init : 0},
		{name : "milliseconds", init : 0},
	]
}
//
// 2. Function which used to be methods for the Date object for Milesian dates.
//
/**
* getMilesianDate(theDate) : from time serie figure to compound Milesian date, local time. 
* @param {Date} theDate, the number of milliseconds from 1970-01-01T00:00:00.000Z
* @return {Object}, an object composed with elements of the complete system local date in Milesian, see "canvas".
*/
function getMilesianDate(theDate) { //Extract Milesian components of a Date object 
//Date.prototype.getMilesianDate = function () {
  return cbcceDecompose (theDate.getTime() - (theDate.getTimezoneOffset() * Chronos.MINUTE_UNIT), Milesian_time_params);
}
/**
* getUTCMilesianDate(theDate) : from time serie figure to compound Milesian date, UTC. 
* @param {Date} theDate, the number of milliseconds from 1970-01-01T00:00:00.000Z
* @return {Object}, an object composed with elements of the complete UTC date in Milesian, see "canvas".
*/
function getUTCMilesianDate (theDate) {
//Date.prototype.getMilesianUTCDate = function () {
  return cbcceDecompose (theDate.getTime(), Milesian_time_params);
}
/**
* setDateFromMilesian (year, month, date) : set date at 00h from Milesian local date figures,  
* @param {number} year, the year in the common (Christian) era, in relative i.e.: 0 and negative years are possible, year 0 is 1 B.C.
* @param {number} month, the Milesian month number 0 to 11 for 1m to 12m
* @param {number} date, the day number in the month, 1 to 31. If day number does not exist in month, computation is made from date 0 of this month, e.g. 31 1m -> 1 2m 
* @return {Object}, an object composed with elements of the complete local date in Milesian, see "canvas".
*/
function setDateFromMilesian (year, month, date) { //Compose a JS Date object with Milesian element. Time is 00h local time and date.
  var theDate = new Date(0);		// Create a new date object, initialize to 0.									   
  theDate.setTime(cbcceCompose({    // This computes UTC time.
	  'year' : year, 'month' : month, 'date' : date, 'hours' : 0, 'minutes' : 0, 'seconds' : 0, 'milliseconds' : 0
	  }, Milesian_time_params));
  theDate.setTime (theDate - theDate.getTimezoneOffset() * Chronos.MINUTE_UNIT); // Compensate UTC date with local and timed time-zone shift
  return theDate;
}
/**
* setUTCDateFromMilesian (year, month, date) : set date at 00h from Milesian UTC date figures,  
* @param {number} year, the year in the common (Christian) era, in relative i.e.: 0 and negative years are possible, year 0 is 1 B.C.
* @param {number} month, the Milesian month number 0 to 11 for 1m to 12m
* @param {number} date, the day number in the month, 1 to 31. If day number does not exist in month, computation is made from date 0 of this month, e.g. 31 1m -> 1 2m 
* @return {Object}, an object composed with elements of the complete UTC date in Milesian, see "canvas".
*/
function setUTCDateFromMilesian (year, month, date) {
//Date.prototype.setUTCTimeFromMilesian = function (year, month = 0, date = 1,
                                              // hours = this.getUTCHours(), minutes = this.getUTCMinutes(), seconds = this.getUTCSeconds(),
                                              // milliseconds = this.getUTCMilliseconds()) {
  var theDate = new Date(0);		// Create a new date object, initialize to 0.									   
  theDate.setTime(cbcceCompose({
	  'year' : year, 'month' : month, 'date' : date, 'hours' : 0, 'minutes' : 0, 'seconds' : 0, 'milliseconds' : 0
	  }, Milesian_time_params));
   return theDate;
}
/**
* toIntlMilesianDateString (theDate) : write a string displaying the Milesian local date as day-in-month, month (xm), year (possibly with minus sign).  
* @param {Date} theDate, the number of milliseconds from 1970-01-01T00:00:00.000Z
* @return {String}, a string giving the local date in Milesian.
*/
function toIntlMilesianDateString (theDate) {
//Date.prototype.toIntlMilesianDateString = function () {
	var dateElements = cbcceDecompose (theDate.getTime()- (theDate.getTimezoneOffset() * Chronos.MINUTE_UNIT), Milesian_time_params );
	return dateElements.date+" "+(++dateElements.month)+"m "+dateElements.year;
}
/**
* toUTCIntlMilesianDateString (theDate) : write a string displaying the Milesian UTC date as day-in-month, month (xm), year (possibly with minus sign).  
* @param {Date} theDate, the number of milliseconds from 1970-01-01T00:00:00.000Z
* @return {String}, a string giving the UTC date in Milesian.
*/
function toUTCIntlMilesianDateString (theDate) {
//Date.prototype.toUTCIntlMilesianDateString = function () {
	var dateElements = cbcceDecompose (theDate.getTime(), Milesian_time_params );
	return dateElements.date+" "+(++dateElements.month)+"m "+dateElements.year;
}
