/* The Cycle-based Calendar Computation Engine/ CBCCE
Character set is UTF-8

The functions of this package performs intercalation computations for calendars
that set intercalation elements following regular cycles.
This applies to the Milesian calendar, that adds or substracts intercalary days only at end of cycles, including months.
This applies also to calendars that shift the long year by one at cyclic periods 
e.g. the long year comes after 5 years instead of 4 years in certain circumstances.
A possible algorithmic implementation of the French Revolutionary "Franciade" uses such cycles.
For other calendars, including Gregorian and Julian, this routines may be used to compute
the rank of a day within a year, and then hours, minutes, seconds and milliseconds.
Computations on months require more specific algorithms.
The principles of these routines are explained in "L'heure milésienne" (for the first version)
a book by Louis-Aimé de Fouquières.

Version for GAS : 
  M2018-02-19 - no change with respect to the original version,
    except: replace "let" statements with "var".
  M2018-06-12 - update comments
*/
/* Copyright Miletus 2016-2018 - Louis A. de Fouquières
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

var Chronos = { // Set of chronological constants generally used when handling Unix dates.
  DAY_UNIT : 86400000, // One day in Unix time units
  HOUR_UNIT : 3600000,
  MINUTE_UNIT : 60000,
  SECOND_UNIT : 1000
}
/* Parameter object structure. Replace # with numbers or literals.
var decomposeParameterExample = {
	timeepoch : #, //origin time in milliseconds (or in the suitable unit) to be used for the decomposition, with respect to 1/1/1970 00:00 UTC.
    coeff : [ // This array holds the coefficient used to decompose a time stamp into time cycles like eras, quadrisaeculae, centuries etc.
		{cyclelength : #, //length of the cycle, expressed in milliseconds.
		ceiling : #, // Infinity, or the maximum number of cycles of this size minus one in the upper cycle; the last cycle may hold an intercalary unit.
		subCycleShift : #, // number (-1, 0 or +1) to add to the ceiling of the cycle of the next level when the ceiling is reached at this level.
		multiplier : #, // multiplies the number of cycle of this level to convert into target units.
		target : #, // the unit (e.g. "year") of the decomposition element at this level. 
		} ,
		{ // similar elements at the lower cycle level 
		} // end of array element
	], // End of this array, but not end of object
	canvas : [ // this last array is the canvas of the decomposition , e.g. "year", "month", "date", with suitable properties at each level.
		{ name : #, // the name of the property at this level, which must match one target property of the coeff component,
		init : #, // value of this component at epoch, and lowest value (except for the first component), e.g. 0 for month, 1 for date, 0 for hours, minutes, seconds.
		} // End of array element (only two properties)
	] // End of second array
}	// End of object.
Constraints: 
  1. The cycles and the canvas elements shall be definined from the larger to the smaller 
        e.g. quadrisaeculum, then century, then quadriannum, then year, etc.
  2. The same names shall bu used for the "coeff" and the "canvas" properties, elsewise applications may return "NaN".	
*/
var
Day_milliseconds = { 	// To convert a time or a duration to and from days + milliseconds in day.
	timeepoch : 0, 
	coeff : [ // to be used with a Unix timestamp in ms. Decompose into days and milliseconds in day.
	  {cyclelength : 86400000, ceiling : Infinity, subCycleShift : 0, multiplier : 1, target : "day_number"}, 
	  {cyclelength : 1, ceiling : Infinity, subCycleShift : 0, multiplier : 1, target : "milliseconds_in_day"}
	],
	canvas : [
		{name : "day_number", init : 0},
		{name : "milliseconds_in_day", init : 0},
	]
}
/**
* cbcceDecompose : from time serie figure to compound object with calendar figures. 
* @param {number} quantity, the number of milliseconds from 1970-01-01T00:00:00.000Z
* @param {Object} params, the paramaters of the calendar.
* @return {Object}, an object composed as specified in "canvas" part of params.
*/
function cbcceDecompose (quantity, params) { // from a chronological number, build an compound object holding the elements as required by cparams.
  quantity -= params.timeepoch; // set at initial value the quantity to decompose into cycles.
  var result = new Object(); // Construct intitial compound result 
  var i, addCycle, r, ceiling; // let is not used in GAS
  for (i = 0; i < params.canvas.length; i++) {	// Define property of result object (a date or date-time) // let i
    Object.defineProperty (result, params.canvas[i].name, {enumerable : true, writable : true, value : params.canvas[i].init});
  }
  addCycle = 0; 	// flag that upper cycle has one element more or less (i.e. a 5 years franciade or 13 months luni-solar year) // let addCycle
  for (i = 0; i < params.coeff.length; ++i) {	// Perform decomposition by dividing by the successive cycle length //let i
    r = 0; // r is the computed quotient for this level of decomposition //let r
    if (params.coeff[i].cyclelength == 1) r = quantity; // avoid performing a trivial divison by 1.
    else {		// at each level, search at the same time the quotient (r) and the modulus (quantity)
      while (quantity < 0) {
        --r; 
        quantity += params.coeff[i].cyclelength;
      }
	  ceiling = params.coeff[i].ceiling + addCycle; //let ceiling
      while ((quantity >= params.coeff[i].cyclelength) && (r < ceiling)) {
        ++r; 
        quantity -= params.coeff[i].cyclelength;
      }
	  addCycle = (r == ceiling) ? params.coeff[i].subCycleShift : 0; // if at last section of this cycle, add or substract 1 to the ceiling of next cycle
    }
    result[params.coeff[i].target] += r*params.coeff[i].multiplier; // add result to suitable part of result array	
  }	
  return result;
}
/**
* cbcceCompose : from compound object to time series figure. 
* @param {Object} cells, contains the date elements following params.canvas
* @param {Object} params, the paramaters of the calendar.
* @return {number}, the number of milliseconds from 1970-01-01T00:00:00.000Z
*/
function cbcceCompose (cells, params) { // from an object cells structured as params.canvas, compute the chronological number
	var quantity = params.timeepoch ; // initialise Unix quantity to computation epoch
	var i, currentTarget, currentCounter, addCycle, f, ceiling; // let is not used in GAS
    for (i = 0; i < params.canvas.length; i++) { // cells value shifted as to have all 0 if at epoch // let i
		cells[params.canvas[i].name] -= params.canvas[i].init
	}
	currentTarget = params.coeff[0].target; 	// Set to uppermost unit used for date (year, most often) // let
	currentCounter = cells[params.coeff[0].target];	// This counter shall hold the successive remainders // let
	addCycle = 0; 	// This flag says whether there is an additionnal period at end of cycle, e.g. a 5th year in the Franciade or a 13th month // let
	for (i = 0; i < params.coeff.length; i++) { // let
		f = 0;				// Number of "target" values (number of years, to begin with) // let
		if (currentTarget != params.coeff[i].target) {	// If we go to the next level (e.g. year to month), reset variables
			currentTarget = params.coeff[i].target;
			currentCounter = cells[currentTarget];
		}
		ceiling = params.coeff[i].ceiling + addCycle; // let
		while (currentCounter < 0) {
			--f;
			currentCounter += params.coeff[i].multiplier;
		}
		while ((currentCounter >= params.coeff[i].multiplier) && (f < ceiling)) {
			++f;
			currentCounter -= params.coeff[i].multiplier;
		}
		addCycle = (f == ceiling) ? params.coeff[i].subCycleShift : 0;
		quantity += f * params.coeff[i].cyclelength;
	}
	return quantity ;	
}
