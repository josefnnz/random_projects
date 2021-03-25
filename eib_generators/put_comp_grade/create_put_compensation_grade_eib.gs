/*
 * ASSUMPTIONS:
 * - Generate EIB to create brand new comp grades and grade profiles.
 * - Each default pay range has all zeroes.
 * - Each pay range has 3 segments.
 * - Each grade profile does not calculate segments.
 * - Each pay range prohibits override.
 * - Comp Grade and associated Comp Grade Profiles use same Effective Date.
 * - Each grade profile uses exactly one eligibility rule.
 * - Number of Segments = 3
 * - Comp Element = Standard Base Pay
 * - Effective Date = "1900-01-01"
 * - Add Only = Y
 * - Default Pay Ranges = 0
 * - Allow Override = N
 * - Primary Comp Basis ranges have a zero Total Base Pay range
 */

// create alphabet for column references
var NUM_COLUMNS = 300;
var alphabet = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];
var column_letters = alphabet;
for (var i = 26; i < NUM_COLUMNS; i++) {
	var quotient = Math.trunc(i / 26);
	var remainder = i % 26;
	column_letters[i] = column_letters[quotient - 1] + column_letters[remainder];
}

var column_array_indices = {};
for (var i = 0; i < column_letters.length; i++) {
	column_array_indices[column_letters[i]] = i;
}

// column number index for data source sheet
var INDEX_COMP_GRADE_NAME = column_array_indices["B"];
var INDEX_COMP_GRADE_EFFECTIVE_DATE = column_array_indices["E"];
var INDEX_COMP_GRADE_INACTIVE = column_array_indices["F"];
var INDEX_COMP_GRADE_NUM_SEGMENTS = column_array_indices["G"];
var INDEX_COMP_GRADE_COMP_ELEMENT = column_array_indices["I"];
var INDEX_COMP_GRADE_CURRENCY = column_array_indices["J"];
var INDEX_COMP_GRADE_FREQUENCY = column_array_indices["K"];
var INDEX_COMP_GRADE_COMP_BASIS = column_array_indices["L"];
var INDEX_COMP_GRADE_PROFILE_NAME = column_array_indices["U"];
var INDEX_COMP_GRADE_PROFILE_ELIGIBILITY_RULE = column_array_indices["W"];
var INDEX_COMP_GRADE_PROFILE_COMP_ELEMENT = column_array_indices["X"];
var INDEX_COMP_GRADE_PROFILE_CURRENCY = column_array_indices["Y"];
var INDEX_COMP_GRADE_PROFILE_FREQUENCY = column_array_indices["Z"];
var INDEX_COMP_GRADE_PROFILE_COMP_BASIS = column_array_indices["AA"];
var INDEX_COMP_GRADE_PROFILE_MIN = column_array_indices["AC"];
var INDEX_COMP_GRADE_PROFILE_MID = column_array_indices["AD"];
var INDEX_COMP_GRADE_PROFILE_MAX = column_array_indices["AE"];
var INDEX_COMP_GRADE_PROFILE_SEGMENT_1_TOP = column_array_indices["AF"];
var INDEX_COMP_GRADE_PROFILE_SEGMENT_2_TOP = column_array_indices["AG"];

// column number index for required EIB columns
var INDEX_EIB_COMP_GRADE_SPREADSHEET_KEY = column_array_indices["B"];
var INDEX_EIB_COMP_GRADE_ADD_ONLY = column_array_indices["C"];
var INDEX_EIB_COMP_GRADE_ID = column_array_indices["E"];
var INDEX_EIB_COMP_GRADE_EFFECTIVE_DATE = column_array_indices["F"];
var INDEX_EIB_COMP_GRADE_NAME = column_array_indices["G"];
var INDEX_EIB_COMP_GRADE_COMP_ELEMENT = column_array_indices["I"];
var INDEX_EIB_COMP_GRADE_NUM_SEGMENTS = column_array_indices["K"];
var INDEX_EIB_COMP_GRADE_MIN = column_array_indices["L"];
var INDEX_EIB_COMP_GRADE_MID = column_array_indices["M"];
var INDEX_EIB_COMP_GRADE_MAX = column_array_indices["N"];
var INDEX_EIB_COMP_GRADE_SEGMENT_1_TOP = column_array_indices["P"];
var INDEX_EIB_COMP_GRADE_SEGMENT_2_TOP = column_array_indices["Q"];
var INDEX_EIB_COMP_GRADE_CURRENCY = column_array_indices["U"];
var INDEX_EIB_COMP_GRADE_FREQUENCY = column_array_indices["V"];
var INDEX_EIB_COMP_GRADE_ALLOW_OVERRIDE = column_array_indices["X"];
var INDEX_EIB_COMP_GRADE_PROFILE_ROW_ID = column_array_indices["AJ"];
var INDEX_EIB_COMP_GRADE_PROFILE_DELETE = column_array_indices["AK"];
var INDEX_EIB_COMP_GRADE_PROFILE_ID = column_array_indices["AM"];
var INDEX_EIB_COMP_GRADE_PROFILE_EFFECTIVE_DATE = column_array_indices["AN"];
var INDEX_EIB_COMP_GRADE_PROFILE_NAME = column_array_indices["AO"];
var INDEX_EIB_COMP_GRADE_PROFILE_COMP_ELEMENT = column_array_indices["AQ"];
var INDEX_EIB_COMP_GRADE_PROFILE_ELIGIBILITY_RULE = column_array_indices["AR"];
var INDEX_EIB_COMP_GRADE_PROFILE_INACTIVE = column_array_indices["AS"];
var INDEX_EIB_COMP_GRADE_PROFILE_NUM_SEGMENTS = column_array_indices["AT"];
var INDEX_EIB_COMP_GRADE_PROFILE_MIN = column_array_indices["AU"];
var INDEX_EIB_COMP_GRADE_PROFILE_MID = column_array_indices["AV"];
var INDEX_EIB_COMP_GRADE_PROFILE_MAX = column_array_indices["AW"];
var INDEX_EIB_COMP_GRADE_PROFILE_SEGMENT_1_TOP = column_array_indices["AY"];
var INDEX_EIB_COMP_GRADE_PROFILE_SEGMENT_2_TOP = column_array_indices["AZ"];
var INDEX_EIB_COMP_GRADE_PROFILE_CURRENCY = column_array_indices["BD"];
var INDEX_EIB_COMP_GRADE_PROFILE_FREQUENCY = column_array_indices["BE"];
var INDEX_EIB_COMP_GRADE_PROFILE_ALLOW_OVERRIDE = column_array_indices["BG"];
var INDEX_EIB_COMP_BASIS_PROFILE_ROW_ID = column_array_indices["BT"];
var INDEX_EIB_COMP_BASIS_PROFILE_DELETE = column_array_indices["BU"];
var INDEX_EIB_COMP_BASIS_PROFILE_COMP_BASIS = column_array_indices["BV"];
var INDEX_EIB_COMP_BASIS_PROFILE_MIN = column_array_indices["BW"];
var INDEX_EIB_COMP_BASIS_PROFILE_MID = column_array_indices["BX"];
var INDEX_EIB_COMP_BASIS_PROFILE_MAX = column_array_indices["BY"];
var INDEX_EIB_COMP_BASIS_PROFILE_SEGMENT_1_TOP = column_array_indices["CA"];
var INDEX_EIB_COMP_BASIS_PROFILE_SEGMENT_2_TOP = column_array_indices["CB"];

// use a javascript set to find the unique list of compensation grades
var FIRST_COL_EXTRACTED = column_array_indices["A"] + 1; // VZ job code column
var LAST_COL_EXTRACTED = column_array_indices["AH"] + 1; // VZ union job column
var FIRST_ROW_EXTRACTED = 2; // first VZ job
var LAST_ROW_EXTRACTED = 32; // last VZ job
var NUM_ROWS_TO_EXTRACT = LAST_ROW_EXTRACTED - FIRST_ROW_EXTRACTED + 1;
var NUM_COLS_TO_EXTRACT = LAST_COL_EXTRACTED - FIRST_COL_EXTRACTED + 1;
var GOOGLE_ID_SS_COMP_GRADES = "103-ZG2ZaQuuPiAni_MeRcYrI5P1M2r2tstoyz61gT-A";
var SHEET_NAME_COMP_GRADES = "DATA_SOURCE";

function create_eib() {
	// load sheet with VZ job profiles
	var sheet_payranges = SpreadsheetApp.openById(GOOGLE_ID_SS_COMP_GRADES).getSheetByName(SHEET_NAME_COMP_GRADES);
	// extract VZ job profiles
	var values_payranges = sheet_payranges.getRange(FIRST_ROW_EXTRACTED, FIRST_COL_EXTRACTED, NUM_ROWS_TO_EXTRACT, NUM_COLS_TO_EXTRACT).getValues();
	var NUM_PAYRANGES = values_payranges.length;

	var unique_grades = new Set();
	for (var i = 0; i < NUM_PAYRANGES; i++) {
		unique_grades.add(values_payranges[i][INDEX_COMP_GRADE_NAME]);
	}

	var NUM_EIB_ROWS = NUM_ROWS_TO_EXTRACT;
	var NUM_EIB_COLS = column_array_indices["CU"] + 1;
	var eib = new Array(NUM_ROWS_TO_EXTRACT);
	for (var i = 0; i < NUM_EIB_ROWS; i++) {
		eib[i] = new Array(NUM_EIB_COLS);
	}

	var row_eib = 0;
	var grades = Array.from(unique_grades); // convert set to array
	var NUM_GRADES = grades.length;
	for (var row = 0; row < NUM_GRADES; row++) {
		// extract current job
		var curr_grade = grades[row];
		var curr_payranges = values_payranges.filter(function(r) {return r[INDEX_COMP_GRADE_NAME] === curr_grade});
		var NUM_GRADE_PROFILES = curr_payranges.length;

		var curr_default_payrange = curr_payranges[0];

		// fill out spreadsheet key column
		for (var i = 0; i < NUM_GRADE_PROFILES; i++) {
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_SPREADSHEET_KEY] = row + 1;
		}

		// fill out default pay range info
		eib[row_eib][INDEX_EIB_COMP_GRADE_ADD_ONLY] = "Y";
		eib[row_eib][INDEX_EIB_COMP_GRADE_ID] = curr_default_payrange[INDEX_COMP_GRADE_NAME];
		eib[row_eib][INDEX_EIB_COMP_GRADE_EFFECTIVE_DATE] = "1900-01-01";
		eib[row_eib][INDEX_EIB_COMP_GRADE_NAME] = curr_default_payrange[INDEX_COMP_GRADE_NAME];
		eib[row_eib][INDEX_EIB_COMP_GRADE_COMP_ELEMENT] = "Standard_Base_Pay";
		eib[row_eib][INDEX_EIB_COMP_GRADE_NUM_SEGMENTS] = 3;
		eib[row_eib][INDEX_EIB_COMP_GRADE_MIN] = 0;
		eib[row_eib][INDEX_EIB_COMP_GRADE_MID] = 0;
		eib[row_eib][INDEX_EIB_COMP_GRADE_MAX] = 0;
		eib[row_eib][INDEX_EIB_COMP_GRADE_SEGMENT_1_TOP] = 0;
		eib[row_eib][INDEX_EIB_COMP_GRADE_SEGMENT_2_TOP] = 0;
		eib[row_eib][INDEX_EIB_COMP_GRADE_CURRENCY] = curr_default_payrange[INDEX_COMP_GRADE_PROFILE_CURRENCY];
		eib[row_eib][INDEX_EIB_COMP_GRADE_FREQUENCY] = curr_default_payrange[INDEX_COMP_GRADE_FREQUENCY];
		eib[row_eib][INDEX_EIB_COMP_GRADE_ALLOW_OVERRIDE] = "N";

		// fill out comp grade profiles info
		for (var i = 0; i < NUM_GRADE_PROFILES; i++) {
			var grade_profile = curr_payranges[i];
			var grade_profile_name = grade_profile[INDEX_COMP_GRADE_PROFILE_NAME];
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_ROW_ID] = i + 1;
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_DELETE] = "N";
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_ID] = curr_grade + "_" + grade_profile_name;
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_EFFECTIVE_DATE] = "1900-01-01";
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_NAME] = grade_profile_name;
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_COMP_ELEMENT] = grade_profile[INDEX_COMP_GRADE_PROFILE_COMP_ELEMENT];
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_ELIGIBILITY_RULE] = grade_profile_name;
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_INACTIVE] = "N";
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_NUM_SEGMENTS] = 3;
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_MIN] = grade_profile[INDEX_COMP_GRADE_PROFILE_MIN];
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_MID] = grade_profile[INDEX_COMP_GRADE_PROFILE_MID];
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_MAX] = grade_profile[INDEX_COMP_GRADE_PROFILE_MAX];
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_SEGMENT_1_TOP] = grade_profile[INDEX_COMP_GRADE_PROFILE_SEGMENT_1_TOP];
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_SEGMENT_2_TOP] = grade_profile[INDEX_COMP_GRADE_PROFILE_SEGMENT_2_TOP];
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_CURRENCY] = grade_profile[INDEX_COMP_GRADE_PROFILE_CURRENCY];
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_FREQUENCY] = grade_profile[INDEX_COMP_GRADE_PROFILE_FREQUENCY];
			eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_ALLOW_OVERRIDE] = "N";
			if (grade_profile[INDEX_COMP_GRADE_PROFILE_COMP_BASIS] === "Total Cost to Company") {
				eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_MIN] = 0;
				eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_MID] = 0;
				eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_MAX] = 0;
				eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_SEGMENT_1_TOP] = 0;
				eib[row_eib + i][INDEX_EIB_COMP_GRADE_PROFILE_SEGMENT_2_TOP] = 0;
				eib[row_eib + i][INDEX_EIB_COMP_BASIS_PROFILE_ROW_ID] = i + 1;
				eib[row_eib + i][INDEX_EIB_COMP_BASIS_PROFILE_DELETE] = "N";
				eib[row_eib + i][INDEX_EIB_COMP_BASIS_PROFILE_COMP_BASIS] = grade_profile[INDEX_COMP_GRADE_PROFILE_COMP_BASIS];
				eib[row_eib + i][INDEX_EIB_COMP_BASIS_PROFILE_MIN] = grade_profile[INDEX_COMP_GRADE_PROFILE_MIN];
				eib[row_eib + i][INDEX_EIB_COMP_BASIS_PROFILE_MID] = grade_profile[INDEX_COMP_GRADE_PROFILE_MID];
				eib[row_eib + i][INDEX_EIB_COMP_BASIS_PROFILE_MAX] = grade_profile[INDEX_COMP_GRADE_PROFILE_MAX];
				eib[row_eib + i][INDEX_EIB_COMP_BASIS_PROFILE_SEGMENT_1_TOP] = grade_profile[INDEX_COMP_GRADE_PROFILE_SEGMENT_1_TOP];
				eib[row_eib + i][INDEX_EIB_COMP_BASIS_PROFILE_SEGMENT_2_TOP] = grade_profile[INDEX_COMP_GRADE_PROFILE_SEGMENT_2_TOP];
			}
		}
		row_eib = row_eib + NUM_GRADE_PROFILES;
	}

	// get empty eib sheet
	var sheet_eib = SpreadsheetApp.openById(GOOGLE_ID_SS_COMP_GRADES).getSheetByName("Compensation Grade");
	sheet_eib.getRange(38, 1, NUM_EIB_ROWS, NUM_EIB_COLS).setValues(eib);
}