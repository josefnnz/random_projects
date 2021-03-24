/*
 * ASSUMPTIONS:
 * - Generate EIB to create brand new comp grades and grade profiles.
 * - Each default pay range has all zeroes.
 * - Each pay range has 3 segments.
 * - Each grade profile does not calculate segments.
 * - Each pay range prohibits override.
 * - Comp Grade and associated Comp Grade Profiles use same Effective Date.
 * - Each grade profile uses exactly one eligibility rule.
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
	column_array_indices[column_letters[i]] = i - 1;
}

// column number index for data source sheet
var ARRAY_INDEX_COMP_GRADE_NAME = column_array_indices["B"];
var ARRAY_INDEX_COMP_GRADE_EFFECTIVE_DATE = column_array_indices["E"];
var ARRAY_INDEX_COMP_GRADE_INACTIVE = column_array_indices["F"];
var ARRAY_INDEX_COMP_GRADE_NUM_SEGMENTS = column_array_indices["G"];
var ARRAY_INDEX_COMP_GRADE_COMP_ELEMENT = column_array_indices["I"];
var ARRAY_INDEX_COMP_GRADE_CURRENCY = column_array_indices["J"];
var ARRAY_INDEX_COMP_GRADE_FREQUENCY = column_array_indices["K"];
var ARRAY_INDEX_COMP_GRADE_COMP_BASIS = column_array_indices["L"];
var ARRAY_INDEX_COMP_GRADE_PROFILE_NAME = column_array_indices["U"];
var ARRAY_INDEX_COMP_GRADE_PROFILE_ELIGIBILITY_RULE = column_array_indices["W"];
var ARRAY_INDEX_COMP_GRADE_PROFILE_COMP_ELEMENT = column_array_indices["X"];
var ARRAY_INDEX_COMP_GRADE_PROFILE_CURRENCY = column_array_indices["Y"];
var ARRAY_INDEX_COMP_GRADE_PROFILE_FREQUENCY = column_array_indices["Z"];
var ARRAY_INDEX_COMP_GRADE_PROFILE_COMP_BASIS = column_array_indices["AA"];
var ARRAY_INDEX_COMP_GRADE_PROFILE_MIN = column_array_indices["AC"];
var ARRAY_INDEX_COMP_GRADE_PROFILE_MID = column_array_indices["AD"];
var ARRAY_INDEX_COMP_GRADE_PROFILE_MAX = column_array_indices["AE"];
var ARRAY_INDEX_COMP_GRADE_PROFILE_SEGMENT_1_TOP = column_array_indices["AF"];
var ARRAY_INDEX_COMP_GRADE_PROFILE_SEGMENT_2_TOP = column_array_indices["AG"];

// column number index for required EIB columns
var ARRAY_INDEX_EIB_COMP_GRADE_ADD_ONLY = column_array_indices["B"];
var ARRAY_INDEX_EIB_COMP_GRADE_ID = column_array_indices["E"];
var ARRAY_INDEX_EIB_COMP_GRADE_EFFECTIVE_DATE = column_array_indices["F"];
var ARRAY_INDEX_EIB_COMP_GRADE_NAME = column_array_indices["G"];
var ARRAY_INDEX_EIB_COMP_GRADE_COMP_ELEMENT = column_array_indices["I"];
var ARRAY_INDEX_EIB_COMP_GRADE_NUM_SEGMENTS = column_array_indices["K"];
var ARRAY_INDEX_EIB_COMP_GRADE_MIN = column_array_indices["L"];
var ARRAY_INDEX_EIB_COMP_GRADE_MID = column_array_indices["M"];
var ARRAY_INDEX_EIB_COMP_GRADE_MAX = column_array_indices["N"];
var ARRAY_INDEX_EIB_COMP_GRADE_SEGMENT_1_TOP = column_array_indices["P"];
var ARRAY_INDEX_EIB_COMP_GRADE_SEGMENT_2_TOP = column_array_indices["Q"];
var ARRAY_INDEX_EIB_COMP_GRADE_CURRENCY = column_array_indices["U"];
var ARRAY_INDEX_EIB_COMP_GRADE_FREQUENCY = column_array_indices["V"];
var ARRAY_INDEX_EIB_COMP_GRADE_ALLOW_OVERRIDE = column_array_indices["X"];
var ARRAY_INDEX_EIB_COMP_GRADE_PROFILE_ROW_ID = column_array_indices["AJ"];
var ARRAY_INDEX_EIB_COMP_GRADE_PROFILE_DELETE = column_array_indices["AK"];
var ARRAY_INDEX_EIB_COMP_GRADE_PROFILE_ID = column_array_indices["AM"];
var ARRAY_INDEX_EIB_COMP_GRADE_PROFILE_EFFECTIVE_DATE = column_array_indices["AN"];
var ARRAY_INDEX_EIB_COMP_GRADE_PROFILE_NAME = column_array_indices["AO"];
var ARRAY_INDEX_EIB_COMP_GRADE_PROFILE_COMP_ELEMENT = column_array_indices["AQ"];
var ARRAY_INDEX_EIB_COMP_GRADE_PROFILE_ELIGIBILITY_RULE = column_array_indices["AR"];
var ARRAY_INDEX_EIB_COMP_GRADE_PROFILE_INACTIVE = column_array_indices["AS"];
var ARRAY_INDEX_EIB_COMP_GRADE_PROFILE_NUM_SEGMENTS = column_array_indices["AT"];
var ARRAY_INDEX_EIB_COMP_GRADE_PROFILE_MIN = column_array_indices["AU"];
var ARRAY_INDEX_EIB_COMP_GRADE_PROFILE_MID = column_array_indices["AV"];
var ARRAY_INDEX_EIB_COMP_GRADE_PROFILE_MAX = column_array_indices["AW"];
var ARRAY_INDEX_EIB_COMP_GRADE_PROFILE_SEGMENT_1_TOP = column_array_indices["AY"];
var ARRAY_INDEX_EIB_COMP_GRADE_PROFILE_SEGMENT_2_TOP = column_array_indices["AZ"];
var ARRAY_INDEX_EIB_COMP_GRADE_PROFILE_CURRENCY = column_array_indices["BD"];
var ARRAY_INDEX_EIB_COMP_GRADE_PROFILE_FREQUENCY = column_array_indices["BE"];
var ARRAY_INDEX_EIB_COMP_GRADE_PROFILE_ALLOW_OVERRIDE = column_array_indices["BG"];
var ARRAY_INDEX_EIB_COMP_BASIS_PROFILE_ROW_ID = column_array_indices["BT"];
var ARRAY_INDEX_EIB_COMP_BASIS_PROFILE_DELETE = column_array_indices["BU"];
var ARRAY_INDEX_EIB_COMP_BASIS_PROFILE_COMP_BASIS = column_array_indices["BV"];
var ARRAY_INDEX_EIB_COMP_BASIS_PROFILE_MIN = column_array_indices["BW"];
var ARRAY_INDEX_EIB_COMP_BASIS_PROFILE_MID = column_array_indices["BX"];
var ARRAY_INDEX_EIB_COMP_BASIS_PROFILE_MAX = column_array_indices["BY"];
var ARRAY_INDEX_EIB_COMP_BASIS_PROFILE_SEGMENT_1_TOP = column_array_indices["CA"];
var ARRAY_INDEX_EIB_COMP_BASIS_PROFILE_SEGMENT_2_TOP = column_array_indices["CB"];