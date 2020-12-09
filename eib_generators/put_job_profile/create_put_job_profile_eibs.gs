var EIB_EFFECTIVE_DATE = "2020-12-09";
var FIRST_COL_EXTRACTED = 1; // VZ job code column
var LAST_COL_EXTRACTED = 22; // VZ union job column
var FIRST_ROW_EXTRACTED = 7; // first VZ job
var LAST_ROW_EXTRACTED = 1007; // last VZ job
var NUM_ROWS_TO_EXTRACT = LAST_ROW_EXTRACTED - FIRST_ROW_EXTRACTED + 1;
var NUM_COLS_TO_EXTRACT = LAST_COL_EXTRACTED - FIRST_COL_EXTRACTED + 1;
var FIRST_ROW_EIB = 6;
var FIRST_COL_EIB = 1;
var LAST_COL_EIB = 94;
var NUM_EIB_COLS = LAST_COL_EIB - FIRST_COL_EIB + 1; // compensation grade is last column we use -- ignore further columns
var NUM_EIB_ROWS_PER_JOB = 30;
var NUM_EIB_ROWS = NUM_ROWS_TO_EXTRACT * NUM_EIB_ROWS_PER_JOB;

// put_job_profile eib template array indices (column index minus one)
var FIELDS = 1 - 1;
var SPREADSHEET_KEY = 2 - 1;
var ADD_ONLY = 3 - 1;
var JOB_CODE = 5 - 1;
var EFFECTIVE_DATE = 6 - 1;
var INACTIVE = 7 - 1;
var JOB_TITLE = 8 - 1;
var INCLUDE_JOB_CODE_IN_NAME = 9 - 1;
var JOB_PROFILE_PRIVATE_TITLE = 10 - 1;
var WORK_SHIFT_REQUIRED = 14 - 1;
var PUBLIC_JOB = 15 - 1;
var MANAGEMENT_LEVEL = 16 - 1;
var JOB_CATEGORY = 17 - 1;
var JOB_LEVEL = 18 - 1;
var ROW_ID_JOB_FAMILY = 19 - 1;
var DELETE_JOB_FAMILY = 20 - 1;
var JOB_FAMILY = 21 - 1;
var RESTRICT_TO_COUNTRY = 28 - 1;
var ROW_ID_JOB_CLASSIFICATION = 29 - 1;
var DELETE_JOB_CLASSIFICATION = 30 - 1;
var JOB_CLASSIFICATIONS = 31 - 1;
var ROW_ID_PAY_RATE_TYPE = 32 - 1;
var DELETE_PAY_RATE_TYPE = 33 - 1;
var PAY_RATE_TYPE_COUNTRY = 34 - 1;
var PAY_RATE_TYPE = 35 - 1;
var ROW_ID_JOB_EXEMPT = 36 - 1;
var DELETE_JOB_EXEMPT = 37 - 1;
var JOB_EXEMPT_LOCATION_CONTEXT = 38 - 1;
var JOB_EXEMPT = 40 - 1;
var ROW_ID_WORKERS_COMPENSATION_CODE = 41 - 1;
var WORKERS_COMPENSATION_CODE = 42 - 1;
var COMPENSATION_GRADE = 94 - 1;



// spreadsheet template fields and array indices
var GOOGLE_ID_SS_VZ_JOB_PROFILES = "13yag7K9-IgPjKkEdl4EYFiHsXDiW3uqHPvRwq0cCiuE";
var SHEET_NAME_VZ_JOB_PROFILES = "IMPORT_Job_Profiles";
var SHEET_NAME_EIB = "Job Profile";
var VZ_JOB_CODE = 2 - 1;
var VZ_JOB_PROFILE_NAME = 3 - 1;
var VZ_JOB_TITLE_DEFAULT = 4 - 1;
var VZ_JOB_FAMILY_GROUP = 5 - 1;
var VZ_JOB_FAMILY = 6 - 1;
var VZ_JOB_CATEGORY = 8 - 1;
var VZ_JOB_LEVEL = 10 - 1;
var VZ_FLSA_STATUS = 11 - 1;
var VZ_PAY_RATE_TYPE = 12 - 1;
var VZ_COMPENSATION_GRADE = 15 - 1;
var VZ_MANAGEMENT_LEVEL = 16 - 1;
var VZ_JOB_CLASSIFICATION_AAP = 17 - 1;
var VZ_JOB_CLASSIFICATION_EEO = 18 - 1;
var VZ_WORKERS_COMPENSATION_CODE = 19 - 1;
var VZ_WORK_SHIFT_REQUIRED = 20 - 1;
var VZ_PUBLIC_JOB = 21 - 1;
var VZ_UNION_JOB = 22 - 1;

var pay_rate_type_countries_non_usa = ["AU","BE","BR","CA","CN","DK","FR","DE","HK","IN","ID","IE","IL","IT","JP","KR","LU","NL","NZ","NO","RO","SG","ES","SE","CH","TW","TH","GB","VN"];

function create_put_job_profile_eibs()
{
	// load sheet with VZ job profiles
	var sheet_jobs = SpreadsheetApp.openById(GOOGLE_ID_SS_VZ_JOB_PROFILES).getSheetByName(SHEET_NAME_VZ_JOB_PROFILES);
	// extract VZ job profiles
	var values_jobs = sheet_jobs.getRange(FIRST_ROW_EXTRACTED, FIRST_COL_EXTRACTED, NUM_ROWS_TO_EXTRACT, NUM_COLS_TO_EXTRACT).getValues();
	var NUM_JOBS = values_jobs.length;

	var eib = new Array(NUM_EIB_ROWS);
	for (var i = 0; i < NUM_EIB_ROWS; i++)
	{
		eib[i] = new Array(NUM_EIB_COLS);
	}

	var row_eib = 0;
	for (var row = 0; row < NUM_JOBS; row++)
	{
		// extract current job
		var curr = values_jobs[row];
		row_eib = NUM_EIB_ROWS_PER_JOB * row;

		// multi spreadsheet key
		curr_spreadsheet_key = row+1;
		eib[row_eib][SPREADSHEET_KEY] = curr_spreadsheet_key;
		eib[row_eib][ADD_ONLY] = "Y";
		eib[row_eib][JOB_CODE] = curr[VZ_JOB_CODE];
		eib[row_eib][EFFECTIVE_DATE] = EIB_EFFECTIVE_DATE;
		eib[row_eib][INACTIVE] = "Y";
		eib[row_eib][JOB_TITLE] = curr[VZ_JOB_PROFILE_NAME];
		eib[row_eib][INCLUDE_JOB_CODE_IN_NAME] = "N";
		eib[row_eib][JOB_PROFILE_PRIVATE_TITLE] = curr[VZ_JOB_TITLE_DEFAULT];
		// single work shift required
		eib[row_eib][PUBLIC_JOB] = "Y";
		eib[row_eib][MANAGEMENT_LEVEL] = curr[VZ_MANAGEMENT_LEVEL];
		eib[row_eib][JOB_CATEGORY] = curr[VZ_JOB_CATEGORY];
		eib[row_eib][JOB_LEVEL] = curr[VZ_JOB_LEVEL];
		eib[row_eib][ROW_ID_JOB_FAMILY] = 1;
		eib[row_eib][DELETE_JOB_FAMILY] = "N";
		eib[row_eib][JOB_FAMILY] = curr[VZ_JOB_FAMILY];
		// multi restrict to country
		eib[row_eib][ROW_ID_JOB_CLASSIFICATION] = 1;
		eib[row_eib][DELETE_JOB_CLASSIFICATION] = "N";
		eib[row_eib][JOB_CLASSIFICATIONS] = curr[VZ_JOB_CLASSIFICATION_EEO];
		eib[row_eib+1][ROW_ID_JOB_CLASSIFICATION] = 2; // add 2nd job classification
		eib[row_eib+1][DELETE_JOB_CLASSIFICATION] = "N"; // add 2nd job classification
		eib[row_eib+1][JOB_CLASSIFICATIONS] = curr[VZ_JOB_CLASSIFICATION_AAP]; // add 2nd job classification
		eib[row_eib][ROW_ID_PAY_RATE_TYPE] = 1;
		eib[row_eib][DELETE_PAY_RATE_TYPE] = "N";
		eib[row_eib][PAY_RATE_TYPE_COUNTRY] = "US";
		eib[row_eib][PAY_RATE_TYPE] = curr[VZ_PAY_RATE_TYPE];
		// multi workers compensation code
		eib[row_eib][ROW_ID_JOB_EXEMPT] = "1";
		eib[row_eib][DELETE_JOB_EXEMPT] = "N";
		eib[row_eib][JOB_EXEMPT_LOCATION_CONTEXT] = "US";
		eib[row_eib][JOB_EXEMPT] = curr[VZ_FLSA_STATUS];
		eib[row_eib][COMPENSATION_GRADE] = curr[VZ_COMPENSATION_GRADE];
		// fill in extra repetitive rows
		for (var j = 0; j < (NUM_EIB_ROWS_PER_JOB-1); j++)
		{
			row_eib++;
			eib[row_eib][SPREADSHEET_KEY] = curr_spreadsheet_key; // fill in spreadsheet key column
			eib[row_eib][ROW_ID_PAY_RATE_TYPE] = j+2;
			eib[row_eib][DELETE_PAY_RATE_TYPE] = "N";
			eib[row_eib][PAY_RATE_TYPE_COUNTRY] = pay_rate_type_countries_non_usa[j];
			eib[row_eib][PAY_RATE_TYPE] = "Salaried";
		}
	}

	// get empty eib sheet
	var sheet_eib = SpreadsheetApp.openById(GOOGLE_ID_SS_VZ_JOB_PROFILES).getSheetByName(SHEET_NAME_EIB);
	sheet_eib.getRange(FIRST_ROW_EIB, FIRST_COL_EIB, NUM_EIB_ROWS, NUM_EIB_COLS).setValues(eib);
	// SpreadsheetApp.flush();
}

create_put_job_profile_eibs();

