import pandas

oath_colnames = ['eeid','legal_name','CEO','L2','L3','L4','L5','L6','L7','L8','L9','L10',\
	        'job_level','comp_grade','mgmt_level','status','regular_or_temp','full_time_or_part_time',\
	        'work_email','userid','mgr_eeid','mgr_name','mgr_work_email','mgr_userid','acquired_company',\
	        'birth_date','original_hire_date','last_hire_date','department_entry_date','job_level_entry_date',\
	        'job_entry_date','base_annualized_local','local_currency','base_annualized_usd','target_abp_plan_yr',\
	        'target_abp_pct','target_abp_exception_flag','target_abp_amt_local','target_abp_amt_usd',\
	        'sales_incentive_plan_yr','sales_beginning_date','sales_end_date','sales_incentive_target_amt_local',\
	        'sales_incentive_target_amt_usd','sales_incentive_guarantee','work_from_home_flag','work_country_code',\
	        'location_code','work_office','work_country','work_state_code','work_city','work_postal_code','home_country',\
	        'home_state_code','home_city','home_postal_code','company_code','company','market_cluster','job_code',\
	        'job_profile','job_family','aol_eeo_group','eeo_job_classification','comp_frequency','standard_hours',\
	        'fte_pct','flsa','department_code','department_name','business_unit','division','region_code',\
	        'reporting_schema_level_1','reporting_schema_level_2','hr_support_eeid','hr_support_name','hr_support_userid',\
	        'separations_group','separations_date','layer','mgr_userid_hierarchy','span']

oath_filepath = "/Users/josefnunez/workforce/roster.xlsx"
oath_sheetname = "Sheet1"

oath = pandas.ExcelFile(oath_filepath).parse(oath_sheetname)
oath.columns = oath_colnames

mappings_filepath = "/Users/josefnunez/workforce/mappings.xlsx"
orgnames_sheetname = "SupOrgNames"
offices_sheetname = "Offices"

mappings = pandas.ExcelFile(mappings_filepath)
orgnames = mappings.parse(orgnames_sheetname)
orgnames.columns = ['L2_or_L3','L2','L3','orgname']
offices = mappings.parse(offices_sheetname)
offices.columns = ['peoplesoft_office_name','worday_office_name','wfh_flag']

oath = pandas.merge(oath, offices, how='left', left_on='work_office', right_on='peoplesoft_office_name')