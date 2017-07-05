import pandas
import numpy

# FOR REFERENCE: counts by group is df.groupby('colname').size()

################################################################################
##### Setup options

pandas.set_option('display.max_columns', None)
pandas.set_option('display.max_rows', None)

################################################################################
##### Helper functions

def vlookup(left, right, left_key, right_key, right_col):
    mleft = left.loc[:, left_key].to_frame()
    mright = right.loc[:, [right_key, right_col]]
    result = pandas.merge(mleft, mright, how='left', left_on=left_key, right_on=right_key)
    return result.loc[:, right_col].to_frame()

def vlookup_update(left, right, left_key, right_key, left_col, right_col):
    mleft = left.loc[:, [left_key, left_col]]
    mright = right.loc[:, [right_key, right_col]]
    mleft.columns, mright.columns = ['key','lval'], ['key','rval']
    output = pandas.merge(mleft, mright, how='left', on='key')
    output.loc[output.loc[:, 'rval'].notnull(), 'lval'] = output.loc[:, 'rval']
    return output.loc[:, 'lval'].to_frame()

################################################################################

setwd = "/Users/josefnunez/workforce/"

# load Oath workforce roster, i.e. PeopleSoft report
oath_filepath, oath_sheet = setwd+"roster.xlsx", "Sheet1"
oath = pandas.ExcelFile(oath_filepath).parse(oath_sheet)
# oath_filepath = setwd+"roster.csv"
# oath = pandas.read_csv(oath_filepath, low_memory=False)
oath.columns = ['wf_group','worker_type','yahoo_eeid','yahoo_userid','aol_eeid',\
                'legal_name','CEO_name','L2_eeid','L2_name','L3_eeid','L3_name',\
                'L4_eeid','L4_name','L5_eeid','L5_name','L6_eeid','L6_name',\
                'L7_eeid','L7_name','L8_name','L9_name','L10_name','comp_grade',\
                'comp_grade_profile','headcount_group','status','regular_or_temp',\
                'ft_or_pt','gender','ethnicity','marital_status','military_status',\
                'email','userid','mgr_eeid','mgr_legal_name','mgr_email','manager_userid',\
                'acquired_company','contract_type','contractor_type',\
                'contract_number','contract_status','msp_or_nonmsp',\
                'contract_start_date','contract_end_date','contract_provider_id',\
                'contract_provider','birth_date','original_hire_date','last_hire_date',\
                'department_entry_date','comp_grade_entry_date','job_entry_date',\
                'benefit_program','base_annualized_local','currency_code',\
                'base_annualized_usd','target_abp_plan_yr','target_abp_pct',\
                'target_abp_exception_flag','target_abp_local','target_abp_usd',\
                'sip_year','sales_begin_date','sales_end_date','sip_target_local',\
                'sip_target_usd','sip_guarantee','wfh_flag','work_country_code',\
                'work_location_code','work_office','work_country','work_state',\
                'work_city','work_postal_code','home_country','home_state',\
                'home_ciity','home_postal_code','company_code','company',\
                'market_cluster','job_code','job_profile','job_family',\
                'aap_job_classification','eeo_job_classification','comp_freq',\
                'std_hrs','fte_pct','flsa','department_code','department_name',\
                'business_unit','division','region_code','reporting_schema_level_1',\
                'reporting_schema_level_2','hr_support_eeid','hr_support_name',\
                'hr_support_userid','separations_group','separation_date','layer',\
                'userid_hierarchy','direct_headcount']		
oath['job_code'] = oath['job_code'].apply('{0:0>6}'.format)	# reformat job code for lookup

# load Yahoo comp
ycomp_filepath, ycomp_sheet, ycomp_skiprows = setwd+"yahoo_comp.xlsx", "Sheet1", 2
ycomp = pandas.ExcelFile(ycomp_filepath).parse(ycomp_sheet, skiprows=ycomp_skiprows)
ycomp.columns = ['yahoo_eeid','emp_preferred_name','email','emp_type','yahoo_job_code',\
                 'yahoo_job_profile','yahoo_job_family_group','yahoo_job_family','yahoo_job_level',\
                 'yahoo_job_category','yahoo_comp_grade','yahoo_comp_grade_profile','local_currency',\
                 'base_annualized_in_local','base_annualized_in_usd','fx_rate',\
                 'yahoo_bonus_plan','target_bonus_pct','last_day_of_work']

# load Yahoo active workers
yactive_filepath, yactive_sheet, yactive_skiprows = setwd+"yahoo_active_workers.xlsx", "Sheet1", 1
yactive = pandas.ExcelFile(yactive_filepath).parse(yactive_sheet, skiprows=yactive_skiprows)
yactive.columns = ['yahoo_eeid','email','yahoo_userid','worker_type','emp_type']

# load final offboards list from AlixPartners
# offboards = pandas.ExcelFile("/Users/josefnunez/workforce/offboards.xlsx").parse("Sheet1")
# offboards.columns = ['work_email','badge_id','company','last_day_of_work']

################################################################################
##### Load Mappings Tables

# Excel workbook containing tabs of mapping tables
mappings = pandas.ExcelFile("/Users/josefnunez/workforce/mappings.xlsx")

# load PeopleSoft to Workday office names
offices = mappings.parse("Offices")
offices.columns = ['ps_office_name','wd_office_name','wfh_flag']
offices.drop_duplicates('ps_office_name', inplace=True)

# load L2 legal name to L2 org name
l2orgnames = mappings.parse("L2OrgNames")
l2orgnames.columns = ['L2_or_L3','ps_L2_name','wd_L2_name','ps_L3_name','wd_L3_name','orgname']
l2orgnames.drop_duplicates('ps_L2_name', inplace=True)

# load L3 legal name to L3 org name
l3orgnames = mappings.parse("L3OrgNames")
l3orgnames.columns = ['L2_or_L3','ps_L2_name','wd_L2_name','ps_L3_name','wd_L3_name','orgname']
l3orgnames.drop_duplicates('ps_L3_name', inplace=True)

# load Oath job catalog
oath_jobs = mappings.parse("Jobs")
oath_jobs.columns = ['oath_job_code','oath_job_profile','oath_job_family_group','oath_job_family',\
                     'oath_job_category_sort_order','oath_job_category','oath_job_level',\
                     'oath_mgmt_level','oath_eeo_job_classification','oath_aap_job_classification',\
                     'oath_pay_rate_type','oath_is_exempt','oath_comp_grade']
oath_jobs['oath_job_code'] = oath_jobs['oath_job_code'].apply('{0:0>6}'.format) # reformat job code for lookup
oath_jobs.drop_duplicates('oath_job_code', inplace=True)

# load city to Oath comp grade profile
oath_cgps = mappings.parse("CompGradeProfiles")
oath_cgps.columns = ['country','city_and_state','city','state','oath_georegion']
oath_cgps.drop_duplicates('city', inplace=True)

# load country to region
regions = mappings.parse('Regions')
regions.columns = ['country', 'region']
regions.drop_duplicates('country', inplace=True)

################################################################################

# immediately remove workers marked as not used for headcount reporting
oath = oath.loc[~oath["wf_group"].str.contains("Not Used for WF Report",case=False)]

# create acquired company field (AOL/Yahoo) and merge AOL/Yahoo numeric eeids into one column
is_yahoo = oath['company'].str.contains("yahoo",case=False)
oath.loc[is_yahoo, 'acquired_company'] = 'Yahoo'
oath.loc[~is_yahoo, 'acquired_company'] = 'AOL'
oath['eeid'] = None
is_yahoo_with_workday_id = (is_yahoo) & (oath['yahoo_eeid'].notnull())
oath.loc[is_yahoo_with_workday_id, 'eeid'] = oath['yahoo_eeid']
oath.loc[oath['eeid'].isnull(), 'eeid'] = oath['aol_eeid']
oath['eeid'] = oath['eeid'].apply('{0:0>6}'.format)
oath.loc[is_yahoo, 'eeid'] = 'Y' + oath['eeid']
oath.loc[~is_yahoo, 'eeid'] = 'A' + oath['eeid']

# change name fields format from "Last, First" to "First Last"
for x in ['legal_name','mgr_legal_name','CEO_name','L2_name','L3_name','L4_name','L5_name','L6_name','L7_name','L8_name','L9_name','L10_name']:
    oath[x] = [" ".join(reversed(w.split(", "))) for w in oath[x]]

# reformat WFH flag
is_wfh = oath['wfh_flag'].str.contains("Yes",case=False)
oath.loc[is_wfh, 'wfh_flag'] = 'WFH'
oath.loc[~is_wfh, 'wfh_flag'] = None

# reformat Full time / Part time field
oath.loc[oath['ft_or_pt'].str.contains("Full-Time",case=False), 'ft_or_pt'] = 'Full time'
oath.loc[oath['ft_or_pt'].str.contains("Part-Time",case=False), 'ft_or_pt'] = 'Part time'

# merge in Workday office names
oath['work_office'] = vlookup(oath, offices, 'work_office', 'ps_office_name', 'wd_office_name')
oath = pandas.merge(oath, offices, how='left', left_on='work_office', right_on='ps_office_name')

# merge in Workday region names
oath['work_region'] = vlookup(oath, regions, 'work_country', 'country', 'region')

oath['active_status'] = 'Yes'
oath['L2_org_name'] = 'placeholder'
oath['L3_org_name'] = 'placeholder'

# merge in Oath job details
oath['job_profile'] = vlookup_update(oath, oath_jobs, 'job_code', 'oath_job_code', 'job_profile', 'oath_job_profile')
oath['job_family_group'] = vlookup(oath, oath_jobs, 'job_code', 'oath_job_code', 'oath_job_family_group')
oath['job_family'] = vlookup_update(oath, oath_jobs, 'job_code', 'oath_job_code', 'job_family', 'oath_job_family')
oath['job_level'] = vlookup(oath, oath_jobs, 'job_code', 'oath_job_code', 'oath_job_level')
oath['job_category'] = vlookup(oath, oath_jobs, 'job_code', 'oath_job_code', 'oath_job_category')
oath['mgmt_level'] = vlookup(oath, oath_jobs, 'job_code', 'oath_job_code', 'oath_mgmt_level')
oath['comp_grade'] = vlookup_update(oath, oath_jobs, 'job_code', 'oath_job_code', 'comp_grade', 'oath_comp_grade')
oath['pay_rate_type'] = vlookup(oath, oath_jobs, 'job_code', 'oath_job_code', 'oath_pay_rate_type')
oath['flsa'] = vlookup_update(oath, oath_jobs, 'job_code', 'oath_job_code', 'flsa', 'oath_is_exempt')

# merge Yahoo comp details
oath['base_annualized_local'] = vlookup_update(oath, ycomp, 'yahoo_eeid', 'yahoo_eeid', 'base_annualized_local', 'base_annualized_in_local')
oath['base_annualized_usd'] = vlookup_update(oath, ycomp, 'yahoo_eeid', 'yahoo_eeid', 'base_annualized_usd', 'base_annualized_in_usd')
oath['currency_code'] = vlookup_update(oath, ycomp, 'yahoo_eeid', 'yahoo_eeid', 'currency_code', 'local_currency')
oath['bonus_plan'] = vlookup(oath, ycomp, 'yahoo_eeid', 'yahoo_eeid', 'yahoo_bonus_plan')
oath['target_bonus_pct'] = vlookup(oath, ycomp, 'yahoo_eeid', 'yahoo_eeid', 'target_bonus_pct')

# compute target bonus amount and target TTC
oath['target_bonus_amt_local'] = oath['base_annualized_local'] * oath['target_bonus_pct']
oath['target_bonus_amt_usd'] = oath['base_annualized_usd'] * oath['target_bonus_pct']
oath['ttc_annualized_local'] = oath['base_annualized_local'] + oath['target_bonus_amt_local']
oath['ttc_annualized_usd'] = oath['base_annualized_usd'] + oath['target_bonus_amt_usd']

# complete worker and employee types -- NEEDS TO BE DONE AFTER JOB DETAILS HAVE BEEN MERGED INTO OATH DATA FRAME
oath['worker_type'], oath['emp_type'] = None, None
is_employee = oath['wf_group'].str.contains("Employees for WF Report",case=False)
oath.loc[is_employee, 'worker_type'] = 'Employee'
oath.loc[~is_employee, 'worker_type'] = 'Contingent Worker'
is_intern = oath['job_category'].str.contains('INT',case=False).fillna(False)
oath.loc[is_intern, 'emp_type'] = 'Employee Type - Intern'
oath.loc[(is_employee) & (~is_intern), 'emp_type'] = 'Employee Type - Regular'

# Current Worker Details columns (includes contingent workers)
cwd_nonsens_cols = ['worker_type','emp_type','eeid','legal_name','mgr_eeid','mgr_legal_name','mgr_email','userid','last_hire_date',\
                    'original_hire_date','active_status','ft_or_pt','fte_pct','email',\
                    'acquired_company','job_code','job_profile','job_family_group','job_family',\
                    'job_level','job_category','mgmt_level','comp_grade_profile','pay_rate_type',\
                    'work_office','wfh_flag','work_country','work_region','CEO_name','L2_name',\
                    'L3_name','L4_name','L5_name','L6_name','L7_name','L8_name','L9_name']

cwd_nonsens = oath.loc[:, cwd_nonsens_cols]

# Comp Kitchen Sink columns (includes contingent workers)
cks_cols = ['worker_type','emp_type','eeid','legal_name','mgr_eeid','mgr_legal_name','mgr_email','userid','last_hire_date',\
            'original_hire_date','active_status','ft_or_pt','fte_pct','std_hrs','email',\
            'acquired_company','job_code','job_profile','job_family_group','job_family','job_level',\
            'job_category','mgmt_level','comp_grade','comp_grade_profile','pay_rate_type','flsa',\
            'currency_code','base_annualized_local','base_annualized_usd','bonus_plan',\
            'target_bonus_pct','target_bonus_amt_local','target_bonus_amt_usd','ttc_annualized_local',\
            'ttc_annualized_usd','wfh_flag','work_office',\
            'work_city','work_state','work_country','work_region','CEO_name','L2_name','L3_name',\
            'L4_name','L5_name','L6_name','L7_name','L8_name','L9_name','L2_org_name','L3_org_name']

cks = oath.loc[:, cks_cols]

# # remove employees either (1) offboarded, (2) on transition, (3) future term
# # oath = oath[~oath['work_email'].isin(offboards['work_email'])]
# oath = oath[oath['separations_date'].isnull()]
# #oath = oath[(~oath['company'].str.contains("yahoo",case=False)) | (oath['work_email'].isin(yactive))] # remove inactive Yahoos -- keep all Aolers

# # merge in Oath geo regions
# oath['work_country'].replace({"Korea, Republic of" : "Republic of Korea"}, inplace=True) # reformat Republic of Korea
# oath = pandas.merge(oath, georegions[['city','georegion']], how='left', left_on='work_city', right_on='city')
# oath.loc[~oath['work_country'].str.contains("united states of america",case=False), 'georegion'] = oath['work_country']
# oath.loc[oath['wfh_flag'].notnull(), 'georegion'] = "WFH"

# oath = pandas.merge(oath, ycomp[['work_email','comp_grade','local_currency','base_annualized_in_local','target_bonus_pct','bonus_plan','fx_rate']], how='left', on='work_email')
# oath.loc[oath['base_annualized_in_local'].notnull(), 'base_annualized_local'] = oath['base_annualized_in_local'] # put Yahoo base values in main column
# oath.loc[oath['target_bonus_pct'].notnull(), 'target_abp_pct'] = oath['target_bonus_pct'] # put Yahoo target bonus % in main column
# oath.loc[oath['local_currency_y'].notnull(), 'local_currency_x'] = oath['local_currency_y'] # put Yahoo local currency in main column
# oath.loc[(oath['sales_beginning_date'].notnull()) & (oath['sales_incentive_target_amt_local'] > 0), 'target_abp_pct'] = oath['sales_incentive_target_amt_local'] / oath['base_annualized_local'] # compute Sales bonus targets

# # merge in L2 and L3 org names
# oath = pandas.merge(oath, l2orgnames.query('L2_or_L3 == "L2"')[['ps_L2_name','orgname']], how='left', left_on='L2', right_on='ps_L2_name')
# oath.rename(columns={'orgname' : 'L2_orgname'}, inplace=True)
# oath = pandas.merge(oath, l3orgnames.query('L2_or_L3 == "L3"')[['ps_L3_name','orgname']], how='left', left_on='L3', right_on='ps_L3_name')
# oath.rename(columns={'orgname' : 'L3_orgname'}, inplace=True)

# writer = pandas.ExcelWriter('output.xlsx')
# final_report.to_excel(writer,'Sheet1', index=False)
# writer.save()
