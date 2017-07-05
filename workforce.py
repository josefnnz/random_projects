import pandas
import numpy

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
oath_jobs.loc[:,'oath_job_code'] = oath_jobs.loc[:,'oath_job_code'].astype('object')
oath_jobs.drop_duplicates('oath_job_code', inplace=True)

# load city to Oath comp grade profile
oath_cgps = mappings.parse("CompGradeProfiles")
oath_cgps.columns = ['country','city_and_state','city','state','oath_georegion']
oath_cgps.drop_duplicates('city', inplace=True)

# load country to region
regions = mappings.parse('Regions')
regions.columns = ['country', 'region']
regions.drop_duplicates('country', inplace=True)

# # load AOL to Yahoo email
# emailremap = mappings.parse("EmailRemap")
# emailremap.columns = ['emp_name','aol_work_email','yahoo_work_email']
# emailremap.drop_duplicates('aol_work_email', inplace=True)
################################################################################

# change name fields format from "Last, First" to "First Last"
for x in ['legal_name','mgr_legal_name','CEO_name','L2_name','L3_name','L4_name','L5_name','L6_name','L7_name','L8_name','L9_name','L10_name']:
	oath[x] = [" ".join(reversed(w.split(", "))) for w in oath[x]]

# reformat WFH flag
is_wfh = oath['wfh_flag'].str.contains("Yes",case=False)
oath.loc[is_wfh, 'wfh_flag'] = 'WFH'
oath.loc[~is_wfh, 'wfh_flag'] = None

# create acquired company field (AOL/Yahoo)
is_yahoo = oath['company'].str.contains("yahoo",case=False)
oath.loc[is_yahoo, 'acquired_company'] = 'Yahoo'
oath.loc[~is_yahoo, 'acquired_company'] = 'AOL'

# reformat Full time / Part time field
oath.loc[oath['ft_or_pt'].str.contains("Full-Time",case=False), 'ft_or_pt'] = 'Full time'
oath.loc[oath['ft_or_pt'].str.contains("Part-Time",case=False), 'ft_or_pt'] = 'Part time'

# merge in Workday office names
oath['work_office'] = vlookup(oath, offices, 'work_office', 'ps_office_name', 'wd_office_name')
oath = pandas.merge(oath, offices, how='left', left_on='work_office', right_on='ps_office_name')

# merge in Workday region names
oath['work_region'] = vlookup(oath, regions, 'work_country', 'country', 'region')

oath['active_status'] = 'Yes'
oath['emp_type'] = 'placeholder'
oath['job_family_group'] = 'placeholder'
oath['mgmt_level'] = 'placeholder'
oath['pay_rate_type'] = 'placeholder'
oath['L2_org_name'] = 'placeholder'
oath['L3_org_name'] = 'placeholder'
oath['target_bonus_amt_local'] = 'placeholder'
oath['ttc_annualized_local'] = 'placeholder'

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


ced_nonsens_cols = ['eeid','legal_name','mgr_eeid','mgr_legal_name','mgr_email','userid','last_hire_date',\
                    'original_hire_date','active_status','emp_type','ft_or_pt','fte_pct','email',\
                    'acquired_company','job_code','job_profile','job_family_group','job_family',\
                    'job_level','job_category','mgmt_level','comp_grade_profile','pay_rate_type',\
                    'work_office','wfh_flag','work_country','work_region','CEO_name','L2_name',\
                    'L3_name','L4_name','L5_name','L6_name','L7_name','L8_name','L9_name']

ced_nonsens = oath.loc[:, ced_nonsens_cols]

cks_cols = ['eeid','legal_name','mgr_eeid','mgr_legal_name','mgr_email','userid','last_hire_date',\
            'original_hire_date','active_status','emp_type','ft_or_pt','fte_pct','std_hrs','email',\
            'acquired_company','job_code','job_profile','job_family_group','job_family','job_level',\
            'job_category','mgmt_level','comp_grade','comp_grade_profile','pay_rate_type','flsa',\
            'currency_code','base_annualized_local','base_annualized_usd','bonus_plan',\
            'target_bonus_pct','target_bonus_amt_local','target_bonus_amt_usd','ttc_annualized_local',\
            'ttc_annualized_usd','wfh_flag','work_office',\
            'work_city','work_state','work_country','work_region','CEO_name','L2_name','L3_name',\
            'L4_name','L5_name','L6_name','L7_name','L8_name','L9_name','L2_org_name','L3_org_name']

cks = oath.loc[:, cks_cols]

# # update Yahoos with AOL email address with their current Yahoo email
# # oath.replace(emailremap.set_index('aol_work_email').to_dict()['yahoo_work_email'], inplace=True)
# oath.loc[:,'email'] = vlookup(oath, ycomp, 'yahoo_eeid', 'eeid', 'email')


# # remove employees either (1) offboarded, (2) on transition, (3) future term
# # oath = oath[~oath['work_email'].isin(offboards['work_email'])]
# oath = oath[oath['separations_date'].isnull()]
# #oath = oath[(~oath['company'].str.contains("yahoo",case=False)) | (oath['work_email'].isin(yactive))] # remove inactive Yahoos -- keep all Aolers

# # merge in Yahoo job profiles for Yahoos, and then merge in oath job details
# # oath = pandas.merge(oath, ycomp[['work_email','job_profile']], how='left', on='work_email')
# # oath.loc[oath['job_profile_y'].notnull(), 'job_profile_x'] = oath['job_profile_y']
# # oath = pandas.merge(oath, jobs, how='left', left_on='job_profile', right_on='old_job_profile')
# # oath = pandas.merge(oath, jobs, how='left', on='job_code')
# oath.update(jobs, on='job_code')

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


# # extract relevant fields
# final_report = oath[['eeid','legal_name','work_email','userid','status','last_hire_date','regular_or_temp','full_time_or_part_time',\
#                         'acquired_company','mgr_eeid','mgr_legal_name','mgr_work_email','mgr_userid','fte_pct','job_code','job_profile_y',\
#                         'job_family_grp','job_family_y','job_category','job_level_y','mgmt_level_y','eeo_job_classification_y',\
#                         'aap_job_classification','comp_frequency','pay_rate_type','standard_hours','flsa','local_currency_x','fx_rate',\
#                         'base_annualized_local','target_abp_pct','target_abp_exception_flag','sales_incentive_guarantee','wd_office_name',\
#                         'wfh_flag','work_city','work_state_code','work_postal_code','work_country','comp_grade_y','georegion','span','CEO',\
#                         'L2','L3','L4','L5','L6','L7','L8','L9','L10','L2_orgname','L3_orgname','gender','ethnicity','marital_status','military_status','benefit_program']]

# final_report.columns = ['Employee ID','Legal Name','Work Email','User ID','Status','Last Hire Date','Regular / Temp','FT / PT','Acquired Company',\
#                         'Direct Supervisor - Emp ID','Direct Supervisor - Legal Name','Direct Supervisor - Work Email','Direct Supervisor - User ID',\
#                         'FTE %','Job Code','Job Profile','Job Family Group','Job Family','Job Category','Job Level','Management Level','EEO Job Group',\
#                         'EEO Job Family','Compensation Frequency','Pay Rate Type','Standard Hours','FLSA','Local Currency','FX Rate','Base Annualized (Local)',\
#                         'Target Bonus %','Target ABP Exception','Sales Incentive Guarantee','Work Location - Office','Work Location - Workspace',\
#                         'Work Location - City','Work Location - State','Work Location - Postal Code','Work Location - Country','Comp Grade','Comp Grade Profile',\
#                         'Direct Headcount','CEO','L2','L3','L4','L5','L6','L7','L8','L9','L10','L2 Org Name','L3 Org Name','gender','ethnicity','marital_status','military_status','benefit_program']

# writer = pandas.ExcelWriter('output.xlsx')
# final_report.to_excel(writer,'Sheet1', index=False)
# writer.save()
