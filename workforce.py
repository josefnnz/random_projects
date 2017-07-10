import pandas
import numpy
import datetime

# FOR REFERENCE: counts by group is df.groupby('colname').size()

################################################################################
##### Setup options

pandas.set_option('display.max_columns', None)
pandas.set_option('display.max_rows', None)

################################################################################
##### Helper functions

def vlookup(left, right, left_key, right_key, right_col, new_col_name):
    mleft = left.copy()
    mright = right.loc[:, [right_key, right_col]]
    mright.columns = [left_key, 'rval_lookup_col']
    output = pandas.merge(mleft, mright, how='left', on=left_key)
    output.rename(columns={'rval_lookup_col':new_col_name}, inplace=True)
    return output

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
oath.columns = ['wf_group','worker_type','yahoo_eeid','yahoo_userid','aol_eeid',\
                'legal_name','CEO_name','L2_pseeid','L2_name','L3_pseeid','L3_name',\
                'L4_pseeid','L4_name','L5_pseeid','L5_name','L6_pseeid','L6_name',\
                'L7_pseeid','L7_name','L8_name','L9_name','L10_name','comp_grade',\
                'comp_grade_profile','headcount_group','status','regular_or_temp',\
                'ft_or_pt','gender','ethnicity','marital_status','military_status',\
                'email','userid','mgr_pseeid','mgr_legal_name','mgr_email','manager_userid',\
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
                 'yahoo_bonus_plan','target_bonus_pct','last_day_of_work','last_hire_date',\
                 'original_hire_date','yahoo_userid']

# # load Yahoo active workers
# yactive_filepath, yactive_sheet, yactive_skiprows = setwd+"yahoo_active_workers.xlsx", "Sheet1", 1
# yactive = pandas.ExcelFile(yactive_filepath).parse(yactive_sheet, skiprows=yactive_skiprows)
# yactive.columns = ['yahoo_eeid','email','yahoo_userid','worker_type','emp_type']

# load final offboards list from AlixPartners
offboards = pandas.ExcelFile("/Users/josefnunez/workforce/offboards.xlsx").parse("Sheet1")
offboards.columns = ['is_reduction','comment','pending_term_entry','ldw_wd','ldw_wd_ps',\
                     'transition_date','final_ldw','final_term_date','eeid','text_badge_id',\
                     'badge_id','emp_name','company','l2_org_name','alixpartners_transition_date',\
                     'talent_decision']

# load Yahoo promos
ypromos_filepath, ypromos_sheet = setwd+"yahoo_promos.xlsx", "Sheet1"
ypromos = pandas.ExcelFile(ypromos_filepath).parse(ypromos_sheet)
ypromos.columns = ['emp_name', 'eeid', 'company', 'oath_job_code', 'oath_job_profile', 'comment']
ypromos['oath_job_code'] = ypromos['oath_job_code'].apply('{0:0>6}'.format) # reformat job code for lookup

# acomp_filepath, acomp_sheet = setwd+"aol_bonuses.xlsx", "Sheet1"
# acomp = pandas.ExcelFile(acomp_filepath).parse(acomp_sheet)
# acomp.columns = ['eeid','aol_eeid','emp_name','aol_job_code','is_aol_sales_ee','company',\
#                  'comp_freq','local_currency','sales_incentive_guarantee','sales_incentive_plan_yr',\
#                  'sales_incentive_target_amt_local','sales_incentive_target_amt_usd','target_abp_pct',\
#                  'target_abp_amt_local','target_abp_amt_usd','target_abp_exception_flag','target_abp_plan_yr']

# # load exchange rates
# fxrates_filepath, fxrates_sheet, fxrates_skiprows = setwd+"fxrates.xlsx", "Currency Rates", 3
# fxrates = pandas.ExcelFile(fxrates_filepath).parse(fxrates_sheet, skiprows=fxrates_skiprows)
# fxrates.columns = ['source_currency','target_currency','fxrate_type','fxrate_timestamp','fxrate']

################################################################################
##### Load Mappings Tables

# Excel workbook containing tabs of mapping tables
mappings = pandas.ExcelFile("/Users/josefnunez/workforce/mappings.xlsx")

# load PeopleSoft to Workday office names
offices = mappings.parse("Offices")
offices.columns = ['ps_office_name','wd_office_name','wfh_flag']
offices.drop_duplicates('ps_office_name', inplace=True)

# load L2 legal name to L2 org name
orgnames = mappings.parse("OrgNames")
orgnames.columns = ['layer','eeid','leader_name','leader_org_name']
orgnames.drop_duplicates('eeid', inplace=True)

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

# lookup AOL/Yahoo numeric eeids for management chain (CEO -> L10)
oath['CEO_eeid'] = 'A188900'
L2_L7_pseeid_cols = ['mgr_pseeid','L2_pseeid','L3_pseeid','L4_pseeid','L5_pseeid','L6_pseeid','L7_pseeid','L8_name','L9_name','L10_name']
L2_L7_eeid_cols = ['mgr_eeid','L2_eeid','L3_eeid','L4_eeid','L5_eeid','L6_eeid','L7_eeid','L8_eeid','L9_eeid','L10_eeid']
L2_L7_lookup_cols = ['aol_eeid']*7 + ['legal_name']*3
for i in range(len(L2_L7_pseeid_cols)):
    curr_pseeid, curr_eeid, curr_lookup_col = L2_L7_pseeid_cols[i], L2_L7_eeid_cols[i], L2_L7_lookup_cols[i]
    oath[curr_eeid] = vlookup(oath, oath, curr_pseeid, curr_lookup_col, 'eeid')     

# change name fields format from "Last, First" to "First Last"
for x in ['legal_name','mgr_legal_name','CEO_name','L2_name','L3_name','L4_name','L5_name','L6_name','L7_name','L8_name','L9_name','L10_name']:
    oath[x] = [" ".join(reversed(w.split(", "))) for w in oath[x]]

# reformat Full time / Part time field
oath.loc[oath['ft_or_pt'].str.contains("Full-Time",case=False), 'ft_or_pt'] = 'Full time'
oath.loc[oath['ft_or_pt'].str.contains("Part-Time",case=False), 'ft_or_pt'] = 'Part time'

# merge in Workday office names
oath['work_office'] = vlookup_update(oath, offices, 'work_office', 'ps_office_name', 'work_office', 'wd_office_name')
oath['wfh_flag'].replace({'No':None, 'Yes':'WFH'}, inplace=True)

# merge in Workday region names
oath['work_region'] = vlookup(oath, regions, 'work_country', 'country', 'region')

# set active status value for all active workers
oath['active_status'] = 'Yes'

# update Yahoo promos with correct Oath job code -- NEEDS TO BE DONE BEFORE MERGING OATH JOB DETAILS
oath['job_code'] = vlookup_update(oath, ypromos, 'eeid', 'eeid', 'job_code', 'oath_job_code')

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
oath['yahoo_bonus_plan'] = vlookup(oath, ycomp, 'yahoo_eeid', 'yahoo_eeid', 'yahoo_bonus_plan') # for AlixPartners report
oath['target_bonus_pct'] = vlookup(oath, ycomp, 'yahoo_eeid', 'yahoo_eeid', 'target_bonus_pct')
oath['yahoo_target_bonus_pct'] = vlookup(oath, ycomp, 'yahoo_eeid', 'yahoo_eeid', 'target_bonus_pct') # for AlixPartners report
oath['last_hire_date'] = vlookup_update(oath, ycomp, 'yahoo_eeid', 'yahoo_eeid', 'last_hire_date', 'last_hire_date')
oath['original_hire_date'] = vlookup_update(oath, ycomp, 'yahoo_eeid', 'yahoo_eeid', 'original_hire_date', 'original_hire_date')
oath['userid'] = vlookup_update(oath, ycomp, 'yahoo_eeid', 'yahoo_eeid', 'userid', 'yahoo_userid')
# oath['fxrate'] = vlookup(oath, ycomp, 'yahoo_eeid', 'yahoo_eeid', 'fx_rate') # for AlixPartners report

# # merge AOL comp details
# oath['sales_incentive_guarantee'] = vlookup_update(oath, acomp, 'eeid', 'eeid', 'sales_incentive_guarantee', 'sales_incentive_guarantee')
# oath['sales_incentive_plan_yr'] = vlookup(oath, acomp, 'eeid', 'eeid', 'sales_incentive_plan_yr')
# oath['sales_incentive_target_amt_local'] = vlookup_update(oath, acomp, 'eeid', 'eeid', 'sales_incentive_target_amt_local', 'sales_incentive_target_amt_local')
# oath['sales_incentive_target_amt_usd'] = vlookup_update(oath, acomp, 'eeid', 'eeid', 'sales_incentive_target_amt_usd', 'sales_incentive_target_amt_usd')
# oath['target_abp_pct'] = vlookup_update(oath, acomp, 'eeid', 'eeid', 'target_abp_pct', 'target_abp_pct')
# oath['target_abp_amt_local'] = vlookup_update(oath, acomp, 'eeid', 'eeid', 'target_abp_amt_local', 'target_abp_amt_local')
# oath['target_abp_amt_usd'] = vlookup_update(oath, acomp, 'eeid', 'eeid', 'target_abp_amt_usd', 'target_abp_amt_usd')
# oath['target_abp_exception_flag'] = vlookup_update(oath, acomp, 'eeid', 'eeid', 'target_abp_exception_flag', 'target_abp_exception_flag')
# oath['target_abp_plan_yr'] = vlookup_update(oath, acomp, 'eeid', 'eeid', 'target_abp_plan_yr', 'target_abp_plan_yr')

# oath.loc[oath['target_abp_pct']>0, 'target_bonus_pct'] = oath['target_abp_pct']
# oath.loc[oath['target_abp_pct']>0, 'bonus_plan'] = 'AOL Bonus Plan'
# oath.loc[oath['sales_incentive_target_amt_local']>0, 'target_bonus_pct'] = oath['sales_incentive_target_amt_local'] / oath['base_annualized_local']
# oath.loc[oath['sales_incentive_target_amt_local']>0, 'bonus_plan'] = 'AOL Sales Incentive Plan'

# compute target bonus amount and target TTC
oath['target_bonus_amt_local'] = oath['base_annualized_local'].astype(float) * oath['target_bonus_pct'].astype(float)
oath['target_bonus_amt_usd'] = oath['base_annualized_usd'].astype(float) * oath['target_bonus_pct'].astype(float)
oath['ttc_annualized_local'] = oath['base_annualized_local'].astype(float) + oath['target_bonus_amt_local'].astype(float)
oath['ttc_annualized_usd'] = oath['base_annualized_usd'].astype(float) + oath['target_bonus_amt_usd'].astype(float)

# complete worker and employee types -- NEEDS TO BE DONE AFTER JOB DETAILS HAVE BEEN MERGED INTO OATH DATA FRAME
oath['worker_type'], oath['emp_type'] = None, None
is_employee = oath['wf_group'].str.contains("Employees for WF Report",case=False)
oath.loc[is_employee, 'worker_type'] = 'Employee'
oath.loc[~is_employee, 'worker_type'] = 'Contingent Worker'
is_intern = oath['job_category'].str.contains('INT',case=False).fillna(False)
oath.loc[is_intern, 'emp_type'] = 'Employee Type - Intern'
oath.loc[(is_employee) & (~is_intern), 'emp_type'] = 'Employee Type - Regular'

# reformat FLSA field
oath['flsa'].replace({'N':'Non-Exempt', 'Nonexempt':'Non-Exempt', 'Y':'Exempt', 'Exempt':'Exempt'}, inplace=True)

# merge in Last Day of Work and Term Date -- NEEDS TO BE DONE AFTER EEID FORMATTED
oath['last_day_of_work'] = vlookup(oath, offboards, 'eeid', 'eeid', 'final_ldw')
oath['term_date'] = vlookup(oath, offboards, 'eeid', 'eeid', 'final_term_date')

# merge in L2-L4 org names
oath['L2_org_name'] = vlookup(oath, orgnames, 'L2_eeid', 'eeid', 'leader_org_name')
oath['L3_org_name'] = vlookup(oath, orgnames, 'L3_eeid', 'eeid', 'leader_org_name')
oath['L4_org_name'] = vlookup(oath, orgnames, 'L4_eeid', 'eeid', 'leader_org_name')

# Remove duplicate employees -- AOLers with laptops deployed on the Yahoo network
oath = oath.loc[oath['eeid']!='Y00000 ']

# Create badge_id field for AlixPartners -- NEEDS TO BE DONE AFTER USERID FIELD HAS BEEN CLEANED
oath['badge_id'] = None
oath.loc[is_yahoo, 'badge_id'] = oath['userid']
oath.loc[(~is_yahoo), 'badge_id'] = oath['aol_eeid']

# Remove employees with Last Days of Work < Current Day
oath = oath.loc[(oath['last_day_of_work'].isnull()) | (oath['last_day_of_work'] >= datetime.datetime.now().strftime("%Y-%m-%d 0:00:00"))]

oath = oath.sort_values(by='eeid', ascending=True)
employees = oath.loc[oath['worker_type']=='Employee']

DATETIMESTAMP = datetime.datetime.now().strftime("%Y-%m-%d %H_%M PDT")

# Current Worker Details columns (includes contingent workers)
cwd_nonsens_cols = ['worker_type','emp_type','eeid','legal_name','CEO_eeid','CEO_name','L2_eeid','L2_name','L3_eeid','L3_name',\
                    'L4_eeid','L4_name','L5_eeid','L5_name','L6_eeid','L6_name','L7_eeid','L7_name',\
                    'L8_eeid','L8_name','L9_eeid','L9_name','L10_eeid','L10_name','L2_org_name',\
                    'L3_org_name','L4_org_name']

cwd_nonsens = oath.loc[:, cwd_nonsens_cols]

writer_cwd = pandas.ExcelWriter('outputs/test '+DATETIMESTAMP+'.xlsx')
cwd_nonsens.to_excel(writer_cwd, 'Sheet1', index=False)
writer_cwd.save()

# # Current Worker Details columns (includes contingent workers)
# cwd_nonsens_cols = ['worker_type','emp_type','eeid','legal_name','mgr_eeid','mgr_legal_name',\
#                     'mgr_email','userid','last_hire_date','original_hire_date','active_status',\
#                     'ft_or_pt','fte_pct','email','acquired_company','job_code','job_profile',\
#                     'job_family_group','job_family','job_level','job_category','mgmt_level',\
#                     'comp_grade_profile','pay_rate_type','work_office','wfh_flag','work_country',\
#                     'work_region','CEO_eeid','CEO_name','L2_eeid','L2_name','L3_eeid','L3_name',\
#                     'L4_eeid','L4_name','L5_eeid','L5_name','L6_eeid','L6_name','L7_eeid','L7_name',\
#                     'L8_eeid','L8_name','L9_eeid','L9_name','L10_eeid','L10_name','L2_org_name',\
#                     'L3_org_name','L4_org_name']

# cwd_nonsens = oath.loc[:, cwd_nonsens_cols]

# writer_cwd = pandas.ExcelWriter('outputs/Current Worker Details - Non-Sensitive '+DATETIMESTAMP+'.xlsx')
# cwd_nonsens.to_excel(writer_cwd, 'Sheet1', index=False)
# writer_cwd.save()

# # Comp Kitchen Sink columns (includes contingent workers)
# cks_cols = ['worker_type','emp_type','eeid','legal_name','mgr_eeid','mgr_legal_name','mgr_email',\
#             'userid','last_hire_date','original_hire_date','active_status','ft_or_pt','fte_pct',\
#             'std_hrs','email','acquired_company','job_code','job_profile','job_family_group',\
#             'job_family','job_level','job_category','mgmt_level','comp_grade','comp_grade_profile',\
#             'pay_rate_type','flsa','base_annualized_local','currency_code','base_annualized_usd',\
#             'bonus_plan','target_bonus_pct','target_bonus_amt_local','target_bonus_amt_usd',\
#             'ttc_annualized_local','ttc_annualized_usd','wfh_flag','work_office','work_city',\
#             'work_state','work_country','work_region','CEO_eeid','CEO_name','L2_eeid','L2_name',\
#             'L3_eeid','L3_name','L4_eeid','L4_name','L5_eeid','L5_name','L6_eeid','L6_name',\
#             'L7_eeid','L7_name','L8_eeid','L8_name','L9_eeid','L9_name','L10_eeid','L10_name',\
#             'L2_org_name','L3_org_name','L4_org_name']
# cks = employees.loc[:, cks_cols]
# writer = pandas.ExcelWriter('outputs/Oath Comp Kitchen Sink '+DATETIMESTAMP+'.xlsx')
# cks.to_excel(writer,'Sheet1', index=False)
# writer.save()

# cwd_sens_cols = ['worker_type','emp_type','eeid','legal_name','mgr_eeid','mgr_legal_name','mgr_email',\
#                  'userid','last_hire_date','original_hire_date','active_status','ft_or_pt','fte_pct',\
#                  'std_hrs','email','acquired_company','job_code','job_profile','job_family_group',\
#                  'job_family','job_level','job_category','mgmt_level','comp_grade','comp_grade_profile',\
#                  'pay_rate_type','flsa','base_annualized_local','currency_code','base_annualized_usd',\
#                  'bonus_plan','target_bonus_pct','target_bonus_amt_local','target_bonus_amt_usd',\
#                  'ttc_annualized_local','ttc_annualized_usd','wfh_flag','work_office','work_city',\
#                  'work_state','work_country','work_region','CEO_eeid','CEO_name','L2_eeid','L2_name',\
#                  'L3_eeid','L3_name','L4_eeid','L4_name','L5_eeid','L5_name','L6_eeid','L6_name',\
#                  'L7_eeid','L7_name','L8_eeid','L8_name','L9_eeid','L9_name','L10_eeid','L10_name',\
#                  'L2_org_name','L3_org_name','L4_org_name','last_day_of_work','term_date']
# cwd_sens = oath.loc[:, cwd_sens_cols]
# writer_cwd_sens = pandas.ExcelWriter('outputs/Current Worker Details - Sensitive with Demographic Data ' + DATETIMESTAMP + '.xlsx')
# cwd_sens.to_excel(writer_cwd_sens, 'Sheet1', index=False)
# writer_cwd_sens.save()

# alixpartners_cols = ['worker_type','emp_type','eeid','badge_id','legal_name','mgr_eeid','mgr_legal_name','mgr_email',\
#                      'userid','last_hire_date','original_hire_date','active_status','ft_or_pt','fte_pct',\
#                      'std_hrs','email','acquired_company','job_code','job_profile','job_family_group',\
#                      'job_family','job_level','job_category','mgmt_level','comp_grade','comp_grade_profile',\
#                      'pay_rate_type','flsa','currency_code','base_annualized_local','base_annualized_usd',\
#                      'target_abp_plan_yr','target_abp_pct','target_abp_amt_local','target_abp_amt_usd',\
#                      'target_abp_exception_flag','sales_incentive_plan_yr','sales_incentive_target_amt_local',\
#                      'sales_incentive_target_amt_usd','sales_incentive_guarantee','yahoo_bonus_plan','yahoo_target_bonus_pct',\
#                      'target_bonus_amt_local','target_bonus_amt_usd',\
#                      'ttc_annualized_local','ttc_annualized_usd','wfh_flag','work_office','work_city',\
#                      'work_state','work_country','work_region','CEO_eeid','CEO_name','L2_eeid','L2_name',\
#                      'L3_eeid','L3_name','L4_eeid','L4_name','L5_eeid','L5_name','L6_eeid','L6_name',\
#                      'L7_eeid','L7_name','L8_eeid','L8_name','L9_eeid','L9_name','L10_eeid','L10_name',\
#                      'L2_org_name','L3_org_name','L4_org_name','last_day_of_work','term_date']
# alixpartners = oath.loc[:, alixpartners_cols]
# alixpartners.columns = ['Worker Type','Employee Type','EEID','Badge ID','Legal Name','Direct Supervisor - EEID','Direct Supervisor - Legal Name','Direct Supervisor - Email','User ID','Last Hire Date','Original Hire Date','Active Status','Full time / Part time','FTE %','Standard Hours','Email','Acquired Company','Job Code','Job Profile','Job Family Group','Job Family','Job Level','Job Category','Management Level','Comp Grade','Comp Grade Profile','Pay Rate Type','FLSA','Local Currency','Base Annualized (Local)','Base Annualized (USD)','Target ABP Plan Year','Target ABP %','Target ABP Amount (Local)','Target ABP Amount (USD)','Target ABP Exception Flag','AOL Sales Incentive Plan Year','AOL Sales Incentive Target Amount (Local)','AOL Sales Incentive Target Amount (USD)','AOL Sales Incentive Guarantee','Yahoo Bonus Plan','Yahoo Target Bonus %','Yahoo Target Bonus Amount (Local)','Yahoo Target Bonus Amount (USD)','TTC Annualized (Local)','TTC Annualized (USD)','WFH Flag','Work Location - Office','Work Location - City','Work Location - State','Work Location - Country','Work Location - Region','CEO EEID','CEO','L2 EEID','L2','L3 EEID','L3','L4 EEID','L4','L5 EEID','L5','L6 EEID','L6','L7 EEID','L7','L8 EEID','L8','L9 EEID','L9','L10 EEID','L10','L2 Org Name','L3 Org Name','l4 Org Name','Last Day of Work','Term Date']
# writer_alixpartners = pandas.ExcelWriter('outputs/Oath Current Employee Details for AlixPartners ' + DATETIMESTAMP + '.xlsx')
# alixpartners.to_excel(writer_alixpartners, 'Sheet1', index=False)
# writer_alixpartners.save()


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

