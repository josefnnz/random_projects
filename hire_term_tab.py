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
    mleft = left.copy()
    mright = right.loc[:, [right_key, right_col]]
    mright.columns = [left_key, 'rval_lookup_col']
    output = pandas.merge(mleft, mright, how='left', on=left_key)
    output.loc[output.loc[:, 'rval_lookup_col'].notnull(), left_col] = output.loc[:, 'rval_lookup_col']
    output.drop('rval_lookup_col', axis=1, inplace=True)
    return output

################################################################################

setwd = "/Users/josefnunez/hire_terms/"

# load PeopleSoft hires and terms report
ps_action_filepath, ps_action_sheet = setwd+"ps_actions.xlsx", "Sheet1"
ps_action = pandas.ExcelFile(ps_action_filepath).parse(ps_action_sheet)
ps_action.columns = ['worker_type','yahoo_eeid','badge_id','aol_eeid','legal_name','effective_date',\
				     'action_date','action','action_reason','CEO_name','L2_pseeid','L2_name','L3_pseeid',\
				     'L3_name','L4_pseeid','L4_name','L5_pseeid','L5_name','L6_pseeid','L6_name','L7_pseeid',\
				     'L7_name','L8_name','L9_name','L10_name','comp_grade','comp_grade_profile',\
				     'headcount_group','status','regular_or_temp','ft_or_pt','gender','ethnicity',\
				     'marital_status','military_status','email','userid','mgr_pseeid','mgr_legal_name',\
				     'mgr_email','manager_userid','acquired_company','contract_type','contractor_type',\
				     'contract_number','contract_status','msp_or_nonmsp','contract_start_date',\
				     'contract_end_date','contract_provider_id','contract_provider','birth_date',\
				     'original_hire_date','last_hire_date','department_entry_date',\
				     'comp_grade_entry_date','job_entry_date','benefit_program',\
				     'base_annualized_local','currency_code','base_annualized_usd','target_abp_plan_yr',\
				     'target_abp_pct','target_abp_exception_flag','target_abp_local','target_abp_usd',\
				     'sip_year','sales_begin_date','sales_end_date','sip_target_local','sip_target_usd',\
				     'sip_guarantee','wfh_flag','work_country_code','work_location_code','work_office',\
				     'work_country','work_state','work_city','work_postal_code','home_country','home_state',\
				     'home_ciity','home_postal_code','company_code','company','market_cluster','job_code',\
				     'job_profile','job_family','aap_job_classification','eeo_job_classification',\
				     'comp_freq','std_hrs','fte_pct','flsa','department_code','department_name',\
				     'business_unit','division','region_code','reporting_schema_level_1',\
				     'reporting_schema_level_2','hr_support_eeid','hr_support_name','hr_support_userid',\
				     'separations_group','separation_date']

wd_action_filepath, wd_action_sheet, wd_action_skiprows = setwd+"wd_actions.xlsx", "Sheet1", 5
wd_action = pandas.ExcelFile(wd_action_filepath).parse(wd_action_sheet, skiprows=wd_action_skiprows)
wd_action.columns = ['worker_type','emp_type','eeid','emp_preferred_name','hire_or_term',\
                     'hire_or_term_effective_date','term_type','email','userid','acquired_company',\
                     'mgr_eeid','mgr_preferred_name','mgr_email','mgr_userid','job_code','job_profile',\
                     'job_family_group','job_family','job_category','job_level','mgmt_level','wfh_flag',\
                     'work_office','work_city','work_state','work_country','work_region','CEO_name','L2_name',\
                     'L3_name','L4_name','L5_name','L6_name','L7_name','L8_name','L9_name','L2_org_name','L3_org_name','L4_org_name']
wd_action['CEO_eeid'] = None; wd_action['L2_eeid'] = None; wd_action['L3_eeid'] = None; wd_action['L4_eeid'] = None;
wd_action['L5_eeid'] = None; wd_action['L6_eeid'] = None; wd_action['L7_eeid'] = None; wd_action['L8_eeid'] = None; 
wd_action['L9_eeid'] = None; wd_action['L10_eeid'] = None;

################################################################################
################################################################################
################################################################################
##### Load Mappings Tables

# Excel workbook containing tabs of mapping tables
mappings = pandas.ExcelFile("/Users/josefnunez/hire_terms/mappings.xlsx")

# load PeopleSoft to Workday office names
offices = mappings.parse("Offices")
offices.columns = ['ps_office_name','wd_office_name','wfh_flag']
offices.drop_duplicates('ps_office_name', inplace=True)

# load country to region
regions = mappings.parse('Regions')
regions.columns = ['country', 'region']
regions.drop_duplicates('country', inplace=True)

# load USA state names
usa_states = mappings.parse("States")
usa_states.columns = ['state_name','state_code']
usa_states.drop_duplicates('state_code', inplace=True)

# load L2 legal name to L2 org name
orgnames = mappings.parse("OrgNames")
orgnames.columns = ['layer','eeid','leader_name','leader_org_name']
orgnames.drop_duplicates('eeid', inplace=True)

# load L2 Workday names to EEID
L2_name_to_eeid = mappings.parse("L2OrgNames")
L2_name_to_eeid.columns = ['L2_L3_flag','L2_eeid','PSFT_L2_name','WD_L2_name','PSFT_L3_name','WD_L3_name','orgname']
L2_name_to_eeid.drop_duplicates('L2_eeid', inplace=True)

# load PeopleSoft legal name to employee id mapping
name_to_eeid = mappings.parse('LegalNameToEEID')
name_to_eeid.colunmns = ['eeid','legal_name']
name_to_eeid.drop_duplicates('eeid', inplace=True)

# load manager hierarchies for orphans
orphan_cleanup = mappings.parse("OrphanCleanup")
orphan_cleanup.columns = ['eeid','legal_name','CEO_eeid','CEO_name','L2_eeid','L2_name','L3_eeid','L3_name',\
                          'L4_eeid','L4_name','L5_eeid','L5_name','L6_eeid','L6_name','L7_eeid','L7_name',\
                          'L8_eeid','L8_name','L9_eeid','L9_name','L10_eeid','L10_name']
orphan_cleanup.drop_duplicates('eeid', inplace=True)

################################################################################
################################################################################
################################################################################
##### Reformat PeopleSoft Data

# create worker type field and filter out contingents
ps_action['worker_type'].replace({'CWR':'Contingent Worker',\
	                              'POI':'Contingent Worker',\
	                              'EMP':'Employee'},\
	                              inplace=True)
ps_action = ps_action.loc[ps_action['worker_type'] == 'Employee']

# create acquired company field (AOL/Yahoo) and merge AOL/Yahoo numeric eeids into one column
is_yahoo = ps_action['company'].str.contains("yahoo",case=False)
ps_action.loc[is_yahoo, 'acquired_company'] = 'Yahoo'
ps_action.loc[~is_yahoo, 'acquired_company'] = 'AOL'
ps_action['eeid'] = None
is_yahoo_with_workday_id = (is_yahoo) & (ps_action['yahoo_eeid'].notnull())
ps_action.loc[is_yahoo_with_workday_id, 'eeid'] = ps_action['yahoo_eeid']
ps_action.loc[ps_action['eeid'].isnull(), 'eeid'] = ps_action['aol_eeid']
ps_action['eeid'] = ps_action['eeid'].apply('{0:0>6}'.format)
ps_action.loc[is_yahoo, 'eeid'] = 'Y' + ps_action['eeid']
ps_action.loc[~is_yahoo, 'eeid'] = 'A' + ps_action['eeid']

# remove Yahoos
ps_action = ps_action.loc[ps_action['acquired_company'] == 'AOL']

# change name fields format from "Last, First" to "First Last"
for x in ['legal_name','mgr_legal_name','CEO_name','L2_name','L3_name','L4_name','L5_name','L6_name','L7_name','L8_name','L9_name','L10_name']:
    ps_action[x] = [" ".join(reversed(w.split(", "))) for w in ps_action[x]]

# lookup AOL/Yahoo numeric eeids for management chain (CEO -> L10)
ps_action['CEO_eeid'] = 'A188900'
L2_L7_pseeid_cols = ['mgr_legal_name','L2_name','L3_name','L4_name','L5_name','L6_name','L7_name','L8_name','L9_name','L10_name']
L2_L7_eeid_cols = ['mgr_eeid','L2_eeid','L3_eeid','L4_eeid','L5_eeid','L6_eeid','L7_eeid','L8_eeid','L9_eeid','L10_eeid']
L2_L7_lookup_cols = ['legal_name']*10
for i in range(len(L2_L7_pseeid_cols)):
    curr_pseeid, curr_eeid, curr_lookup_col = L2_L7_pseeid_cols[i], L2_L7_eeid_cols[i], L2_L7_lookup_cols[i]
    ps_action = vlookup(ps_action, name_to_eeid, curr_pseeid, curr_lookup_col, 'eeid', curr_eeid)     

# create layer field (is L1,L2,L3,...)
L10_L1_eeid_cols = ['L10_eeid','L9_eeid','L8_eeid','L7_eeid','L6_eeid','L5_eeid','L4_eeid','L3_eeid','L2_eeid','CEO_eeid']
ps_action['layer'] = 10
for i in range(len(L10_L1_eeid_cols)):
    ps_action.loc[ps_action[L10_L1_eeid_cols[i]].isnull(), 'layer'] = 10 - i # assign layer number
ps_action['layer'] = ps_action['layer'] - 1
ps_action.loc[ps_action['CEO_name'] == 'Orphan', 'layer'] = None


# merge in Workday office names
ps_action = vlookup_update(ps_action, offices, 'work_office', 'ps_office_name', 'work_office', 'wd_office_name')
ps_action['wfh_flag'].replace({'No':None, 'Yes':'WFH'}, inplace=True)
ps_action.loc[(ps_action['wfh_flag'] == 'WFH') & (ps_action['work_country'] == 'United States Of America'), 'work_city'] = 'WFH' # fill in empty city values for USA WFH employees

# merge in Workday region names and USA states
ps_action = vlookup(ps_action, regions, 'work_country', 'country', 'region', 'work_region')
ps_action = vlookup_update(ps_action, usa_states, 'work_state', 'state_code', 'work_state', 'state_name')
ps_action.loc[ps_action['work_office'] == 'UK - VDMS Offsite Contractors', 'work_country'] = 'United Kingdom'
ps_action.loc[ps_action['work_office'] == 'UK - VDMS Offsite Contractors', 'work_region'] = 'EMEA'

# merge in L2-L4 org names and L2/L3 org name grouping for workforce report
ps_action = vlookup(ps_action, orgnames, 'L2_eeid', 'eeid', 'leader_org_name', 'L2_org_name')
ps_action = vlookup(ps_action, orgnames, 'L3_eeid', 'eeid', 'leader_org_name', 'L3_org_name')
ps_action = vlookup(ps_action, orgnames, 'L4_eeid', 'eeid', 'leader_org_name', 'L4_org_name')
ps_action.loc[ps_action['eeid'] == 'A188900', 'L2_org_name'] = 'CEO Office'
ps_action['L2_or_L3_org_name'] = ps_action['L2_org_name']
ps_action.loc[ps_action['L3_org_name'] == 'Facilities', 'L2_or_L3_org_name'] = 'Facilities'
ps_action.loc[ps_action['L3_org_name'] == 'Small Business', 'L2_or_L3_org_name'] = 'Small Business'
ps_action.loc[ps_action['L3_org_name'] == 'Small Business Engineering', 'L2_or_L3_org_name'] = 'Small Business'
ps_action.loc[(ps_action['layer'] == 1) | (ps_action['layer'] == 2), 'L3_org_name'] = ps_action['legal_name']
ps_action.loc[ps_action['layer'] == 3, 'L4_org_name'] = ps_action['legal_name']
ps_action.loc[ps_action['L4_org_name'].isnull(), 'L4_org_name'] = ps_action['L4_name']

# create term / hire flag
ps_action['hire_or_term'] = ps_action['action']
ps_action['hire_or_term'].replace({'Hire':'Hire',\
	                               'Rehire':'Hire',\
	                               'Termination - Involuntary':'Term',\
	                               'Termination - Voluntary':'Term'},\
	                               inplace=True)

# create term type field
ps_action['term_type'] = None
ps_action.loc[(ps_action['action'] == 'Termination - Voluntary'), 'term_type'] = 'Voluntary'
ps_action.loc[(ps_action['action'] == 'Termination - Involuntary'), 'term_type'] = 'Involuntary'

ps_action['L2_L3_lookup_val'] = ps_action['L2_org_name'] + '-' + ps_action['L3_org_name']

ps_action = ps_action.sort_values(by='eeid', ascending=True)

################################################################################
################################################################################
################################################################################
##### Reformat Workday data

# flag as Yahoo hires/terms
wd_action['company'] = 'Yahoo'

# merge in CEO/L2/L3 eeids
wd_action['CEO_eeid'] = 'A188900'
wd_action = vlookup_update(wd_action, L2_name_to_eeid, 'L2_name', 'WD_L2_name', 'L2_eeid', 'L2_eeid')

# merge in L2 eeids and then merge in L2 org name using L2 eeid
wd_action = vlookup_update(wd_action, orgnames, 'L2_eeid', 'eeid', 'L2_org_name', 'leader_org_name')


hierarchy_cols = ['CEO_eeid','CEO_name','L2_eeid','L2_name','L3_eeid','L3_name','L4_eeid','L4_name','L5_eeid','L5_name',\
                  'L6_eeid','L6_name','L7_eeid','L7_name','L8_eeid','L8_name','L9_eeid','L9_name','L10_eeid','L10_name']
for i in range(len(hierarchy_cols)):
    wd_action = vlookup_update(wd_action, orphan_cleanup, 'eeid', 'eeid', hierarchy_cols[i], hierarchy_cols[i])
wd_action = vlookup_update(wd_action, orgnames, 'L3_eeid', 'eeid', 'L3_org_name', 'leader_org_name')

wd_action['L2_or_L3_org_name'] = wd_action['L2_org_name']
wd_action['L2_L3_lookup_val'] = wd_action['L2_org_name'] + '-' + wd_action['L3_org_name']


DATETIMESTAMP = datetime.datetime.now().strftime("%Y-%m-%d %H_%M PDT")

# aol_actions_cols = ['acquired_company','action']
aol_actions_cols = ['acquired_company','eeid','effective_date','hire_or_term','term_type','CEO_eeid','L2_eeid','L3_eeid',\
                    'CEO_name','L2_name','L3_name','L2_org_name','L3_org_name','L2_or_L3_org_name','L2_L3_lookup_val']
aol_actions = ps_action.loc[:, aol_actions_cols]

writer_aol_actions = pandas.ExcelWriter('outputs/hires_terms_tab '+DATETIMESTAMP+'.xlsx')
aol_actions.to_excel(writer_aol_actions, 'Sheet1', index=False)
writer_aol_actions.save()

yahoo_actions_cols = ['acquired_company','eeid','hire_or_term_effective_date','hire_or_term','term_type','CEO_eeid','L2_eeid','L3_eeid',\
                      'CEO_name','L2_name','L3_name','L2_org_name','L3_org_name','L2_or_L3_org_name','L2_L3_lookup_val']
yahoo_actions = wd_action.loc[:, yahoo_actions_cols]
writer_yahoo_actions = pandas.ExcelWriter('outputs/hires_terms_yahoo_tab '+DATETIMESTAMP+'.xlsx')
yahoo_actions.to_excel(writer_yahoo_actions, 'Sheet1', index=False)
writer_yahoo_actions.save()



