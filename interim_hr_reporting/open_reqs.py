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

setwd = "/Users/josefnunez/open_reqs/"

# Excel workbook containing tabs of mapping tables
mappings = pandas.ExcelFile("/Users/josefnunez/open_reqs/mappings.xlsx")

# load country to region
regions = mappings.parse('Regions')
regions.columns = ['country', 'region']
regions.drop_duplicates('country', inplace=True)

taleo_L2s = mappings.parse('TaleoL2s')
taleo_L2s.columns = ['ATS','L2_or_L3','eeid','psft_leader_name','wd_leader_name','orgname']
taleo_L2s.drop_duplicates('eeid', inplace=True)

sr_L2s = mappings.parse('SmartRecruiterL2s')
sr_L2s.columns = ['ATS','L2_or_L3','eeid','psft_leader_name','wd_leader_name','orgname']
sr_L2s.drop_duplicates('eeid', inplace=True)

# load Taleo open reqs
taleo_filepath, taleo_sheet = "taleo.xlsx", "Sheet1"
taleo = pandas.ExcelFile(taleo_filepath).parse(taleo_sheet)
taleo.columns = ['reqid','L2_name','L3_name','L4_name','L5_name','recruiter','hiring_mgr','job_code',\
                 'job_profile','job_category','date_opened','work_office','work_city','work_state',\
                 'work_country','work_region']

# create Taleo flag
taleo['ATS'] = 'Taleo'
taleo['CEO_eeid'] = 'A188900'
taleo['CEO_name'] = 'Timothy Armstrong'

# change name fields format from "Last, First" to "First Last"
for x in ['L2_name','L3_name','L4_name','L5_name','hiring_mgr','recruiter']:
    taleo[x] = taleo[x].astype('str')
    taleo[x] = [" ".join(reversed(w.split(", "))) for w in taleo[x]]
    taleo.loc[taleo[x] == 'nan', x] = None

taleo = vlookup(taleo, taleo_L2s, 'L2_name', 'wd_leader_name', 'eeid', 'L2_eeid')
taleo = vlookup(taleo, taleo_L2s, 'L2_eeid', 'eeid', 'orgname', 'L2_or_L3_org_name')

# reformat region names
taleo['work_region'].replace({'Americas':'AMER','Europe - Middle East - Africa':'EMEA','Asia Pacific':'APAC'}, inplace=True)

# load SmartRecruiter open reqs
sr_filepath, sr_sheet = "sr.xlsx", "Sheet1"
sr = pandas.ExcelFile(sr_filepath).parse(sr_sheet)
sr.columns = ['req_status','reqid','job_id','posted_job_id','position_id','position_status',\
              'num_positions','posting_title','current_req_status','job_function','job_category',\
              'hc_group','sr_cost_center','cost_center','dpmt_desc','business_unit','division',\
              'region_code','reporting_schema_lvl_1','reporting_schema_lvl_2','L2_pseeid','L2_name',\
              'L3_pseeid','L3_name','L4_pseeid','L4_name','L5_pseeid','L5_name','L6_pseeid','L6_name',\
              'L7_pseeid','L7_name','hiring_mgr_pseeid','hiring_mgr','recruiter_id','recruiter',\
              'work_country','country_code','work_city','work_state','sr_job_code','sr_job_title',\
              'job_code','job_profile','job_family','range_min','range_mid','range_max','currency',\
              'range_min_usd','range_mid_usd','range_max_usd','sr_ftc','sr_job_category','sr_business',\
              'sr_brand','sr_function','sr_industry','sr_hiring_process','sr_experience_lvl',\
              'flsa_status','area_of_interest','relo_program','bonus_eligible','location_code',\
              'sr_office_location','hr_coordinator','created_on','updated_on','sourcing_on',\
              'interview_on','offer_made_on','closed_on','canceled_on','applicants',\
              'reviewed_or_rejected','new_applicant_lead','in_review','interview','offered',\
              'req_created_on','position_created_on','position_opened_on','target_start_date',\
              'actual_start_date','position_type','add_or_replace','incumbent','internal_hire',\
              'internal_hire_last_name','internal_hire_first_name']

['ATS','L2_or_L3','eeid','psft_leader_name','wd_leader_name','orgname']

# filter out closed reqs
sr = sr.loc[sr['req_status'] != 'Closed Since Close']

# create SmartRecruiter flag
sr['ATS'] = 'SmartRecruiter'
sr['CEO_eeid'] = 'A188900'
sr['CEO_name'] = 'Timothy Armstrong'

# extend United States label and rename region values to  AMER/EMEA/APAC
sr['work_country'].replace({'United States':'United States Of America'}, inplace=True)
sr = vlookup_update(sr, regions, 'work_country', 'country', 'work_region', 'region')

# change name fields format from "Last, First" to "First Last"
for x in ['L2_name','L3_name','L4_name','L5_name','hiring_mgr','recruiter']:
    sr[x] = [" ".join(reversed(w.split(", "))) for w in sr[x]]

sr = vlookup(sr, sr_L2s, 'L2_name', 'wd_leader_name', 'eeid', 'L2_eeid')
sr = vlookup(sr, sr_L2s, 'L2_eeid', 'eeid', 'orgname', 'L2_or_L3_org_name')

DATETIMESTAMP = datetime.datetime.now().strftime("%Y-%m-%d %H_%M PDT")

common_column_names = ['ATS','date_opened','req_id','legacy_job_code','legacy_job_title','recruiter',\
                       'hiring_manager','work_office','work_city','work_state','work_country','work_region',\
                       'CEO_eeid','CEO_name','L2_eeid','L2_name','L3_name','L4_name','L5_name','L2_or_L3_org_name']

taleo = taleo.loc[:, ['ATS','date_opened','reqid','job_code','job_profile','recruiter','hiring_mgr',\
                      'work_office','work_city','work_state','work_country','work_region','CEO_eeid','CEO_name',\
                      'L2_eeid','L2_name','L3_name','L4_name','L5_name','L2_or_L3_org_name']]
# writer = pandas.ExcelWriter('outputs/combined_reqs ' + DATETIMESTAMP + '.xlsx')
# taleo.to_excel(writer, 'taleo_reqs', index=False)

sr = sr.loc[:, ['ATS','date_opened','reqid','job_code','job_profile','recruiter','hiring_mgr',\
                'sr_office_location','work_city','work_state','work_country','work_region','CEO_eeid','CEO_name',\
                'L2_eeid','L2_name','L3_name','L4_name','L5_name','L2_or_L3_org_name']]
# sr.to_excel(writer, 'sr_reqs', index=False)
# writer.save()

taleo.columns = common_column_names
sr.columns = common_column_names

output = taleo.append(sr, ignore_index=True)
writer = pandas.ExcelWriter('outputs/Oath Open Reqs ' + DATETIMESTAMP + '.xlsx')
output.to_excel(writer, 'Sheet1', index=False)
writer.save()


