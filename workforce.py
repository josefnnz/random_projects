import pandas

# load Oath workforce roster, i.e. PeopleSoft report
oath = pandas.ExcelFile("/Users/josefnunez/workforce/roster.xlsx").parse("Sheet1")
oath.columns = ['eeid','legal_name','CEO','L2','L3','L4','L5','L6','L7','L8','L9','L10','job_level','comp_grade',\
                'mgmt_level','status','regular_or_temp','full_time_or_part_time','gender','ethnicity','marital_status',\
                'military_status','work_email','userid','mgr_eeid','mgr_legal_name','mgr_work_email','mgr_userid',\
                'acquired_company','birth_date','original_hire_date','last_hire_date','department_entry_date',\
                'job_level_entry_date','job_entry_date','benefit_program','base_annualized_local','local_currency',\
                'base_annualized_usd','target_abp_plan_yr','target_abp_pct','target_abp_exception_flag',\
                'target_abp_amt_local','target_abp_amt_usd','sales_incentive_plan_yr','sales_beginning_date',\
                'sales_end_date','sales_incentive_target_amt_local','sales_incentive_target_amt_usd',\
                'sales_incentive_guarantee','work_from_home_flag','work_country_code','location_code','work_office',\
                'work_country','work_state_code','work_city','work_postal_code','home_country','home_state_code',\
                'home_city','home_postal_code','company_code','company','market_cluster','job_code','job_profile',\
                'job_family','aol_eeo_group','eeo_job_classification','comp_frequency','standard_hours','fte_pct',\
                'flsa','department_code','department_name','business_unit','division','region_code','reporting_schema_level_1',\
                'reporting_schema_level_2','hr_support_eeid','hr_support_name','hr_support_userid','separations_group',\
                'separations_date','layer','mgr_userid_hierarchy','span']

# load Yahoo comp
ycomp = pandas.ExcelFile("/Users/josefnunez/workforce/yahoo_comp.xlsx").parse("Sheet1", skiprows=2)
ycomp.columns = ['eeid','emp_preferred_name','work_email','emp_type','job_code','job_profile','job_family_group','job_family','job_level',\
                 'job_category','comp_grade','comp_grade_profile','local_currency','base_annualized_in_local',\
                 'base_annualized_in_usd','fx_rate','bonus_plan','target_bonus_pct']

# load Yahoo active workers
yactive = pandas.ExcelFile("/Users/josefnunez/workforce/yahoo_active_workers.xlsx").parse("Sheet1", skiprows=1)
yactive.columns = ['eeid','work_email','userid','worker_type','emp_type']

# load final offboards list from AlixPartners
offboards = pandas.ExcelFile("/Users/josefnunez/workforce/offboards.xlsx").parse("Sheet1")
offboards.columns = ['work_email','badge_id','company','last_day_of_work']

# load mappings tables
mappings = pandas.ExcelFile("/Users/josefnunez/workforce/mappings.xlsx") # Excel workbook containing tabs of mapping tables
offices = mappings.parse("Offices") # tab containing PS office name to WD office name
offices.columns = ['ps_office_name','wd_office_name','wfh_flag']
orgnames = mappings.parse("SupOrgNames") # tab containing L2/L3 to supervisory org name
orgnames.columns = ['L2_or_L3','ps_L2_name','wd_L2_name','ps_L3_name','wd_L2_name','orgname']
jobs = mappings.parse("Jobs") # tab containing Aol/Yahoo jobs to Oath jobs
jobs.columns = ['company','old_job_profile','job_code','job_profile','job_family_grp','job_family','job_category_sort_order',\
                'job_category','job_level','mgmt_level','eeo_job_classification','aap_job_classification','pay_rate_type',\
                'is_exempt','comp_grade']
georegions = mappings.parse("GeoRegions")
georegions.columns = ['country','city_and_state','city','state','georegion']
emailremap = mappings.parse("EmailRemap")
emailremap.columns = ['emp_name','aol_work_email','yahoo_work_email']

# update Yahoos with AOL email address with their current Yahoo email
oath.replace(emailremap.set_index('aol_work_email').to_dict()['yahoo_work_email'], inplace=True)

# merge in Workday office names
oath = pandas.merge(oath, offices, how='left', left_on='work_office', right_on='ps_office_name')

# remove employees either (1) offboarded, (2) on transition, (3) future term
oath = oath[~oath['work_email'].isin(offboards['work_email'])]
oath = oath[oath['separations_date'].isnull()]
#oath = oath[(~oath['company'].str.contains("yahoo",case=False)) | (oath['work_email'].isin(yactive))] # remove inactive Yahoos -- keep all Aolers

# merge in Yahoo job profiles for Yahoos, and then merge in oath job details
oath = pandas.merge(oath, ycomp[['work_email','job_profile']], how='left', on='work_email')
oath.loc[oath['job_profile_y'].notnull(), 'job_profile_x'] = oath['job_profile_y']
oath = pandas.merge(oath, jobs, how='left', left_on='job_profile_x', right_on='old_job_profile')

# merge in Oath geo regions
oath['work_country'].replace({"Korea, Republic of" : "Republic of Korea"}, inplace=True) # reformat Republic of Korea
oath = pandas.merge(oath, georegions[['city','georegion']], how='left', left_on='work_city', right_on='city')
oath.loc[~oath['work_country'].str.contains("united states of america",case=False), 'georegion'] = oath['work_country']
oath.loc[oath['wfh_flag'].notnull(), 'georegion'] = "WFH"

# merge Yahoo comp details
oath = pandas.merge(oath, ycomp[['work_email','comp_grade','local_currency','base_annualized_in_local','target_bonus_pct','bonus_plan','fx_rate']], how='left', on='work_email')
oath.loc[oath['base_annualized_in_local'].notnull(), 'base_annualized_local'] = oath['base_annualized_in_local'] # put Yahoo base values in main column
oath.loc[oath['target_bonus_pct'].notnull(), 'target_abp_pct'] = oath['target_bonus_pct'] # put Yahoo target bonus % in main column
oath.loc[oath['local_currency_y'].notnull(), 'local_currency_x'] = oath['local_currency_y'] # put Yahoo local currency in main column
oath.loc[(oath['sales_beginning_date'].notnull()) & (oath['sales_incentive_target_amt_local'] > 0), 'target_abp_pct'] = oath['sales_incentive_target_amt_local'] / oath['base_annualized_local'] # compute Sales bonus targets


# merge in L2 and L3 org names
oath = pandas.merge(oath, orgnames.query('L2_or_L3 == "L2"')[['ps_L2_name','orgname']], how='left', left_on='L2', right_on='ps_L2_name')
oath.rename(columns={'orgname' : 'L2_orgname'}, inplace=True)
oath = pandas.merge(oath, orgnames.query('L2_or_L3 == "L3"')[['ps_L3_name','orgname']], how='left', left_on='L3', right_on='ps_L3_name')
oath.rename(columns={'orgname' : 'L3_orgname'}, inplace=True)

# change name fields format from "Last, First" to "First Last"
for x in ['legal_name','mgr_legal_name','CEO','L2','L3','L4','L5','L6','L7','L8','L9','L10']:
	oath[x] = [" ".join(reversed(w.split(", "))) for w in oath[x]]

# extract relevant fields
final_report = oath[['eeid','legal_name','work_email','userid','status','last_hire_date','regular_or_temp','full_time_or_part_time',\
                        'acquired_company','mgr_eeid','mgr_legal_name','mgr_work_email','mgr_userid','fte_pct','job_code_y','job_profile',\
                        'job_family_grp','job_family_y','job_category','job_level_y','mgmt_level_y','eeo_job_classification_y',\
                        'aap_job_classification','comp_frequency','pay_rate_type','standard_hours','flsa','local_currency_x','fx_rate',\
                        'base_annualized_local','target_abp_pct','target_abp_exception_flag','sales_incentive_guarantee','wd_office_name',\
                        'wfh_flag','work_city','work_state_code','work_postal_code','work_country','comp_grade_y','georegion','span','CEO',\
                        'L2','L3','L4','L5','L6','L7','L8','L9','L10','L2_orgname','L3_orgname','gender','ethnicity','marital_status','military_status','benefit_program']]

final_report.columns = ['Employee ID','Legal Name','Work Email','User ID','Status','Last Hire Date','Regular / Temp','FT / PT','Acquired Company',\
                        'Direct Supervisor - Emp ID','Direct Supervisor - Legal Name','Direct Supervisor - Work Email','Direct Supervisor - User ID',\
                        'FTE %','Job Code','Job Profile','Job Family Group','Job Family','Job Category','Job Level','Management Level','EEO Job Group',\
                        'EEO Job Family','Compensation Frequency','Pay Rate Type','Standard Hours','FLSA','Local Currency','FX Rate','Base Annualized (Local)',\
                        'Target Bonus %','Target ABP Exception','Sales Incentive Guarantee','Work Location - Office','Work Location - Workspace',\
                        'Work Location - City','Work Location - State','Work Location - Postal Code','Work Location - Country','Comp Grade','Comp Grade Profile',\
                        'Direct Headcount','CEO','L2','L3','L4','L5','L6','L7','L8','L9','L10','L2 Org Name','L3 Org Name','gender','ethnicity','marital_status','military_status','benefit_program']

writer = pandas.ExcelWriter('output.xlsx')
final_report.to_excel(writer,'Sheet1', index=False)
writer.save()
