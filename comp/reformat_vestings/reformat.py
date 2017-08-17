import pandas as pd
import numpy
import datetime

# FOR REFERENCE: counts by group is df.groupby('colname').size()

################################################################################
##### Setup options

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

################################################################################

# setwd = "/Users/josefnunez/workforce/"

# Load equity report. Assumes report is in current working directory.
equity_filepath, equity_sheet = "equity.xlsx", "Sheet1"
equity = pd.ExcelFile(equity_filepath).parse(equity_sheet)
equity.columns = ['product_type_name','plan_id','product_id','grant_type','grant_date','grant_sequence_number',\
                  'client_grant_id_or_grant_number','grant_custom_field_1','grant_custom_field_2',\
                  'grant_custom_field_3','grant_custom_field_4','legacy_yahoo_grant_type','participant_name',\
                  'participant_id','employee_id','active_indicator','division_code','qty_granted','qty_outstanding',\
                  'vesting_date_1','qty_vesting_1','vesting_date_2','qty_vesting_2','vesting_date_3','qty_vesting_3',\
                  'vesting_date_4','qty_vesting_4','vesting_date_5','qty_vesting_5','vesting_date_6','qty_vesting_6',\
                  'vesting_date_7','qty_vesting_7','vesting_date_8','qty_vesting_8','vesting_date_9','qty_vesting_9',\
                  'vesting_date_10','qty_vesting_10','vesting_date_11','qty_vesting_11','vesting_date_12',\
                  'qty_vesting_12','vesting_date_13','qty_vesting_13','vesting_date_14','qty_vesting_14',\
                  'vesting_date_15','qty_vesting_15','vesting_date_16','qty_vesting_16','vesting_date_17',\
                  'qty_vesting_17','vesting_date_18','qty_vesting_18','vesting_date_19','qty_vesting_19',\
                  'vesting_date_20','qty_vesting_20','vesting_date_21','qty_vesting_21','vesting_date_22',\
                  'qty_vesting_22','vesting_date_23','qty_vesting_23','vesting_date_24','qty_vesting_24',\
                  'vesting_date_25','qty_vesting_25','vesting_date_26','qty_vesting_26','vesting_date_27',\
                  'qty_vesting_27','vesting_date_28','qty_vesting_28','vesting_date_29','qty_vesting_29',\
                  'vesting_date_30','qty_vesting_30','vesting_date_31','qty_vesting_31','vesting_date_32',\
                  'qty_vesting_32','vesting_date_33','qty_vesting_33','vesting_date_34','qty_vesting_34',\
                  'vesting_date_35','qty_vesting_35','vesting_date_36','qty_vesting_36','vesting_date_37',\
                  'qty_vesting_37','vesting_date_38','qty_vesting_38','vesting_date_39','qty_vesting_39',\
                  'vesting_date_40','qty_vesting_40','vesting_date_41','qty_vesting_41','vesting_date_42',\
                  'qty_vesting_42','vesting_date_43','qty_vesting_43','vesting_date_44','qty_vesting_44',\
                  'vesting_date_45','qty_vesting_45','vesting_date_46','qty_vesting_46','vesting_date_47',\
                  'qty_vesting_47','vesting_date_48','qty_vesting_48','vesting_date_49','qty_vesting_49',\
                  'vesting_date_50','qty_vesting_50','vesting_date_51','qty_vesting_51','vesting_date_52',\
                  'qty_vesting_52','vesting_date_53','qty_vesting_53','vesting_date_54','qty_vesting_54',\
                  'vesting_date_55','qty_vesting_55','vesting_date_56','qty_vesting_56','vesting_date_57',\
                  'qty_vesting_57','vesting_date_58','qty_vesting_58','vesting_date_59','qty_vesting_59',\
                  'vesting_date_60','qty_vesting_60','vesting_date_61','qty_vesting_61','vesting_date_62',\
                  'qty_vesting_62','vesting_date_63','qty_vesting_63','vesting_date_64','qty_vesting_64',\
                  'vesting_date_65','qty_vesting_65','vesting_date_66','qty_vesting_66','vesting_date_67',\
                  'qty_vesting_67','vesting_date_68','qty_vesting_68','vesting_date_69','qty_vesting_69',\
                  'vesting_date_70','qty_vesting_70','vesting_date_71','qty_vesting_71','vesting_date_72','qty_vesting_72']

# Remove extraneous columns
equity.drop(['product_type_name','plan_id','product_id','grant_type','grant_sequence_number','grant_custom_field_1',\
             'grant_custom_field_2','grant_custom_field_3','grant_custom_field_4','participant_name','participant_id',\
             'active_indicator','division_code','qty_granted','qty_outstanding'],\
             axis=1, inplace=True)

# Get number of row and columns in data frame
# Note: df.shape returns tuple of (number of rows, number of columns) for the data frame
NUM_ROWS_EQUITY_DF, NUM_COLS_EQUITY_DF = equity.shape

# Set max number of vesting events for a grant in the report
MAX_NUM_VESTING_EVENTS = 72

# Fields to add in merge: ['vest_date','shares_vested_year','shares_vested']

### Important Note for Future Python Pandas Use ###
# Trap: When inserting data from an indexed pandas object, only items from the indexed object that 
# have a corresponding index in the DataFrame will be added. The receiving DataFrame is not extended to
# accommodate the new series. 
# Solution: Using DataFrame VALUES attribute to return a numpy array -- a non-indexed object.
# 
# This problem with solution used in code below.

# Create empty dataframe to store final table. INDEX is number of rows. COLUMNS is number of columns.
output = pd.DataFrame(index=range(NUM_ROWS_EQUITY_DF*MAX_NUM_VESTING_EVENTS),\
                      columns=['employee_id','grant_type','grant_number','grant_date','vest_date','shares_vested'])

# Fill-in dataframe with vesting data
static_cols = ['employee_id','grant_type','client_grant_id_or_grant_number','grant_date'] # Static cols per grant
for i in range(MAX_NUM_VESTING_EVENTS):
    first_entry_row = i*NUM_ROWS_EQUITY_DF
    last_entry_row = first_entry_row + NUM_ROWS_EQUITY_DF
    curr_vesting_date, curr_qty_vesting = 'vesting_date_' + str(i+1), 'qty_vesting_' + str(i+1)
    output.iloc[first_entry_row:last_entry_row, :] = equity.loc[:, static_cols+[curr_vesting_date, curr_qty_vesting]].values

# Cleanup resulting dataframe
output = output.loc[output.loc[:,'vest_date'].notnull()] # Remove empty vest events
output.sort_values(['employee_id','grant_number','vest_date'], inplace=True)
output.reset_index(drop=True, inplace=True)

# cwd_sens_cols = ['worker_type','emp_type','eeid','legal_name','gender','ethnicity','mgr_eeid','mgr_legal_name','mgr_email',\
#                  'userid','last_hire_date','original_hire_date','active_status','ft_or_pt','fte_pct',\
#                  'std_hrs','email','acquired_company','job_code','job_profile','job_family_group',\
#                  'job_family','job_level','job_category','mgmt_level','comp_grade','comp_grade_profile',\
#                  'pay_rate_type','flsa','base_annualized_local','local_currency','fx_rate','base_annualized_usd',\
#                  'bonus_plan','target_bonus_pct','target_bonus_amt_local','target_bonus_amt_usd',\
#                  'ttc_annualized_local','ttc_annualized_usd','wfh_flag','work_office','work_city',\
#                  'work_state','work_country','work_region','is_ppl_mgr','layer','CEO_eeid','CEO_name','L2_eeid','L2_name',\
#                  'L3_eeid','L3_name','L4_eeid','L4_name','L5_eeid','L5_name','L6_eeid','L6_name',\
#                  'L7_eeid','L7_name','L8_eeid','L8_name','L9_eeid','L9_name','L10_eeid','L10_name',\
#                  'L2_org_name','L3_org_name','L4_org_name','L2_or_L3_org_name','last_day_of_work','term_date']
# cwd_sens = oath.loc[:, cwd_sens_cols]
# writer_cwd_sens = pandas.ExcelWriter('outputs/Oath Current Employee and Contingent Worker Details - Highly Sensitive ' + DATETIMESTAMP + '.xlsx')
# cwd_sens.to_excel(writer_cwd_sens, 'Sheet1', index=False)
# writer_cwd_sens.save()



