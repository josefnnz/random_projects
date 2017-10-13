import pandas
import numpy
import datetime

# FOR REFERENCE: counts by group is df.groupby('colname').size()

################################################################################
##### Setup options

pandas.set_option('display.max_columns', None)
pandas.set_option('display.max_rows', None)

################################################################################
def create_row_for_grade(grade, sskey, addonly, comp_element="Standard_Base_Pay"):
    # Extract comp grade profiles for given comp grade
    profiles = ranges.loc[ranges.loc[:,'grade']==grade, :]

    # Create empty dataframe to store row values. 
    # 36 rows for 36 unique georegions. 
    # 43 cols for 43 required columns to fill
    NUM_ROWS, NUM_COLS = 36, 43
    row = pandas.DataFrame(index=range(NUM_ROWS), columns=range(NUM_COLS))
    row.columns = ['sskey','grade_addonly','grade','grade_id','grade_effective_date','grade_name','grade_description',\
                   'grade_comp_element','default_num_segments','default_min','default_mid','default_max','default_spread','default_s1_top',\
                   'default_s2_top','default_s3_top','default_s4_top','default_currency','default_frequency','default_allow_overrride','cgp_row_id',\
                   'cgp_delete','cgp_id','cgp_effective_date','cgp_name','cgp_description','cgp_comp_element','cgp_eligibility_rule','cgp_inactive',\
                   'cgp_num_segments','cgp_min','cgp_mid','cgp_max','cgp_spread','cgp_s1_top','cgp_s2_top','cgp_s3_top','cgp_s4_top','cgp_currency',\
                   'cgp_frequency','cgp_allow_overrride','assign_first_step_during_comp_proposal','grade_inactive']
    row.loc[:, 'sskey'] = sskey
    row.loc[0, 'grade_addonly'] = addonly
    row.loc[0, 'grade'] = grade
    row.loc[0, 'grade_id'] = grade
    row.loc[0, 'grade_effective_date'] = "1900-01-01"
    row.loc[0, 'grade_name'] = grade
    row.loc[0, 'grade_comp_element'] = comp_element
    row.loc[0, 'default_num_segments'] = 4
    row.loc[0, 'default_min'] = 0
    row.loc[0, 'default_mid'] = 0
    row.loc[0, 'default_max'] = 0
    row.loc[0, 'default_spread'] = 0
    row.loc[0, 'default_s1_top'] = 0
    row.loc[0, 'default_s2_top'] = 0
    row.loc[0, 'default_s3_top'] = 0
    row.loc[0, 'default_s4_top'] = 0
    row.loc[0, 'default_currency'] = "USD"
    row.loc[0, 'default_frequency'] = "Annual"
    row.loc[0, 'default_allow_overrride'] = "N"
    row.loc[:, 'cgp_row_id'] = range(1, NUM_ROWS+1)
    row.loc[:, 'cgp_delete'] = "N"
    row.loc[:, 'cgp_id'] = profiles.loc[:, 'cgp'].values
    row.loc[:, 'cgp_effective_date'] = "1900-01-01"
    row.loc[:, 'cgp_name'] = profiles.loc[:, 'cgp'].values
    row.loc[:, 'cgp_description'] = ""
    row.loc[:, 'cgp_comp_element'] = comp_element
    row.loc[:, 'cgp_eligibility_rule'] = profiles.loc[:, 'country'].values
    row.loc[:, 'cgp_inactive'] = "N"
    row.loc[:, 'cgp_num_segments'] = 4
    row.loc[:, 'cgp_min'] = profiles.loc[:, 'cgp_min'].values
    row.loc[:, 'cgp_mid'] = profiles.loc[:, 'cgp_mid'].values
    row.loc[:, 'cgp_max'] = profiles.loc[:, 'cgp_max'].values
    row.loc[:, 'cgp_spread'] = profiles.loc[:, 'cgp_spread'].values
    row.loc[:, 'cgp_s1_top'] = profiles.loc[:, 'cgp_s1_top'].values
    row.loc[:, 'cgp_s2_top'] = profiles.loc[:, 'cgp_s2_top'].values
    row.loc[:, 'cgp_s3_top'] = profiles.loc[:, 'cgp_s3_top'].values
    row.loc[:, 'cgp_s4_top'] = profiles.loc[:, 'cgp_s4_top'].values
    row.loc[:, 'cgp_currency'] = profiles.loc[:, 'cgp_currency'].values
    row.loc[:, 'cgp_frequency'] = "Annual"
    row.loc[:, 'cgp_allow_overrride'] = "N"
    row.loc[0, 'assign_first_step_during_comp_proposal'] = "N"
    row.loc[0, 'grade_inactive'] = "N"
    return row

################################################################################

setwd = "/Users/josefnunez/random_projects/eib_generators/put_comp_grade/"

# load Oath workforce roster, i.e. PeopleSoft report
ranges_filepath, ranges_sheet = setwd+"Range Table - DRAFT (v1 10-07-17).xlsx", "Sheet1"
ranges = pandas.ExcelFile(ranges_filepath).parse(ranges_sheet)
ranges.columns = ['grade','cgp','georegion','country','gradetype','jobcategory','cgp_currency','compphilosophy','base_or_ote','cgp_min','cgp_mid','cgp_max']
ranges.sort_values(by=['grade','cgp'], ascending=[True,True], inplace=True)
ranges.loc[:, 'cgp_spread'] = ranges.loc[:, 'cgp_max'] - ranges.loc[:, 'cgp_min']
ranges.loc[:, 'cgp_s1_top'] = (ranges.loc[:, 'cgp_mid'] + ranges.loc[:, 'cgp_min']) / 2
ranges.loc[:, 'cgp_s2_top'] = ranges.loc[:, 'cgp_mid']
ranges.loc[:, 'cgp_s3_top'] = (ranges.loc[:, 'cgp_mid'] + ranges.loc[:, 'cgp_max']) / 2
ranges.loc[:, 'cgp_s4_top'] = ranges.loc[:, 'cgp_max']

unique_grades = ranges.loc[ranges.loc[:, 'grade'].notnull(), 'grade']
unique_grades.drop_duplicates(inplace=True)
# unique_grades = unique_grades.reindex_axis(labels=None, axis=0)

unique_cgps = ranges.loc[ranges.loc[:, 'cgp'].notnull(), 'cgp']
unique_cgps.drop_duplicates(inplace=True)
# unique_cgps = unique_cgps.reindex_axis(labels=None, axis=0)

NUM_UNIQUE_GRADES, NUM_UNIQUE_CGPS, NUM_REQUIRED_COLS = unique_grades.shape[0], unique_cgps.shape[0], 43
indices = pandas.DataFrame(range(NUM_UNIQUE_GRADES))
eib = pandas.DataFrame(index=range(NUM_UNIQUE_GRADES*NUM_UNIQUE_CGPS), columns=range(NUM_REQUIRED_COLS))
i = 1
for g in unique_grades:
    new_row = create_row_for_grade(g, i+1, "Y", comp_element="Standard_Base_Pay")
    shift = i+i*NUM_UNIQUE_CGPS
    eib.iloc[range(0+shift, NUM_UNIQUE_CGPS+shift), :] = new_row
    i += 1

writer = pandas.ExcelWriter('test.xlsx')
eib.to_excel(writer, 'Sheet1', index=False)
writer.save()


# 1. Create empty dataframe to hold entire EIB
# 2. Use function to generate EIB row for one comp grade
# 3. Insert comp grade row into overall dataframe
# 4. Write to excel














