################################################################################################
# create_eib_for_zeroed_oath_comp_grades.R
# Description: Create EIB for Oath comp grades, zeroed out.
# Details: NO FILES ARE IMPORTED IN THIS SCRIPT. The Comp Team provided a list of unique
# Oath Comp Grades and Oath Comp Grade Profiles. The unique lists of CGs / CGPs are sufficient
# to build a Put_Comp_Grade EIB with zeroed ranges.
################################################################################################

require(xlsx)

# Set working directory where the EIBs will be outputted
setwd("/Users/josefnunez/random_projects/eib_generators/put_comp_grade/output")

# CE, SE, TE comp grades below are EXECUTIVE comp grades. 
# The EXECUTIVE comp grades only pertain to levels 92, 94, 96
# Otherwise, the NON-EXECUTIVE comp grades pertain to levels 62, 64, 66, 72, 74, 76, 82, 84, 86, 88
comp_structures = c("CA","CB","CC","CE",
                    "SA","SB","SC","SE", 
                    "TA","TB","TC","TE")

# List out levels for Non-Exec and Exec comp grades. These will be concatenated with the
# Comp Strucutres listed above to create the Comp Grade names.
# NON-EXECUTIVE: 62, 64, 66, 72, 74, 76, 82, 84, 86, 88
# EXECUTIVE: 92, 94, 96
non_exec = c(62,64,66,
             72,74,76,
             82,84,86,88)
exec = c(92,94,96)

NUM_NON_EXEC = length(non_exec)
NUM_EXEC = length(exec)

# QUESTION: Should we create Paris and Dublin comp grade profiles? (currently listed below)
# We currently do NOT have France employees outside of Paris, nor Ireland employees ouside of Dublin.
# Do we see the company expanding in these countries, which would warrant these additional comp grade profiles.
comp_grade_profiles = c("Australia","Sydney","Belgium","Brazil","Canada","Denmark","France","Paris","Germany","Munich",
                        "Hong Kong","India","Ireland","Dublin","Israel","Italy","Japan","Netherlands","New Zealand",
                        "Norway","Singapore","South Korea","Spain","Sweden","Taiwan","United Kingdom",
                        "United States - Tier 1","United States - Tier 2","United States - Tier 3","United States - Tier 4")

comp_grade_profile_descriptions = c("AU","AU","BE","BR","CA","DK","FR","FR","DE","DE","HK","IN","IE","IE","IL","IT","JP",
                                    "NL","NZ","NO","SG","KR","ES","SE","TW","UK","US","US","US","US")

comp_grade_profiles_currencies = c("AUD","AUD","EUR","BRL","CAD","DKK","EUR","EUR","EUR","EUR","HKD","INR","EUR","EUR",
                                   "ILS","EUR","JPY","EUR","NZD","NOK","SGD","KRW","EUR","SEK","TWD","GBP","USD","USD",
                                   "USD","USD")

NUM_COMP_GRADE_PROFILES = length(comp_grade_profiles)

# Store columns names used in create_row_for_grade function
READABLE_COL_HEADERS_FOR_FN = c("sskey","grade_addonly","grade","grade_id","grade_effective_date","grade_name","grade_description",
                                "grade_comp_element","default_num_segments","default_min","default_mid","default_max","default_spread","default_s1_top",
                                "default_s2_top","default_s3_top","default_s4_top","default_currency","default_frequency","default_allow_overrride","cgp_row_id",
                                "cgp_delete","cgp_id","cgp_effective_date","cgp_name","cgp_description","cgp_comp_element","cgp_eligibility_rule","cgp_inactive",
                                "cgp_num_segments","cgp_min","cgp_mid","cgp_max","cgp_spread","cgp_s1_top","cgp_s2_top","cgp_s3_top","cgp_s4_top","cgp_currency",
                                "cgp_frequency","cgp_allow_overrride","assign_first_step_during_comp_proposal","grade_inactive")

NUM_REQUIRED_EIB_FIELDS = length(READABLE_COL_HEADERS_FOR_FN)

# Store ALL columns on the eib
ALL_EIB_COLUMNS = c("area_fields","sskey","grade_addonly","grade","grade_id","grade_effective_date","grade_name",
                    "grade_description","grade_comp_element","compensation_grade_data__eligibility_rule",
                    "default_num_segments","default_min","default_mid","default_max","default_spread","default_s1_top",
                    "default_s2_top","default_s3_top","default_s4_top","compensation_pay_range__segment_5_top",
                    "default_currency","default_frequency","compensation_pay_range__salary_plan",
                    "default_allow_overrride","compensation_step__row_id","compensation_step__delete",
                    "compensation_step__compensation_step","compensation_step_data__compensation_step_id",
                    "compensation_step_data__sequence","compensation_step_data__name","compensation_step_data__amount",
                    "compensation_step_data__interval","compensation_step_data__period",
                    "compensation_step_data__progression_rule","cgp_row_id","cgp_delete",
                    "compensation_grade_profile__compensation_grade_profile","cgp_id","cgp_effective_date","cgp_name",
                    "cgp_description","cgp_comp_element","cgp_eligibility_rule","cgp_inactive","cgp_num_segments",
                    "cgp_min","cgp_mid","cgp_max","cgp_spread","cgp_s1_top","cgp_s2_top","cgp_s3_top","cgp_s4_top",
                    "compensation_pay_range_data__segment_5_top","cgp_currency","cgp_frequency",
                    "compensation_pay_range_data__salary_plan","cgp_allow_overrride","compensation_step__row_id",
                    "compensation_step__delete","compensation_step__compensation_step",
                    "compensation_step_data__compensation_step_id","compensation_step_data__sequence",
                    "compensation_step_data__name","compensation_step_data__amount","compensation_step_data__interval",
                    "compensation_step_data__period","compensation_step_data__progression_rule",
                    "compensation_grade_profile_data__setup_security_segment","alternative_pay_range__row_id",
                    "alternative_pay_range__delete","alternate_pay_range_data__compensation_basis",
                    "alternate_pay_range_data__minimum","alternate_pay_range_data__midpoint",
                    "alternate_pay_range_data__maximum","alternate_pay_range_data__spread",
                    "alternate_pay_range_data__segment_1_top","alternate_pay_range_data__segment_2_top",
                    "alternate_pay_range_data__segment_3_top","alternate_pay_range_data__segment_4_top",
                    "alternate_pay_range_data__segment_5_top","compensation_grade_data__setup_security_segment",
                    "assign_first_step_during_comp_proposal","compensation_grade_data__pay_level","grade_inactive",
                    "alternative_pay_range__row_id","alternative_pay_range__delete",
                    "alternate_pay_range_data__compensation_basis","alternate_pay_range_data__minimum",
                    "alternate_pay_range_data__midpoint","alternate_pay_range_data__maximum",
                    "alternate_pay_range_data__spread","alternate_pay_range_data__segment_1_top",
                    "alternate_pay_range_data__segment_2_top","alternate_pay_range_data__segment_3_top",
                    "alternate_pay_range_data__segment_4_top","alternate_pay_range_data__segment_5_top")

NUM_ALL_EIB_FIELDS = length(ALL_EIB_COLUMNS)

create_row_for_grade = function(grade, sskey, addonly, comp_element="Standard_Base_Pay")
{
  # Create empty dataframe to store row values. 
  # Number of rows = Number of Unique Geo Regions
  # Number of columns = Number of Required Columns to Fill
  entry = data.frame(matrix("",nrow=NUM_COMP_GRADE_PROFILES,ncol=NUM_ALL_EIB_FIELDS),stringsAsFactors=FALSE)
  names(entry) = ALL_EIB_COLUMNS

  entry[["sskey"]] = sskey
  entry[["grade_addonly"]][1] = addonly
  # entry[["grade"]][1] = grade
  entry[["grade_id"]][1] = grade
  entry[["grade_effective_date"]][1] = "1900-01-01"
  entry[["grade_name"]][1] = grade
  entry[["grade_comp_element"]][1] = comp_element
  entry[["default_num_segments"]][1] = 4
  entry[["default_min"]][1] = 0
  entry[["default_mid"]][1] = 0
  entry[["default_max"]][1] = 0
  entry[["default_spread"]][1] = 0
  entry[["default_s1_top"]][1] = 0
  entry[["default_s2_top"]][1] = 0
  entry[["default_s3_top"]][1] = 0
  entry[["default_s4_top"]][1] = 0
  # entry[["default_s5_top"]][1] = 0
  entry[["default_currency"]][1] = "USD"
  entry[["default_frequency"]][1] = "Annual"
  entry[["default_allow_overrride"]][1] = "N"
  entry[["cgp_row_id"]] = 1:NUM_COMP_GRADE_PROFILES
  entry[["cgp_delete"]] = "N"
  entry[["cgp_id"]] = paste0(paste0(grade,"-"),comp_grade_profiles)
  entry[["cgp_effective_date"]] = "1900-01-01"
  entry[["cgp_name"]] = comp_grade_profiles
  entry[["cgp_description"]] = comp_grade_profile_descriptions
  entry[["cgp_comp_element"]] = comp_element
  entry[["cgp_eligibility_rule"]] = comp_grade_profiles
  entry[["cgp_inactive"]] = "N"
  entry[["cgp_num_segments"]] = 4
  entry[["cgp_min"]] = 0
  entry[["cgp_mid"]] = 0
  entry[["cgp_max"]] = 0
  entry[["cgp_spread"]] = 0
  entry[["cgp_s1_top"]] = 0
  entry[["cgp_s2_top"]] = 0
  entry[["cgp_s3_top"]] = 0
  entry[["cgp_s4_top"]] = 0
  entry[["cgp_currency"]] = comp_grade_profiles_currencies
  entry[["cgp_frequency"]] = "Annual"
  entry[["cgp_allow_overrride"]] = "N"
  entry[["assign_first_step_during_comp_proposal"]] = "N"
  entry[["grade_inactive"]] = "N"
  return(entry)
}

create_comp_grades = function(structs)
{
  grades = c()
  for (s in structs)
  {
    lvls = non_exec
    if(grepl("E", s, fixed=TRUE))
    {
      lvls = exec
    }
    grades = c(grades, paste0(s, lvls))
  }
  return(grades)
}

create_eib_for_comp_structure = function(struct)
{
  comp_grades = create_comp_grades(struct)
  NUM_COMP_GRADES = length(comp_grades)
  
  # Create new empty data frame to hold EIB data
  eib = data.frame(matrix("",nrow=NUM_COMP_GRADES*NUM_COMP_GRADE_PROFILES,ncol=NUM_ALL_EIB_FIELDS),
                   stringsAsFactors=FALSE)
  names(eib) = ALL_EIB_COLUMNS
  for (i in 1:NUM_COMP_GRADES)
  {
    grade = comp_grades[i]
    eib[1:NUM_COMP_GRADE_PROFILES+((i-1)*NUM_COMP_GRADE_PROFILES), ] = create_row_for_grade(grade, i, "Y", "Standard_Base_Pay")
  }
  return(eib)
}

create_unique_sskeys_for_consolidated_eib = function(eib)
{
  keys = eib[["sskey"]]
  k = 1
  for (i in 1:length(keys))
  {
    if (keys[i] != "")
    {
      keys[i] = k
      k = k + 1
    }
  }
  eib[["sskey"]] = keys
  return(eib)
}

create_consolidated_eib = function(eibs)
{
  consolidated_eib = do.call(rbind, eibs)
  consolidated_eib = create_unique_sskeys_for_consolidated_eib(consolidated_eib)
  return(consolidated_eib)
}

all_eib_data = list()
for (s in comp_structures)
{
  all_eib_data[[s]] = create_eib_for_comp_structure(s)
}

eib_with_all_structures = create_consolidated_eib(all_eib_data)

DATETIMESTAMP = format(Sys.time(),"%Y-%m-%d %H_%M PDT")
filename = paste0("EIB - Put_Comp_Grade - ",DATETIMESTAMP,".xlsx")
write.xlsx2(eib_with_all_structures, file=filename, sheetName="Compensation Grade", col.names=TRUE, row.names=FALSE, append=FALSE, showNA=FALSE)






