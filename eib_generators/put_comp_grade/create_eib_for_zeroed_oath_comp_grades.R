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
core_comp_structures = c("CA","CB","CC","CE")
sales_comp_structures = c("SA","SB","SC","SE")
tech_comp_structures = c("TA","TB","TC","TE")
comp_structures = c(core_comp_grades, sales_comp_grades, tech_comp_grades)

# List out levels for Non-Exec and Exec comp grades
non_exec = c(62,64,66,
             72,74,76,
             82,84,86,88)
exec = c(92,94,96)

NUM_NON_EXEC = length(non_exec)
NUM_EXEC = length(exec)

core_comp_grades = create_comp_grades(core_comp_structures)
sales_comp_grades = create_comp_grades(sales_comp_structures)
tech_comp_grades = create_comp_grades(tech_comp_structures)

# QUESTION: Should we create Paris and Dublin comp grade profiles? (currently listed below)
# We currently do NOT have France employees outside of Paris, nor Ireland employees ouside of Dublin.
# Do we see the company expanding in these countries, which would warrant these additional comp grade profiles.
comp_grade_profiles = c("Australia","Sydney","Belgium","Brazil","Canada","Denmark","France","Paris","Germany","Munich",
                        "Hong Kong","India","Ireland","Dublin","Israel","Italy","Japan","Netherlands","New Zealand",
                        "Norway","Singapore","South Korea","Spain","Sweden","Taiwan","United Kingdom",
                        "United States - Tier 1","United States - Tier 2","United States - Tier 3","United States - Tier 4")

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

create_row_for_grade = function(grade, sskey, addonly, comp_element="Standard_Base_Pay")
{
  # Create empty dataframe to store row values. 
  # Number of rows = Number of Unique Geo Regions
  # Number of columns = Number of Required Columns to Fill
  entry = data.frame(matrix("",nrow=NUM_COMP_GRADE_PROFILES,ncol=NUM_REQUIRED_EIB_FIELDS),stringsAsFactors=FALSE)
  names(entry) = READABLE_COL_HEADERS_FOR_FN

  entry[["sskey"]][1] = sskey
  entry[["grade_addonly"]][1] = addonly
  entry[["grade"]][1] = grade
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
  entry[["default_currency"]][1] = "USD"
  entry[["default_frequency"]][1] = "Annual"
  entry[["default_allow_overrride"]][1] = "N"
  entry[["cgp_row_id"]] = 1:NUM_COMP_GRADE_PROFILES
  entry[["cgp_delete"]] = "N"
  entry[["cgp_id"]] = comp_grade_profiles
  entry[["cgp_effective_date"]] = "1900-01-01"
  entry[["cgp_name"]] = comp_grade_profiles
  entry[["cgp_description"]] = ""
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

create_eib_for_comp_structure = function()

x = create_row_for_grade("CA10","1","Y","Standard_Base_Pay")



write.xlsx2(x, "TEST.xlsx", sheetName="Sheet1",
            col.names=TRUE, row.names=FALSE, append=FALSE)

for (grade in comp_grades)
{
  create_row_for_grade(grade,"1","Y","Standard_Base_Pay")
}





