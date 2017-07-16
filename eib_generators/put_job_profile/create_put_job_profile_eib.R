setwd("/Users/josefnunez/jobcodes/")

ALL_EIB_FIELDS = c("Fields","Spreadsheet Key*","Add Only","Job Profile","Job Code","Effective Date","Inactive","Job Title",
                   "Include Job Code in Name","Job Profile Private Title","Job Profile Summary","Job Description","Additional Job Description",
                   "Work Shift Required","Public Job","Management Level","Job Category","Job Level","Row ID*","Delete","Job Family*","Row ID*",
                   "Delete","Company Insider Type*","Referral Payment Plan","Critical Job","Difficulty to Fill","Restrict to Country+","Row ID*",
                   "Delete","Job Classification*","Row ID*","Delete","Country*","Pay Rate Type","Row ID*","Delete","Location Context*",
                   "ISO 3166 1 Alpha 2 Code","Job Exempt","Row ID*","Workers Compensation Code","Delete","Row ID*","Responsibility Description*",
                   "Required","Delete","Row ID*","Worker Experience*","Work Experience Rating","Required","Delete","Row ID*","Degree","Field Of Study",
                   "Required","Delete","Row ID*","Language*","Row ID*","Language Ability Type*","Language Proficiency","Required","Delete","Row ID*",
                   "Competency*","Proficiency Rating","Required","Delete","Row ID*","Country","Certification","Name","Issuer","Required","Row ID*",
                   "Specialty","Subspecialty+","Delete","Row ID*","Training Name","Description","Training Type","Required","Delete","Row ID*","Skill",
                   "Name","Required","Row ID*","Field","Integration Document Name","Value","Compensation Grade","Compensation Grade Profile+",
                   "Requirement Option","Union+","Questionnaire for Internal Career Site","Secondary Questionnaire for Internal Career Site",
                   "Questionnaire for External Career Sites","Secondary Questionnaire for External Career Sites","Requirement Option",
                   "Allowed Student Award Sources+")

REQUIRED_EIB_FIELDS = c("Spreadsheet Key*","Add Only","Job Code","Effective Date","Inactive","Job Title","Include Job Code in Name","Management Level",
                        "Job Category","Job Level","Row ID*","Delete","Job Family*","Row ID*","Delete","Job Classification*","Row ID*","Delete",
                        "Country*","Pay Rate Type","Row ID*","Delete","Location Context*","Job Exempt","Compensation Grade")

INDICES_REQUIRED_EIB_FIELDS = c(2,3,5,6,7,8,9,16,17,18,19,20,21,29,30,31,32,33,34,35,36,37,38,40,94)

READABLE_COL_HEADERS_FOR_FN = c("sskey","addonly","code","effdate","inactive","title","codeinname","mgmtlevel","jobcat","level",
                                "rowjobfam","deletejobfam","jobfam","roweeoclass","deleteeeoclass","eeoclass","rowcountry",
                                "deletecountry","country","countrypayratetype","rowflsa","deleteflsa","flsacountry","flsa","compgrade")

JOB_COUNTRIES = c("US","AU","BE","BR","CA","CN","DK","FR","DE","HK","IN","ID","IE","IL","IT","JP","KR","LU","NL","NZ","NO","RO","SG","ES","SE","CH","TW","TH","GB","VN")

NUM_ALL_EIB_FIELDS = length(ALL_EIB_FIELDS)

NUM_REQUIRED_EIB_FIELDS = length(REQUIRED_EIB_FIELDS)

NUM_JOB_COUNTRIES = length(JOB_COUNTRIES)

# Column Definitions: EIB File Section > Field - EIB Column Name
# All > sskey - Spreadsheet Key | addonly - Add Only
# Job Profile Date (All) > code - Job Code | effdate - Effective Date | inactive - Inactive | title - Job Title | 
#   codeinname - Include Job Code in Name | mgmtlevel - Management Level | jobcat - Job Category | level - Job Level
# Job Family Data > rowjobfam - Row ID | deletejobfam - Delete | jobfam - Job Family
# Job Classifcation Data > roweeoclass - Row ID | deleteeeoclass - Delete | eeoclass - Job Classification
# Job Profile Pay Rate Data > rowcountry - Row ID | deletecountry - Delete | country - Country | countrypayratetype - Pay Rate Type
# Job Profile Exempt Data > rowflsa - Row ID | deleteflsa - Delete | flsacountry - Location Context | flsa - Job Exempt
# Job Profile Compensation Data > compgrade - Compensation Grade
create_eib_entry = function(key, job) 
{
  entry = data.frame(matrix("",nrow=NUM_JOB_COUNTRIES,ncol=NUM_REQUIRED_EIB_FIELDS),stringsAsFactors=FALSE)
  names(entry) = READABLE_COL_HEADERS_FOR_FN
  
  entry[["sskey"]][1:30] = rep(as.character(key),times=30)
  entry[["addonly"]][1] = "Y"
  entry[["code"]][1] = job[["code"]]
  entry[["effdate"]][1] = "1900-01-01"
  entry[["inactive"]][1] = "Y"
  entry[["title"]][1] = job[["title"]]
  entry[["codeinname"]][1] = "N"
  entry[["mgmtlevel"]][1] = job[["mgmtlevel"]]
  entry[["jobcat"]][1] = job[["jobcat"]]
  entry[["level"]][1] = job[["level"]]
  entry[["rowjobfam"]][1] = "1"
  entry[["deletejobfam"]][1] = "N"
  entry[["jobfam"]][1] = job[["jobfam"]]
  entry[["roweeoclass"]][1:2] = c("1","2")
  entry[["deleteeeoclass"]][1:2] = c("N","N")
  entry[["eeoclass"]][1] = job[["eeojobgroup"]]
  entry[["eeoclass"]][2] = job[["eeojobcategory"]]
  entry[["rowcountry"]][1:30] = as.character(1:30)
  entry[["deletecountry"]][1:30] = rep("N",times=30)
  entry[["country"]][1:30] = JOB_COUNTRIES
  entry[["countrypayratetype"]][1:30] = c(job[["payratetype"]],rep("Salaried",times=29))
  entry[["rowflsa"]][1] = "1"
  entry[["deleteflsa"]][1] = "N"
  entry[["flsacountry"]][1] = "US"
  entry[["flsa"]][1] = job[["isexempt"]]
  entry[["compgrade"]][1] = job[["compgrade"]]
  return(entry)
}

newcojobs = read.csv("newco.csv",stringsAsFactors=FALSE)
names(newcojobs) = c("code","title","jobfamgrp","jobfam","levelorder","jobcat","level","mgmtlevel","eeojobcategory",
                     "eeojobgroup","payratetype","isexempt","compgrade")

NUM_JOB_PROFILES = nrow(newcojobs)

eib = as.data.frame(matrix("",nrow=NUM_JOB_PROFILES*NUM_JOB_COUNTRIES,ncol=length(READABLE_COL_HEADERS_FOR_FN)),stringsAsFactors=FALSE)
names(eib) = READABLE_COL_HEADERS_FOR_FN
for (i in 1:NUM_JOB_PROFILES) {
  print(i)
  eib[(1:30)+(30*(i-1)),] = create_eib_entry(i,newcojobs[i,])
}

fulleib = data.frame(matrix("",nrow=NUM_JOB_PROFILES*NUM_JOB_COUNTRIES,ncol=NUM_ALL_EIB_FIELDS),stringsAsFactors=FALSE)
fulleib[,INDICES_REQUIRED_EIB_FIELDS] = eib
names(fulleib) = ALL_EIB_FIELDS

DATETIMESTAMP = format(Sys.time(),"%Y-%m-%d %H_%M PDT")
write.csv(fulleib,file=paste0("EIB - Put_Job_Profile - ",DATETIMESTAMP,".csv"),row.names=FALSE)
