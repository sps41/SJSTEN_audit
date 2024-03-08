install.packages("xlsx")
library("readxl")
options(readxl.show_progress = FALSE)
setwd("D:/SJSTEN Audit")

rm(list = ls())
dir <- paste(getwd(),"/Data",sep="")
files <- list.files(dir,full.names = T)
###Read first file and create data frames
colnames <- read_excel(files[1],range = "Core Qs!A2:AL4",col_names=as.character(1:38))
actualColNames <- as.character(c(colnames[1,1:3],colnames[2,4:11],colnames[3,12:38]))
AddColNames <- read_excel(files[1],range = "Additional Qs!B4:AC4",col_names=as.character(1:28))
AddColNamesreal <- as.character(c("Patient",AddColNames[1,]))
ClinNames <- as.character(read_excel(files[1],range = "Core Qs!C11:O11",col_names=as.character(1:13))[1,])

###Read in patient data only
tmp1 <- read_excel(files[1],range = "Core Qs!A6:AL10",col_names=actualColNames)
tmp2 <- read_excel(files[1],range = "Additional Qs!A6:AC10",col_names=AddColNamesreal)
patient_data <- cbind(tmp1,tmp2)
###Read in clinician data
clinician_data <- read_excel(files[1],range = "Core Qs!C12:O12",col_names=ClinNames)


####Read in data for all others
for(f in 2:length(files)){
  ###Read in patient data only
  n1 <- read_excel(files[f],range = "Core Qs!A6:AL10",col_names=actualColNames)
  n2 <- read_excel(files[f],range = "Additional Qs!A6:AC10",col_names=AddColNamesreal)
  n3 <- cbind(n1,n2)
  patient_data <- rbind(patient_data,n3)
  
  ###Read in clinician data
  n8 <- read_excel(files[f],range = "Core Qs!C12:O12",col_names=ClinNames)
  clinician_data <- rbind(clinician_data,n8)
  print(paste("Done:",files[f]))
}

##Tidy up
clinician_data <- clinician_data[,c(1,4,9,13)]
patient_data <- patient_data[,-c(3,11)]
patient_data$Submitter <- rep(clinician_data$NAME,each=5)
patient_data$Sub_Trust <- rep(clinician_data$`TRUST  NAME`,each=5)
patient_data$HOSPITAL[146:150] <- patient_data$Sub_Trust[146:150]
##remove NA rows
patient_data_clean <- patient_data[which(!is.na(patient_data$HOSPITAL)),]
table(patient_data_clean$Sub_Trust)
##Add regions
patient_data_clean$Region <- rep(NA,times=159)
patient_data_clean$Region[which(patient_data_clean$Sub_Trust %in% c("NWAFT","Addenbrooke's Hospital, Cambridge University Hospitals","Norfolk and Norwich University Hospitals NHS Foundation trust"))] <- "East"
patient_data_clean$Region[which(patient_data_clean$Sub_Trust %in% c("Great Western Hospital NHS Foundation Trust","North Bristol NHS Trust","Royal Berkshire Hospital NHS Foundation Trust","Royal Cornwall Hospitals NHS Trust","Royal Devon University Healthcare NHS FT","Torbay and South Devon","University Hospital Dorset NHS Foundation Trust"))] <- "South West"
patient_data_clean$Region[which(patient_data_clean$Sub_Trust %in% c("Ashford and St Peter's NHS trust","University Hospitals Sussex, RSCH"))] <- "South East"
patient_data_clean$Region[which(patient_data_clean$Sub_Trust %in% c("Buckinghamshire Healthcare NHS Trust","Nottingham University Hospitals NHS Trust","Oxford University Hospitals","Sherwood Forest Hospitals NHS F T","University Hospital Birmingham NHS Foundation Trust","University Hospitals Leicester","University Hospitals Of Derby and Burton NHS Foundation Trust"))] <- "Midlands"
patient_data_clean$Region[which(patient_data_clean$Sub_Trust %in% c("Barts Health NHS Trust","Chelsea and Westminster Hospital", "Guy's & St Thomas' NHS Foundation Trust","Epsom and St Helier University Hospitals NHS Trust","Imperial College Healthcare NHS Trust","King's College Hospital, London","Lewisham Hospital","London North West University","Queens hospital, Romford, London","St Georges Hospital Foundation Trust","The Royal Marsden","University College London Hospitals",""))] <- "London"
patient_data_clean$Region[which(patient_data_clean$Sub_Trust %in% c("NHS Newcastle upon Tyne Hospitals",""))] <- "North East"
patient_data_clean$Region[which(patient_data_clean$Sub_Trust %in% c("Leeds Teaching Hospitals NHS Trust","St Helens and Knowsley NHS Teaching Hospitals NHS Trust","Teaching Hospitals of Wrightington, Wigan and Leigh NHS Foundation Trust"))] <- "North West"
patient_data_clean$Region[which(patient_data_clean$Sub_Trust %in% c("NHS Ayrshire and Arran","NHS Lanarkshire","NHS Lothian","NHS Tayside"))] <- "Scotland"
table(patient_data_clean$Region)
hist(patient_data_clean$Region)
##clean up age
patient_data_clean$`Age at diagnosis (number, years)` <- gsub("years","",patient_data_clean$`Age at diagnosis (number, years)`,T)
patient_data_clean$`Age at diagnosis (number, years)` <- as.numeric(patient_data_clean$`Age at diagnosis (number, years)`)
mean(patient_data_clean[,3],na.rm = T)

##Save out
##load("rawData.RData")
save(clinician_data,patient_data,file = "SJSTENrawData.RData")
write.csv2(x = clinician_data,file = "TEN_Audit_Clinician_Master_List.csv")
write.csv2(x = patient_data,file = "TEN_Audit_Patient_Data_List.csv")
