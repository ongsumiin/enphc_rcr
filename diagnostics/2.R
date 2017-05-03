library(ProjectTemplate)
load.project()

workbook.name <- list.files("data/", pattern = "[xlsx]")
filename <- list.files("data/", pattern = "[xlsx]", full.names = TRUE)
sheets <- excel_sheets(filename)

for (sheet.name in sheets) {

    tryCatch(assign(sheet.name,
                    readxl::read_excel(filename,
                                       sheet = sheet.name),
                    envir = ProjectTemplate:::.TargetEnv),
             error = function(e)
             {
                 warning(paste("The worksheet", sheet.name, "didn't load correctly."))
             })
}


##Selected variables of interest##
variables <- c('Syn_ID','VisitDate', 'Site_Name', 'PtAge','PtWeightValue','PtHeightValue',
               'PtBMIValue','PtWaistValue','PtWaistUnit','SmokingStatus',
               'RiskStratification','Morbidity_Type1Db','Morbidity_Type2Db',
               'Morbidity_HPT','Morbidity_HLD','Lab_HbA1cResult','Lab_SBPMean',
               'Lab_DBPMean','Lab_RP_Crn','Lab_LP_TotalCholAnnual',
               'Lab_LP_TotalChol','Lab_LP_TCholResult','Lab_LP_TCholUnit',
               'MoriskyMedAdh','DataCompleted','DateCompleted','CreatedDate',
               'CreatedBy','UpdatedDate','UpdatedBy')

DEF2 <- DEF[variables]

##remove entry with all missing
all.miss <- apply(DEF2, 1, function(x) all(is.na(x)))
DEF2 <- subset(DEF2, !all.miss)

######completion rate######
completed <- capture.output(table(DEF2$DataCompleted), split = TRUE)
completed.prop <- capture.output(round(prop.table(table(DEF2$DataCompleted))*100, 
                                       1), split = TRUE)
cat("Completion Count", completed, "\n", "Completion Rate", 
    completed.prop, file = "diagnostics/Completion Rate.txt", sep = "\n")

######Data entry by DEP######
#aggregate complete cases by data entry person
creator.y <- aggregate(DataCompleted~CreatedBy, 
                     data = DEF2, 
                     function(x) table(x)["Yes"])
#replace NA with 0
creator.y$DataCompleted[is.na(creator.y$DataCompleted)] <- 0
#calculate total complete cases
creator.y[nrow(creator.y)+1, 1:2] <- c("Total", sum(creator.y$DataCompleted))

#aggregate incomplete cases by data entry person
creator.n <- aggregate(DataCompleted~CreatedBy, 
                       data = DEF2, 
                       function(x) table(x)["No"])
#replace NA with 0
creator.n$DataCompleted[is.na(creator.n$DataCompleted)] <- 0
#calculate total incomplete cases
creator.n[nrow(creator.n)+1, 1:2] <- c("Total", sum(creator.n$DataCompleted))
#rename column
names(creator.n)[2] <- "DataIncomplete"

#merge complete and incomplete data frame by data entry person name
creator <- merge(creator.y, creator.n, by = "CreatedBy", sort = FALSE)
#change class to numeric
creator[2:3] <- apply(creator[2:3], 2, function(x) as.numeric(x))
#calculate total cases by each person
creator$Total <- creator$DataCompleted + creator$DataIncomplete
#sort by total ascending order
creator <- creator[order(creator$Total), ]
#export results to excel in diagnostics folder
w.excel(creator, file = "diagnostics/DEF-Data Entry Performance.xlsx")


#######subset complete cases only######
DEF3 <- DEF[DEF$DataCompleted == "Yes", ]
DEF3$month <- factor(format.Date(DEF3$VisitDate, "%b"), 
                     levels = c("Apr", "Mar", "Feb", "Jan", "Dec", "Nov"))

######number of completed DM cases by clinics######
capture.output(addmargins(table(DEF3$Site_Name, DEF3$Morbidity_Type2Db, 
                                dnn = c(" ", "Diabetes"))), 
               file = "diagnostics/DEF by clinics-DM.txt")

capture.output(cat("\n", "By Month", "\n"), 
               addmargins(table(DEF3$Site_Name[DEF3$Morbidity_Type2Db == "Yes"], 
                                DEF3$month[DEF3$Morbidity_Type2Db == "Yes"], 
                                dnn = c(" ", "Diabetes"))), 
               file = "diagnostics/DEF by clinics-DM.txt",
               append = TRUE)

######number of completed HPT cases by clinics######
capture.output(addmargins(table(DEF3$Site_Name, DEF3$Morbidity_HPT, 
                                dnn = c(" ", "Hypertension"))), 
               file = "diagnostics/DEF by clinics-HPT.txt")

capture.output(cat("\n", "By Month", "\n"), 
               addmargins(table(DEF3$Site_Name[DEF3$Morbidity_HPT == "Yes"], 
                                DEF3$month[DEF3$Morbidity_HPT == "Yes"], 
                                dnn = c(" ", "Hypertension"))), 
               file = "diagnostics/DEF by clinics-HPT.txt",
               append = TRUE)

#######Duplicate ID Checks######
#combine IC and other ID
DEF3$PtID <- ifelse(str_count(DEF3$PtIC_New) == 14, DEF3$PtIC_New, 
                    DEF3$PtOtherIDNum)
#tabulate ID number
ICdup <- data.frame(table(DEF3$PtID))
ICdup <- ICdup[order(-ICdup$Freq), ]
#export results
w.excel(ICdup, file = "diagnostics/DEF-IC duplication.xlsx")

a <- tapply(DEF3$PtID, DEF3$Site_Name, table)
capture.output(a, file = "diagnostics/DEF-IC duplication by clinic.txt")

######Completion duration######
#time to complete in min
DEF3$duration <- DEF3$DateCompleted - DEF3$CreatedDate
#duration that is 3 std score below mean
DEF3$duration.low <- scale(DEF3$duration) < -3
#export short duration cases if there are cases
DEF3a <- subset(DEF3, duration.low == TRUE, select = variables)
if(nrow(DEF3a) > 0) {
    w.excel(data.frame(DEF3a), 
            file = "diagnostics/DEF-Duration outliers.xlsx")
}else{
    NULL
}

#export cases less than 5 min if there are cases
DEF3b <- subset(DEF3, duration < 5)
if(nrow(DEF3b) > 0) {
    w.excel(data.frame(DEF3b), 
            file = "diagnostics/DEF-Duration Under 5 min.xlsx")
}else{
    NULL
}

#tabulate cases with short duration by data entry personnel 
capture.output(cat("\n", "Short Duration Outliers", "\n"),
               addmargins(table(DEF3$CreatedBy, DEF3$duration.low,
                                dnn = c("Data Entry Personnel", 
                                        "Short duration"))),
               file = "diagnostics/DEF-duration.txt")
#tabulate cases with duration <5 minits by data entry personnel
capture.output(cat("\n", "Less than 5 minutes", "\n"),
                   addmargins(table(DEF3$CreatedBy, DEF3$duration < 5,
                                dnn = c("Data Entry Personnel", 
                                        "< 5 minutes"))),
               file = "diagnostics/DEF-duration.txt", append = TRUE)


######continuous variables outliers######
#select continuous variables to evaluate
DEF4 <- DEF3[c('Syn_ID', 'CreatedBy', 'Site_Name', 'Morbidity_Type2Db', 
               'Morbidity_HPT', 'PtAge','PtWeightValue',
               'PtHeightValue', 'Lab_HbA1cResult', 'Lab_SBPMean','Lab_DBPMean')]
#calculate mahalanobis score
DEF4$mahal <- mahalanobis(DEF4[(ncol(DEF4)-5):ncol(DEF4)], 
                          colMeans(DEF4[(ncol(DEF4)-5):ncol(DEF4)]), 
                          cov(DEF4[(ncol(DEF4)-5):ncol(DEF4)]))
#compare calculated score with threshold set at 0.1% of chisq distribution
DEF4$out <- DEF4$mahal > qchisq(0.999, df = ncol(DEF4[(ncol(DEF4)-6):ncol(DEF4)])-1)

#tabulation of outlier cases by person
capture.output(addmargins(table(DEF4$CreatedBy, DEF4$out, 
                                dnn = c("Data Entry Personnel", "Outlier Case"))), 
               file = "diagnostics/DEF-outliers by person.txt")

#check age limit for cases
DEF4$agelimit <- DEF4$PtAge < 30

#export results to diagnostics folder
w.excel(data.frame(DEF4), file = "diagnostics/DEF-Outliers.xlsx")








