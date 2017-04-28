library(ProjectTemplate)
load.project()

workbook.name <- "ExportEnPHC_26042017_221404.xlsx"
filename <- "data/ExportEnPHC_26042017_221404.xlsx"
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


DEF2 <- DEF[c('Syn_ID','VisitDate','PtAge','PtWeightValue','PtHeightValue',
              'PtBMIValue','PtWaistValue','PtWaistUnit','SmokingStatus',
              'RiskStratification','Morbidity_Type1Db','Morbidity_Type2Db',
              'Morbidity_HPT','Morbidity_HLD','Lab_HbA1cResult','Lab_SBPMean',
              'Lab_DBPMean','Lab_RP_Crn','Lab_LP_TotalCholAnnual',
              'Lab_LP_TotalChol','Lab_LP_TCholResult','Lab_LP_TCholUnit',
              'MoriskyMedAdh','DataCompleted','DateCompleted','CreatedDate',
              'CreatedBy','UpdatedDate','UpdatedBy')]

all.miss <- apply(DEF2, 1, function(x) all(is.na(x)))
DEF2 <- subset(DEF2, !all.miss)

#completion rate
completed <- capture.output(table(DEF2$DataCompleted), split = TRUE)
completed.prop <- capture.output(round(prop.table(table(DEF2$DataCompleted))*100, 
                                       1), split = TRUE)
cat("Completion Count", completed, "Completion Rate", 
    completed.prop, file = "doc/Completion Rate.txt", sep = "\n")

#Data entry by DEP
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
#export results to excel in doc folder
write.xlsx(creator, file = "doc/DEF-Data Entry Performance.xlsx", 
           sheetName = as.character(Sys.Date() - 1), 
           append = TRUE, row.names = FALSE)

#subset complete cases only
DEF3 <- DEF2[DEF2$DataCompleted == "Yes", ]
summary(DEF3)

#continuous variables outliers
DEF4 <- DEF3[c('Syn_ID', 'CreatedBy', 'PtAge','PtWeightValue','PtHeightValue',
               'Lab_HbA1cResult', 'Lab_SBPMean','Lab_DBPMean')]
DEF4$mahal <- mahalanobis(DEF4[3:ncol(DEF4)], 
                          colMeans(DEF4[3:ncol(DEF4)]), 
                          cov(DEF4[3:ncol(DEF4)]))
DEF4$out <- DEF4$mahal > qchisq(0.999, df = ncol(DEF4[3:ncol(DEF4)])-1)
table(DEF4$CreatedBy, DEF4$out)

