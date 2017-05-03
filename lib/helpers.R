helper.function <- function()
{
  return(1)
}

######Function to export excel with sheet overwrite feature######
w.excel <- function(object, file) {
    if(file.exists(file)) {
        match <- sum(grepl(as.character(Sys.Date() - 1), 
                           excel_sheets(file)))
        if(match > 0) {
            wb <- loadWorkbook(file)
            removeSheet(wb, sheetName = as.character(Sys.Date() - 1))
            newsheet <- createSheet(wb, sheetName = as.character(Sys.Date() - 1))
            addDataFrame(object, newsheet, row.names = FALSE)
            saveWorkbook(wb, file)
        } else {
            write.xlsx(object, file, 
                       sheetName = as.character(Sys.Date() - 1), 
                       append = TRUE, row.names = FALSE)
        }
    } else {
        write.xlsx(object, file, sheetName = as.character(Sys.Date() - 1), 
                   append = TRUE, row.names = FALSE)
    }
}