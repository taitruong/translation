# source: http://rmsharp.me/2014/08/25/save-a-list-of-data-frames-into-an-excel-workbook-with-a-simple-wrapper-of-some-xlconnect-functions/
# Usage: create_wkbk('myWorkbook.xlsx', list(dataframe1, ...), c('sheetname1', ...))
#' \code{create_wkbk()} creates an Excel workbook with worksheets.
#' 
#' @param file filename of workbook to be created
#' @param df_list list of data frames to be added as worksheets to 
#'   workbook
#' @param sheetnames character vector of worksheet names
#' @param create Specifies if the file should be created if it does not 
#' already exist (default is FALSE). Note that create = TRUE has 
#' no effect if the specified file exists, i.e. an existing file is 
#' loaded and not being recreated if create = TRUE.
#' @import XLConnect
#' @export
create_wkbk <- function(file, df_list, sheetnames, create = TRUE) {
        if (length(df_list) != length(sheetnames))
                stop("Number of dataframes does not match number of worksheet names")
        
        if (file.exists(file) & create)
                file.remove(file)
        
        wkbk <- loadWorkbook(filename = file, create = create)
        for (i in seq_along(df_list)) {
                sheetname <- sheetnames[i]
                df <- df_list[[i]]
                createSheet(wkbk, sheetname)
                writeWorksheet(wkbk, df, sheetname, startRow = 1, startCol = 1, 
                               header = TRUE)
                setColumnWidth(wkbk, sheetname, column = 1:ncol(df), width = -1)
        }
        saveWorkbook(wkbk)
}