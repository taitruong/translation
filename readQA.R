require(XLConnect)
importSheets <- function(filename) {
        # import sheets source from: http://stackoverflow.com/a/12945733
        workbook <- loadWorkbook(filename)
        lst <- readWorksheet(workbook, sheet = getSheets(workbook))
}