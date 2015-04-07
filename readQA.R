# problem loading rJava, http://stackoverflow.com/a/9120712
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk1.7.0_75\\jre')
require(XLConnect)
source('loadTranslation.R')
importSheets <- function(filename) {
        translation <- loadTranslation('D:/000/recruitingappTranslation_eng_new.xml', 'D:/000/recruitingappTranslation_Main.xml')
        workbook <- loadWorkbook(filename)
        sheetNames <- getSheets(workbook)
        sheets <- readWorksheet(workbook, sheetNames)
        
        greenCell <- createCellStyle(workbook)
        setFillPattern(greenCell, fill = XLC$FILL.SOLID_FOREGROUND)
        setFillForegroundColor(greenCell, color = XLC$COLOR.GREEN)
        for (sheetName in sheetNames) {
                currentSheet <- readWorksheet(workbook, sheet = sheetName)
                sheetRowNumbers <- nrow(currentSheet)
                if (sheetRowNumbers > 3) {
                        print(c('Processing sheet ', sheetName))
                        # skip row 1 to 3 since they do not contain keys in column 1
                        changedRows <- seq()
                        for (sheetRowNumber in 4:sheetRowNumbers) {
                                keyValue <- currentSheet[sheetRowNumber, 'Key']
                                #print(c('Process row', sheetRowNumber+1))
                                if (!is.null(keyValue) && !is.na(keyValue)) {
                                        #row data frame
                                        translationRow <- translation[translation$Key == keyValue,]
                                        result <- nrow(translationRow)
                                        if (result == 1) {
                                                #print(c('change row', sheetRowNumber))
                                                changedRows <- c(changedRows, sheetRowNumber)
                                                currentSheet[sheetRowNumber, 'ID'] <- translationRow$ID
                                                currentSheet[sheetRowNumber, 'OriginalText'] <- translationRow$OriginalText
                                                currentSheet[sheetRowNumber, 'Text'] <- translationRow$Text
                                        } else {
                                                stop(c(result, ' results found. Could not find key ', keyValue,' in translation file. Error in sheet ', sheetName, ', row ', sheetRowNumber+1))
                                        }
                                }
                        }
                        if (length(changedRows) > 0) {
                                writeWorksheet(workbook, currentSheet, sheet = sheetName)
                                #does not work yet - changing cell styles over two columns 
                                #setCellStyle(workbook, sheet = sheetName, row = changedRows, col = c(1,2), cellstyle = greenCell)
                                print(c(length(changedRows), " rows updated."))
                        }
                } else {
                        print(c('Skip empty sheet', sheetName))
                }
        }
        saveWorkbook(workbook, 'test.xlsx')
}