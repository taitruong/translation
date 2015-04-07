# problem loading rJava, http://stackoverflow.com/a/9120712
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk1.7.0_75\\jre')
require(XLConnect)
source('loadTranslation.R')
fillTranslationSheets <- function(excelFile) {
        translation <- loadTranslation('D:/000/recruitingappTranslation_eng_new.xml', 'D:/000/recruitingappTranslation_Main.xml')
        workbook <- loadWorkbook(excelFile)
        sheetNames <- getSheets(workbook)
        sheets <- readWorksheet(workbook, sheetNames)
        
        # define cell color for header
        headerCellstyle <- createCellStyle(workbook)
        setFillPattern(headerCellstyle, fill = XLC$FILL.SOLID_FOREGROUND)
        setFillForegroundColor(headerCellstyle, color = XLC$COLOR.LIGHT_BLUE)
        
        # define cell color for changes
        changedRowCellstyle <- createCellStyle(workbook)
        setFillPattern(changedRowCellstyle, fill = XLC$FILL.SOLID_FOREGROUND)
        setFillForegroundColor(changedRowCellstyle, color = XLC$COLOR.LIGHT_YELLOW)
        
        # define cell color for original text and text changes
        changedTextRowCellstyle <- createCellStyle(workbook)
        setFillPattern(changedTextRowCellstyle, fill = XLC$FILL.SOLID_FOREGROUND)
        setFillForegroundColor(changedTextRowCellstyle, color = XLC$COLOR.LIGHT_ORANGE)
        
        for (sheetName in sheetNames) {
                currentSheet <- readWorksheet(workbook, sheet = sheetName)
                
                # define col indices for cell styles below
                columnNames <- colnames(currentSheet)
                colLength <- length(columnNames)
                colIndexOriginalText <- which(columnNames == 'OriginalText')
                colIndexText <- which(columnNames == 'Text')
                
                sheetRowNumbers <- nrow(currentSheet)
                if (sheetRowNumbers > 3) {
                        print(paste('Processing sheet', sheetName))
                        # skip row 1 to 3 since they do not contain keys in column 1
                        changedRows <- seq()
                        changedOriginalRows <- seq()
                        changedTextRows <- seq()
                        for (sheetRowNumber in 4:sheetRowNumbers) {
                                keyValue <- currentSheet[sheetRowNumber, 'Key']
                                #print(c('Process row', sheetRowNumber+1))
                                if (!is.null(keyValue) && !is.na(keyValue)) {
                                        #row data frame
                                        translationRow <- translation[translation$Key == keyValue,]
                                        result <- nrow(translationRow)
                                        if (result == 1) {
                                                currentSheet[sheetRowNumber, 'ID'] <- translationRow$ID
                                                changed <- FALSE

                                                originalText <- currentSheet[sheetRowNumber, 'OriginalText']
                                                if (is.na(originalText) || originalText != translationRow$OriginalText) {
                                                        changed <- TRUE
                                                        changedOriginalRows <- c(changedOriginalRows, sheetRowNumber + 1)
                                                        print(paste('change original text ', originalText, 'to', translationRow$OriginalText))
                                                        currentSheet[sheetRowNumber, 'OriginalText'] <- translationRow$OriginalText
                                                }
                                                
                                                text <- currentSheet[sheetRowNumber, 'Text']
                                                if (is.na(text) || text != translationRow$Text) {
                                                        changed <- TRUE
                                                        changedTextRows <- c(changedTextRows, sheetRowNumber + 1)
                                                        print(paste('change text', text, 'to', translationRow$Text))
                                                        currentSheet[sheetRowNumber, 'Text'] <- translationRow$Text
                                                }
                                                if (changed) {
                                                        print(paste('Change row', sheetRowNumber + 1))
                                                        changedRows <- c(changedRows, sheetRowNumber + 1)
                                                }
                                        } else {
                                                stop(c(result, ' results found. Could not find key ', keyValue,' in translation file. Error in sheet ', sheetName, ', row ', sheetRowNumber+1))
                                        }
                                }
                        }
                        if (length(changedRows) > 0) {
                                writeWorksheet(workbook, currentSheet, sheet = sheetName)
                                
                                #does not work yet - changing cell styles over two columns 
                                for (changedRow in changedRows) {
                                        # color code complete row as changed
                                        setCellStyle(workbook, sheet = sheetName, row = changedRow, col = 1:colLength, cellstyle = changedRowCellstyle)
                                        # color code cell for changed original text
                                        if (changedRow %in% changedOriginalRows) {
                                                setCellStyle(workbook, sheet = sheetName, row = changedRow, col = colIndexOriginalText, cellstyle = changedTextRowCellstyle)
                                        }
                                        # color code cell for changed text
                                        if (changedRow %in% changedTextRows) {
                                                setCellStyle(workbook, sheet = sheetName, row = changedRow, col = colIndexText, cellstyle = changedTextRowCellstyle)
                                        }
                                }
                                
                                print(paste(length(changedRows), "row(s) updated."))
                        }
                } else {
                        print(paste('Skip empty sheet', sheetName))
                }
                setCellStyle(workbook, sheet = sheetName, row = 1, col = 1:colLength, cellstyle = headerCellstyle)
        }
        saveWorkbook(workbook, paste('new_', excelFile))
}