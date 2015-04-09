# problem loading rJava, http://stackoverflow.com/a/9120712
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk1.7.0_75\\jre')
require(XLConnect)
source('loadTranslation.R')
fillTranslationSheets <- function(excelFile, currentLangFile, currentMainFile, latestLangFile, latestMainFile) {
        translation <- loadTranslation('D:/000/recruitingappTranslation_eng_new.xml', 'D:/000/recruitingappTranslation_Main.xml')
        
        # read excel workbook and its sheets
        workbook <- loadWorkbook(excelFile)
        sheetNames <- getSheets(workbook)
        sheets <- readWorksheet(workbook, sheetNames)
        
        # define cell style and color for Excel for highlighting:
        # - changes in a row: color for complete row
        # - changes in original text: color for a cell in that row
        # - changes in text: color for a cell in that row
        
        # define cell style and color for header row
        headerCellstyle <- createCellStyle(workbook)
        setFillPattern(headerCellstyle, fill = XLC$FILL.SOLID_FOREGROUND)
        setFillForegroundColor(headerCellstyle, color = XLC$COLOR.LIGHT_BLUE)
        
        # define cell style and color for row changes
        changedRowCellstyle <- createCellStyle(workbook)
        setFillPattern(changedRowCellstyle, fill = XLC$FILL.SOLID_FOREGROUND)
        setFillForegroundColor(changedRowCellstyle, color = XLC$COLOR.LIGHT_YELLOW)
        
        # define cell style and color for original text and text changes in a single cell
        changedTextRowCellstyle <- createCellStyle(workbook)
        setFillPattern(changedTextRowCellstyle, fill = XLC$FILL.SOLID_FOREGROUND)
        setFillForegroundColor(changedTextRowCellstyle, color = XLC$COLOR.LIGHT_ORANGE)
        
        # function to populate a sheet's cell
        # returns TRUE if it has changed else FALSE
        setCell <- function(sheet, # the sheet to be changed
                            rowNumber, # the sheet's row of the cell to be changed
                            colName, # the column name of the cell to be changed
                            translationRow # the translation data to update the cell
                            ) {
                # get value from sheet's cell
                cellValue <- sheet[rowNumber, colName]
                # cell has changed based on these rules:
                # - empty/NA/NULL: no translation yet
                # - value differs from data frame
                if (is.null(cellValue) || is.na(cellValue) || cellValue != translationRow[colName]) {
                        #print(paste('change ', colName, ' \'', cellValue, '\' to \'', translationRow[colName], '\'', sep = ''))
                        
                        # set cell value
                        sheet[rowNumber, colName] <- translationRow[colName]
                        TRUE
                } else {
                        FALSE
                }
        }
        
        # process for each sheet:
        # read key in each row
        # based on the key read the translation data frames and populate cells
        for (sheetName in sheetNames) {
                currentSheet <- readWorksheet(workbook, sheet = sheetName)
                
                # get col positions required to set cell styles below
                columnNames <- colnames(currentSheet)
                # length - last position of last column (cell style for whole row)
                colLength <- length(columnNames)
                # position of original text (cell style when it has changed)
                colIndexOriginalText <- which(columnNames == 'OriginalText')
                # position of text (cell style when it has changed)
                colIndexText <- which(columnNames == 'Text')
                
                sheetRowNumbers <- nrow(currentSheet)
                # is there any data?
                # skip first 3 rows since they contain no keys / only header infos
                if (sheetRowNumbers > 3) {
                        print(paste('Processing sheet', sheetName))
                        # initialize 3 lists holding changed rows:
                        # - row list for any cell changes (union of the other two change lists)
                        changedRows <- list()
                        # - row list for changes in 'Original' cell
                        changedOriginalRows <- list()
                        # - row list for changes in 'Text' cell
                        changedTextRows <- list()
                        # iterate through each row and start filling the sheet
                        # using our translation data frames
                        for (sheetRowNumber in 4:sheetRowNumbers) {
                                keyValue <- currentSheet[sheetRowNumber, 'Key']
                                #print(c('Process row', sheetRowNumber+1))
                                # skip if Key cell is NULL or NA
                                if (is.null(keyValue) || is.na(keyValue)) {
                                        next
                                }
                                # get translation for this key
                                translationRow <- translation[translation$Key == keyValue,]
                                # there must be exactly one translation
                                result <- nrow(translationRow)
                                if (result == 1) {
                                        # set ID in sheet
                                        currentSheet[sheetRowNumber, 'ID'] <- translationRow$ID

                                        # boolean marker to indicate some has changed in this sheet's row
                                        changed <- FALSE
                                        
                                        # fill OriginalText cell
                                        if (setCell(currentSheet, sheetRowNumber, 'OriginalText', translationRow)) {
                                                # store row number in list
                                                changedOriginalRows <- c(changedOriginalRows, sheetRowNumber + 1)
                                                changed <- TRUE
                                        }
                                        
                                        # fill Text cell
                                        if (setCell(currentSheet, sheetRowNumber, 'Text', translationRow)) {
                                                # store row number in list
                                                changedTextRows <- c(changedTextRows, sheetRowNumber + 1)
                                                changed <- TRUE
                                        }

                                        if (changed) {
                                                #print(paste('Changed row', sheetRowNumber + 1))
                                                changedRows <- c(changedRows, sheetRowNumber + 1)
                                        }
                                } else {
                                        stop(c(result,
                                               ' results found. Could not find key ', 
                                               keyValue,
                                               ' in translation file. Error in sheet ',
                                               sheetName,
                                               ', row ',
                                               sheetRowNumber+1,
                                               '. Translation:\n',
                                               translationRow))
                                }
                        }
                        
                        # something has changed?
                        # highlight the rows with our defined cell styles above
                        if (length(changedRows) > 0) {
                                # write the sheet back into the workbook
                                writeWorksheet(workbook, currentSheet, sheet = sheetName)
                                
                                for (changedRow in changedRows) {
                                        # cell style for complete row has changed
                                        setCellStyle(workbook,
                                                     sheet = sheetName,
                                                     row = changedRow,
                                                     col = 1:colLength,
                                                     cellstyle = changedRowCellstyle)
                                        
                                        # cell style for changed original text
                                        if (changedRow %in% changedOriginalRows) {
                                                setCellStyle(workbook, 
                                                             sheet = sheetName, 
                                                             row = changedRow, 
                                                             col = colIndexOriginalText, 
                                                             cellstyle = changedTextRowCellstyle)
                                        }
                                        
                                        # cell style for changed text
                                        if (changedRow %in% changedTextRows) {
                                                setCellStyle(workbook, 
                                                             sheet = sheetName, 
                                                             row = changedRow, 
                                                             col = colIndexText, 
                                                             cellstyle = changedTextRowCellstyle)
                                        }
                                }
                                
                                print(paste(length(changedRows), "row(s) updated."))
                        }
                } else {
                        print(paste('Skip empty sheet', sheetName))
                }
                # set header cell style for each sheet
                setCellStyle(workbook, sheet = sheetName, row = 1, col = 1:colLength, cellstyle = headerCellstyle)
        }
        # save result
        saveWorkbook(workbook, paste('new_', excelFile))
}