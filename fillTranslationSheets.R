# problem loading rJava, http://stackoverflow.com/a/9120712
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk1.7.0_75\\jre')
require(XLConnect)
source('loadTranslation.R')
fillTranslationSheets <- function(excelFile, currentLangFile, currentMainFile, latestLangFile, latestMainFile) {
        print('Reading current translation files')
        currentTranslation <- loadTranslation(currentLangFile, currentMainFile)
        print('Reading latest translation files')
        #latestTranslation <- loadTranslation(latestLangFile, latestMainFile)
        
        # read excel workbook and its sheets
        workbook <- loadWorkbook(excelFile)
        sheetNames <- getSheets(workbook)
        sheets <- readWorksheet(workbook, sheetNames)
        
        #################### Functions, Constants ####################
        
        # skip first 3 rows since they contain no keys / only header infos
        ROW_INDEX_FIRST_KEY <- 4

        # the columns to be populated and checked
        COLUMN_NAME_ORIGINAL_TEXT <- 'OriginalText'
        COLUMN_NAME_TEXT <- 'Text'
        columnList <- list(COLUMN_NAME_ORIGINAL_TEXT, COLUMN_NAME_TEXT)
        
        
        # summary sheet
        SUMMARY_SHEET_NAME <- 'Summary'
        SUMMARY_COLUMN_NAME_SHEET <- 'Sheet'
        SUMMARY_COLUMN_NAME_STATUS <- 'Status'
        SUMMARY_COLUMN_NAME_DESCRIPTION <- 'Description'
        SUMMARY_COLUMN_LENGTH <- 3
        # cell style for row OK
        SUMMARY_CELL_STYLE_OK <- createCellStyle(workbook)
        setFillPattern(SUMMARY_CELL_STYLE_OK, fill = XLC$FILL.SOLID_FOREGROUND)
        setFillForegroundColor(SUMMARY_CELL_STYLE_OK, color = XLC$COLOR.LIGHT_GREEN)
        # cell style for row ERROR
        SUMMARY_CELL_STYLE_ERROR <- createCellStyle(workbook)
        setFillPattern(SUMMARY_CELL_STYLE_ERROR, fill = XLC$FILL.SOLID_FOREGROUND)
        setFillForegroundColor(SUMMARY_CELL_STYLE_ERROR, color = XLC$COLOR.RED)
        # cell style for row details ERROR
        SUMMARY_CELL_STYLE_ERROR_DETAILS <- createCellStyle(workbook)
        setFillPattern(SUMMARY_CELL_STYLE_ERROR_DETAILS, fill = XLC$FILL.SOLID_FOREGROUND)
        setFillForegroundColor(SUMMARY_CELL_STYLE_ERROR_DETAILS, color = XLC$COLOR.LIGHT_ORANGE)
        # indicates position for next summary
        summaryRowIndex <- 1
        # the summary to be filled into the sheet
        summary <- data.frame(
                Sheet = character(),
                Status = character(),
                stringsAsFactors = FALSE
                )
        if (!SUMMARY_SHEET_NAME %in% sheetNames) {
                createSheet(workbook, name = SUMMARY_SHEET_NAME)
        }
        
        # define cell style and color for Excel for highlighting:
        # - changes in a row: color for complete row
        # - changes in original text: color for a cell in that row
        # - changes in text: color for a cell in that row
        
        # define cell style and color for header row
        CELL_STYLE_HEADER <- createCellStyle(workbook)
        setFillPattern(CELL_STYLE_HEADER, fill = XLC$FILL.SOLID_FOREGROUND)
        setFillForegroundColor(CELL_STYLE_HEADER, color = XLC$COLOR.LIGHT_BLUE)
        
        # define cell style and color for row changes
        CELL_STYLE_ROW_CHANGED <- createCellStyle(workbook)
        setWrapText(CELL_STYLE_ROW_CHANGED, wrap = T)
        setFillPattern(CELL_STYLE_ROW_CHANGED, fill = XLC$FILL.SOLID_FOREGROUND)
        setFillForegroundColor(CELL_STYLE_ROW_CHANGED, color = XLC$COLOR.LIGHT_YELLOW)
        
        # define cell style and color for original text and text changes in a single cell
        CELL_STYLE_TEXT_CHANGED <- createCellStyle(workbook)
        setWrapText(CELL_STYLE_TEXT_CHANGED, wrap = T)
        setFillPattern(CELL_STYLE_TEXT_CHANGED, fill = XLC$FILL.SOLID_FOREGROUND)
        setFillForegroundColor(CELL_STYLE_TEXT_CHANGED, color = XLC$COLOR.LIGHT_ORANGE)
        
        # define cell style for wrapping text
        CELL_STYLE_WRAP_TEXT <- createCellStyle(workbook)
        setWrapText(CELL_STYLE_WRAP_TEXT, wrap = T)
        
        updateSheetStyles <- function() {
                # something has changed?
                # highlight the rows with our defined cell styles above
                for (rowNumber in ROW_INDEX_FIRST_KEY:rowNumbers) {
                        if (rowNumber %in% changedRows) {
                                # cell style for complete row has changed
                                setCellStyle(workbook,
                                             sheet = sheetName,
                                             row = rowNumber,
                                             col = 1:colLength,
                                             cellstyle = CELL_STYLE_ROW_CHANGED)
                                
                                # cell style for changed original text
                                if (rowNumber %in% changedOriginalRows) {
                                        setCellStyle(workbook, 
                                                     sheet = sheetName, 
                                                     row = rowNumber, 
                                                     col = colIndexOriginalText, 
                                                     cellstyle = CELL_STYLE_TEXT_CHANGED)
                                }
                                
                                # cell style for changed text
                                if (rowNumber %in% changedTextRows) {
                                        setCellStyle(workbook, 
                                                     sheet = sheetName, 
                                                     row = rowNumber, 
                                                     col = colIndexText, 
                                                     cellstyle = CELL_STYLE_TEXT_CHANGED)
                                }
                        } else {
                                # cell style for wrapping text on complete sheet 
                                setCellStyle(workbook,
                                             sheet = sheetName,
                                             row = rowNumber,
                                             col = 1:colLength,
                                             cellstyle = CELL_STYLE_WRAP_TEXT)
                        }
                }
        }
        
        # function to populate a sheet's cell
        # returns TRUE if it has changed else FALSE
        changedValue <- function(cellValue, translationText) {
                # cell has changed based on these rules:
                # - empty/NA/NULL: no translation yet
                # - value differs from data frame
                if (is.null(cellValue) || is.na(cellValue) || cellValue != translationText) {
                        TRUE
                } else {
                        FALSE
                }
        }
        #################### END ####################
        
        # the lists containing row indices with status ok or error
        summary_row_list_sheet_status_ok <- list()
        summary_row_list_sheet_status_error <- list()
        summary_row_list_sheet_details_error <- list()

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
                colIndexOriginalText <- which(columnNames == COLUMN_NAME_ORIGINAL_TEXT)
                # position of text (cell style when it has changed)
                colIndexText <- which(columnNames == 'Text')
                
                # initialize 3 lists holding changed rows:
                # - row list for any cell changes (union of the other two change lists)
                changedRows <- list()
                # - row list for changes in 'Original' cell
                changedOriginalRows <- list()
                # - row list for changes in 'Text' cell
                changedTextRows <- list()

                # summary of processing
                overallStatus <- 'OK'
                # write sheet name in summary column 'Sheet'
                summary[summaryRowIndex, SUMMARY_COLUMN_NAME_SHEET] <- sheetName
                # row position of sheet name in summary
                summarySheetNameIndex <- summaryRowIndex
                # all other infos goes to next line
                summaryRowIndex <- summaryRowIndex + 1

                # is there any data?
                rowNumbers <- nrow(currentSheet)
                if (rowNumbers >= ROW_INDEX_FIRST_KEY) {
                        print(paste('Processing sheet', sheetName))
                        
                        # iterate through each row and start filling the sheet
                        # using our translation data frames
                        for (rowNumber in ROW_INDEX_FIRST_KEY:rowNumbers) {
                                #print(c('Process row', rowNumber+1))

                                # read key from sheet
                                keyValue <- currentSheet[rowNumber, 'Key']
                                
                                # skip if Key cell is NULL or NA
                                if (is.null(keyValue) || is.na(keyValue)) {
                                        next
                                }
                                
                                # get translation for this key
                                translationRow <- currentTranslation[currentTranslation$Key == keyValue,]
                                # there must be exactly one translation
                                result <- nrow(translationRow)
                                if (result == 1) {
                                        # set ID in sheet
                                        currentSheet[rowNumber, 'ID'] <- translationRow$ID

                                        # boolean marker to indicate some has changed in this sheet's row
                                        changedColumns <- list()
                                        
                                        # fill columns
                                        for (columnName in columnList) {
                                                if (changedValue(currentSheet[rowNumber, columnName], translationRow[columnName])) {
                                                        print(paste('change ', 
                                                                    columnName, 
                                                                    ' \'', 
                                                                    currentSheet[rowNumber, columnName], 
                                                                    '\' to \'', 
                                                                    translationRow[columnName], 
                                                                    '\'', 
                                                                    sep = ''))
                                                        
                                                        # set cell value
                                                        currentSheet[rowNumber, columnName] <- translationRow[columnName]
                                                        
                                                        # store row number in list
                                                        changedColumns <- c(changedColumns, columnName)
                                                }
                                        }
                                        if (length(changedColumns) > 0) {
                                                #print(paste('Changed row', rowNumber + 1))
                                                changedRows <- c(changedRows, rowNumber + 1)
                                                if (COLUMN_NAME_TEXT %in% changedColumns) {
                                                        changedTextRows <- c(changedTextRows, rowNumber + 1)
                                                }
                                                if (COLUMN_NAME_ORIGINAL_TEXT %in% changedColumns) {
                                                        changedOriginalRows <- c(changedOriginalRows, rowNumber + 1)
                                                }
                                        }
                                } else {
                                        overallStatus <- 'Error'
                                        # fill text in column status
                                        summary[summaryRowIndex, SUMMARY_COLUMN_NAME_STATUS] <- 
                                                paste('Unknown key \'', 
                                                      keyValue, '\'', 
                                                      sep = '')
                                        # store row position of error details in list
                                        summary_row_list_sheet_details_error <- c(summary_row_list_sheet_details_error, summaryRowIndex + 1)
                                        # fill text in column description
                                        summary[summaryRowIndex, SUMMARY_COLUMN_NAME_DESCRIPTION] <- 
                                                paste(nrow(translationRow), 
                                                      ' translations found in Sheet \'', 
                                                      sheetName, 
                                                      '\', row ', 
                                                      rowNumber+1, 
                                                      '.', 
                                                      sep = '')
                                        summaryRowIndex <- summaryRowIndex + 1
                                        
                                }
                        }
                        
                        # first write the sheet back into the workbook
                        # before doing further changes e.g. cell styles
                        writeWorksheet(workbook, currentSheet, sheet = sheetName)

                        updateSheetStyles()
                }
                text <- ''
                for (columnName in columnList) {
                        if (columnName == COLUMN_NAME_TEXT) {
                                text <- paste(text, ' ', length(changedTextRows), ' changes in ', columnName, '.', sep = '')
                        } else if (columnName == COLUMN_NAME_ORIGINAL_TEXT) {
                                text <- paste(text, ' ', length(changedOriginalRows), ' changes in ', columnName, '.', sep = '')
                        }
                }
                summary[summaryRowIndex, SUMMARY_COLUMN_NAME_DESCRIPTION] <- text
                summary[summaryRowIndex, SUMMARY_COLUMN_NAME_STATUS] <- paste(length(changedOriginalRows) + length(changedTextRows), ' total changes.')
                
                # set header cell style for each sheet
                setCellStyle(workbook, sheet = sheetName, row = 1, col = 1:colLength, cellstyle = CELL_STYLE_HEADER)
                summaryRowIndex <- summaryRowIndex + 1
                if (overallStatus == 'OK') {
                        summary_row_list_sheet_status_ok <- c(summary_row_list_sheet_status_ok, summarySheetNameIndex + 1)
                } else {
                        summary_row_list_sheet_status_error <- c(summary_row_list_sheet_status_error, summarySheetNameIndex + 1)
                }
                summary[summarySheetNameIndex, SUMMARY_COLUMN_NAME_STATUS] <- overallStatus
        }
        # first write data before updating style sheet
        writeWorksheet(workbook, summary, sheet = SUMMARY_SHEET_NAME)
        setCellStyle(workbook,
                     sheet = SUMMARY_SHEET_NAME,
                     row = 1,
                     col = 1:SUMMARY_COLUMN_LENGTH,
                     cellstyle = CELL_STYLE_HEADER)
        for (i in 2:nrow(summary)) {
                # set style for status ok or error
                if (i %in% summary_row_list_sheet_status_ok) {
                        setCellStyle(workbook,
                                     sheet = SUMMARY_SHEET_NAME,
                                     row = i,
                                     col = 1:SUMMARY_COLUMN_LENGTH,
                                     cellstyle = SUMMARY_CELL_STYLE_OK)
                } else if (i %in% summary_row_list_sheet_status_error) {
                        setCellStyle(workbook,
                                     sheet = SUMMARY_SHEET_NAME,
                                     row = i,
                                     col = 1:SUMMARY_COLUMN_LENGTH,
                                     cellstyle = SUMMARY_CELL_STYLE_ERROR)
                } else if (i %in% summary_row_list_sheet_details_error) {
                        setCellStyle(workbook,
                                     sheet = SUMMARY_SHEET_NAME,
                                     row = i,
                                     col = 2:SUMMARY_COLUMN_LENGTH,
                                     cellstyle = SUMMARY_CELL_STYLE_ERROR_DETAILS)
                }
        }
        setColumnWidth(workbook, sheet = SUMMARY_SHEET_NAME, column = 1:3, width = -1)

        # save result
        saveWorkbook(workbook, paste('new_', excelFile))
}