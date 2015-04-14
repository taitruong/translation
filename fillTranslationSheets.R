# problem loading rJava, http://stackoverflow.com/a/9120712
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk1.7.0_75\\jre')
require(XLConnect)
# needed for cpos functions
require(cwhmisc)
source('loadTranslation.R')


fillTranslationSheets <- function(excelFile, currentLangFile, currentMainFile, latestLangFile, latestMainFile) {
        # read excel workbook and its sheets
        print(paste('Reading workbook', excelFile))
        workbook <- loadWorkbook(excelFile)
        sheetNames <- getSheets(workbook)
        sheets <- readWorksheet(workbook, sheetNames)
        # define cell style and color for Excel for highlighting:
        # - changes in a row: color for complete row
        # - changes in original text: color for a cell in that row
        # - changes in text: color for a cell in that row

        ## cell styles for translation sheets
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
        
        ## cell styles for summary sheets
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
        
        print('Reading current translation files')
        
        # current language summary sheet
        CURRENT_LANG_SUMMARY_SHEET_NAME <- 'Summary Current Translation'
        LANG_SUMMARY_COLUMN_NAME_DESCRIPTION <- 'Description'
        LANG_SUMMARY_COLUMN_NAME_OUTPUT <- 'Output'
        LANG_SUMMARY_COLUMN_LENGTH <- 2
        currentLangStartRow <- 1
        # the summary to be filled into the sheet
        currentLangSummary <- data.frame(
                Description = character(),
                Output = character(),
                stringsAsFactors = FALSE
        )
        currentLangSummary[currentLangStartRow, LANG_SUMMARY_COLUMN_NAME_DESCRIPTION] <- 'Current lang file:'
        currentLangSummary[currentLangStartRow, LANG_SUMMARY_COLUMN_NAME_OUTPUT] <- currentLangFile
        currentLangStartRow <- currentLangStartRow + 1
        currentLangSummary[currentLangStartRow, LANG_SUMMARY_COLUMN_NAME_DESCRIPTION] <- 'Current main file:'
        currentLangSummary[currentLangStartRow, LANG_SUMMARY_COLUMN_NAME_OUTPUT] <- currentMainFile
        currentLangStartRow <- currentLangStartRow + 1

        currentResultHandler <- function(result, langDf, mainDf) {
                currentLangSummary[currentLangStartRow, LANG_SUMMARY_COLUMN_NAME_DESCRIPTION] <- 'Number of translations in lang file:'
                currentLangSummary[currentLangStartRow, LANG_SUMMARY_COLUMN_NAME_OUTPUT] <- nrow(langDf)
                currentLangStartRow <- currentLangStartRow + 1
                
                currentLangSummary[currentLangStartRow, LANG_SUMMARY_COLUMN_NAME_DESCRIPTION] <- 'Number of translations in main file:'
                currentLangSummary[currentLangStartRow, LANG_SUMMARY_COLUMN_NAME_OUTPUT] <- nrow(mainDf)
                currentLangStartRow <- currentLangStartRow + 1

                # check whether there are missing IDs
                langDfNotInMainDf <- langDf[!langDf$ID %in% mainDf$ID,]
                mainDfNotInLangDf <- mainDf[!mainDf$ID %in% langDf$ID,]
                
                langRowNumber = nrow(langDfNotInMainDf)
                mainRowNumber = nrow(mainDfNotInLangDf)
                mainRowError <- if (langRowNumber > 0) {
                        currentLangSummary[currentLangStartRow, LANG_SUMMARY_COLUMN_NAME_DESCRIPTION] <- paste(langRowNumber ,'IDs in Lang file but not in Main file:')
                        currentLangSummary[currentLangStartRow, LANG_SUMMARY_COLUMN_NAME_OUTPUT] <- toString(langDfNotInMainDf$ID)
                        currentLangStartRow <- currentLangStartRow + 1
                        currentLangStartRow
                } else {
                        -1
                }
                langRowError <- if (mainRowNumber > 0) {
                        currentLangSummary[currentLangStartRow, LANG_SUMMARY_COLUMN_NAME_DESCRIPTION] <-  paste(mainRowNumber ,'IDs in Main file but not in Lang file:')
                        currentLangSummary[currentLangStartRow, LANG_SUMMARY_COLUMN_NAME_OUTPUT] <- toString(mainDfNotInLangDf$ID)
                        currentLangStartRow <- currentLangStartRow + 1
                        currentLangStartRow
                } else {
                        -1
                }
                # check for escape characters
                originalTextEscapeList <- list()
                for (i in 1:nrow(result)) {
                        resultRow <- result[i,]
                        if (!is.na(cpos(resultRow$OriginalText, 'amp;amp;')) 
                            || !is.na(cpos(resultRow$OriginalText, 'lt;lt;'))
                            || !is.na(cpos(resultRow$OriginalText, 'gt;gt;'))
                            || !is.na(cpos(resultRow$OriginalText, 'apos;apos;'))
                            || !is.na(cpos(resultRow$OriginalText, 'quot;quot;'))) {
                                originalTextEscapeList <- c(originalTextEscapeList, resultRow$ID)
                        }
                }
                originalTextRowError <- if (length(originalTextEscapeList) > 0) {
                        currentLangSummary[currentLangStartRow, LANG_SUMMARY_COLUMN_NAME_DESCRIPTION] <- paste(length(textEscapeList), 'escape errors in attribute OriginalText with IDs:')
                        currentLangSummary[currentLangStartRow, LANG_SUMMARY_COLUMN_NAME_OUTPUT] <- toString(originalTextEscapeList)
                        currentLangStartRow <- currentLangStartRow + 1
                        currentLangStartRow
                } else {
                        -1
                }
                textEscapeList <- list()
                for (i in 1:nrow(result)) {
                        resultRow <- result[i,]
                        if (!is.na(cpos(resultRow$Text, 'amp;amp;')) 
                            || !is.na(cpos(resultRow$Text, 'lt;lt;'))
                            || !is.na(cpos(resultRow$Text, 'gt;gt;'))
                            || !is.na(cpos(resultRow$Text, 'apos;apos;'))
                            || !is.na(cpos(resultRow$Text, 'quot;quot;'))) {
                                textEscapeList <- c(textEscapeList, resultRow$ID)
                        }
                }
                textRowError <- if (length(textEscapeList) > 0) {
                        currentLangSummary[currentLangStartRow, LANG_SUMMARY_COLUMN_NAME_DESCRIPTION] <- paste(length(textEscapeList), 'escape errors in attribute Text with IDs:')
                        currentLangSummary[currentLangStartRow, LANG_SUMMARY_COLUMN_NAME_OUTPUT] <- toString(textEscapeList)
                        currentLangStartRow <- currentLangStartRow + 1
                        currentLangStartRow
                } else {
                        -1
                }
                
                # create sheet only when it does not exists yet
                if (!CURRENT_LANG_SUMMARY_SHEET_NAME %in% sheetNames) {
                        createSheet(workbook, name = CURRENT_LANG_SUMMARY_SHEET_NAME)
                }
                writeWorksheet(workbook, currentLangSummary, sheet = CURRENT_LANG_SUMMARY_SHEET_NAME)
                # output column description width is 100 characters long
                setColumnWidth(workbook, sheet = CURRENT_LANG_SUMMARY_SHEET_NAME, column = which(colnames(currentLangSummary) == LANG_SUMMARY_COLUMN_NAME_OUTPUT), width = 100 * 256)
                # auto-size for description column
                setColumnWidth(workbook, sheet = CURRENT_LANG_SUMMARY_SHEET_NAME, column = which(colnames(currentLangSummary) == LANG_SUMMARY_COLUMN_NAME_DESCRIPTION), width = 30 * 256)
                
                # cell styles for the sheet
                setCellStyle(workbook,
                             sheet = CURRENT_LANG_SUMMARY_SHEET_NAME,
                             row = 1,
                             col = 1:LANG_SUMMARY_COLUMN_LENGTH,
                             cellstyle = CELL_STYLE_HEADER)
                if (mainRowError != -1) {
                        setCellStyle(workbook,
                                     sheet = CURRENT_LANG_SUMMARY_SHEET_NAME,
                                     row = mainRowError,
                                     col = 1:LANG_SUMMARY_COLUMN_LENGTH,
                                     cellstyle = SUMMARY_CELL_STYLE_ERROR)
                }
                if (langRowError != -1) {
                        setCellStyle(workbook,
                                     sheet = CURRENT_LANG_SUMMARY_SHEET_NAME,
                                     row = langRowError,
                                     col = 1:LANG_SUMMARY_COLUMN_LENGTH,
                                     cellstyle = SUMMARY_CELL_STYLE_ERROR)
                }
                if (originalTextRowError != -1) {
                        setCellStyle(workbook,
                                     sheet = CURRENT_LANG_SUMMARY_SHEET_NAME,
                                     row = originalTextRowError,
                                     col = 1:LANG_SUMMARY_COLUMN_LENGTH,
                                     cellstyle = SUMMARY_CELL_STYLE_ERROR)
                }
                if (textRowError != -1) {
                        setCellStyle(workbook,
                                     sheet = CURRENT_LANG_SUMMARY_SHEET_NAME,
                                     row = textRowError,
                                     col = 1:LANG_SUMMARY_COLUMN_LENGTH,
                                     cellstyle = SUMMARY_CELL_STYLE_ERROR)
                }
        }
        currentTranslation <- loadTranslation(currentLangFile, currentMainFile, currentResultHandler)
        
        print('Reading latest translation files')
        # latest language summary sheet
        LATEST_LANG_SUMMARY_SHEET_NAME <- 'Summary Latest Translation'
        latestLangStartRow <- 1
        # the summary to be filled into the sheet
        latestLangSummary <- data.frame(
                Description = character(),
                Output = character(),
                stringsAsFactors = FALSE
        )
        latestLangSummary[latestLangStartRow, LANG_SUMMARY_COLUMN_NAME_DESCRIPTION] <- 'Latest lang file:'
        latestLangSummary[latestLangStartRow, LANG_SUMMARY_COLUMN_NAME_OUTPUT] <- currentLangFile
        latestLangStartRow <- latestLangStartRow + 1
        latestLangSummary[latestLangStartRow, LANG_SUMMARY_COLUMN_NAME_DESCRIPTION] <- 'Latest main file:'
        latestLangSummary[latestLangStartRow, LANG_SUMMARY_COLUMN_NAME_OUTPUT] <- currentMainFile
        latestLangStartRow <- latestLangStartRow + 1
        
        latestResultHandler <- function(result, langDf, mainDf) {
                latestLangSummary[latestLangStartRow, LANG_SUMMARY_COLUMN_NAME_DESCRIPTION] <- 'Number of translations in lang file:'
                latestLangSummary[latestLangStartRow, LANG_SUMMARY_COLUMN_NAME_OUTPUT] <- nrow(langDf)
                latestLangStartRow <- latestLangStartRow + 1
                
                latestLangSummary[latestLangStartRow, LANG_SUMMARY_COLUMN_NAME_DESCRIPTION] <- 'Number of translations in main file:'
                latestLangSummary[latestLangStartRow, LANG_SUMMARY_COLUMN_NAME_OUTPUT] <- nrow(mainDf)
                latestLangStartRow <- latestLangStartRow + 1
                
                # check whether there are missing IDs
                langDfNotInMainDf <- langDf[!langDf$ID %in% mainDf$ID,]
                mainDfNotInLangDf <- mainDf[!mainDf$ID %in% langDf$ID,]
                
                langRowNumber = nrow(langDfNotInMainDf)
                mainRowNumber = nrow(mainDfNotInLangDf)
                mainRowError <- if (langRowNumber > 0) {
                        latestLangSummary[latestLangStartRow, LANG_SUMMARY_COLUMN_NAME_DESCRIPTION] <- paste(langRowNumber , 'IDs in Lang file but not in Main file:')
                        latestLangSummary[latestLangStartRow, LANG_SUMMARY_COLUMN_NAME_OUTPUT] <- toString(langDfNotInMainDf$ID)
                        latestLangStartRow <- latestLangStartRow + 1
                        latestLangStartRow
                } else {
                        -1
                }
                langRowError <- if (mainRowNumber > 0) {
                        latestLangSummary[latestLangStartRow, LANG_SUMMARY_COLUMN_NAME_DESCRIPTION] <-  paste(mainRowNumber ,'IDs in Main file but not in Lang file:')
                        latestLangSummary[latestLangStartRow, LANG_SUMMARY_COLUMN_NAME_OUTPUT] <- toString(mainDfNotInLangDf$ID)
                        latestLangStartRow <- latestLangStartRow + 1
                        latestLangStartRow
                } else {
                        -1
                }
                # check for escape characters
                originalTextEscapeList <- list()
                for (i in 1:nrow(result)) {
                        resultRow <- result[i,]
                        if (!is.na(cpos(resultRow$OriginalText, 'amp;amp;')) 
                            || !is.na(cpos(resultRow$OriginalText, 'lt;lt;'))
                            || !is.na(cpos(resultRow$OriginalText, 'gt;gt;'))
                            || !is.na(cpos(resultRow$OriginalText, 'apos;apos;'))
                            || !is.na(cpos(resultRow$OriginalText, 'quot;quot;'))) {
                                originalTextEscapeList <- c(originalTextEscapeList, resultRow$ID)
                        }
                }
                originalTextRowError <- if (length(originalTextEscapeList) > 0) {
                        currentLangSummary[latestLangStartRow, LANG_SUMMARY_COLUMN_NAME_DESCRIPTION] <- paste(length(textEscapeList), 'escape errors in attribute OriginalText with IDs:')
                        currentLangSummary[latestLangStartRow, LANG_SUMMARY_COLUMN_NAME_OUTPUT] <- toString(originalTextEscapeList)
                        currentLangStartRow <- latestLangStartRow + 1
                        latestLangStartRow
                } else {
                        -1
                }
                textEscapeList <- list()
                for (i in 1:nrow(result)) {
                        resultRow <- result[i,]
                        if (!is.na(cpos(resultRow$Text, 'amp;amp;')) 
                            || !is.na(cpos(resultRow$Text, 'lt;lt;'))
                            || !is.na(cpos(resultRow$Text, 'gt;gt;'))
                            || !is.na(cpos(resultRow$Text, 'apos;apos;'))
                            || !is.na(cpos(resultRow$Text, 'quot;quot;'))) {
                                textEscapeList <- c(textEscapeList, resultRow$ID)
                        }
                }
                textRowError <- if (length(textEscapeList) > 0) {
                        currentLangSummary[latestLangStartRow, LANG_SUMMARY_COLUMN_NAME_DESCRIPTION] <- paste(length(textEscapeList), 'escape errors in attribute Text with IDs:')
                        currentLangSummary[latestLangStartRow, LANG_SUMMARY_COLUMN_NAME_OUTPUT] <- toString(textEscapeList)
                        latestLangStartRow <- latestLangStartRow + 1
                        latestLangStartRow
                } else {
                        -1
                }
                
                # create sheet only when it does not exists yet
                if (!LATEST_LANG_SUMMARY_SHEET_NAME %in% sheetNames) {
                        createSheet(workbook, name = LATEST_LANG_SUMMARY_SHEET_NAME)
                }
                
                writeWorksheet(workbook, latestLangSummary, sheet = LATEST_LANG_SUMMARY_SHEET_NAME)
                # output column description width is 100 characters long
                setColumnWidth(workbook, sheet = LATEST_LANG_SUMMARY_SHEET_NAME, column = which(colnames(latestLangSummary) == LANG_SUMMARY_COLUMN_NAME_OUTPUT), width = 100 * 256)
                # auto-size for description column
                setColumnWidth(workbook, sheet = LATEST_LANG_SUMMARY_SHEET_NAME, column = which(colnames(latestLangSummary) == LANG_SUMMARY_COLUMN_NAME_DESCRIPTION), width = -1)

                # cell styles for the sheet
                setCellStyle(workbook,
                             sheet = LATEST_LANG_SUMMARY_SHEET_NAME,
                             row = 1,
                             col = 1:LANG_SUMMARY_COLUMN_LENGTH,
                             cellstyle = CELL_STYLE_HEADER)
                if (mainRowError != -1) {
                        setCellStyle(workbook,
                                     sheet = LATEST_LANG_SUMMARY_SHEET_NAME,
                                     row = mainRowError,
                                     col = 1:LANG_SUMMARY_COLUMN_LENGTH,
                                     cellstyle = SUMMARY_CELL_STYLE_ERROR)
                }
                if (langRowError != -1) {
                        setCellStyle(workbook,
                                     sheet = LATEST_LANG_SUMMARY_SHEET_NAME,
                                     row = langRowError,
                                     col = 1:LANG_SUMMARY_COLUMN_LENGTH,
                                     cellstyle = SUMMARY_CELL_STYLE_ERROR)
                }
                if (originalTextRowError != -1) {
                        setCellStyle(workbook,
                                     sheet = LATEST_LANG_SUMMARY_SHEET_NAME,
                                     row = originalTextRowError,
                                     col = 1:LANG_SUMMARY_COLUMN_LENGTH,
                                     cellstyle = SUMMARY_CELL_STYLE_ERROR)
                }
                if (textRowError != -1) {
                        setCellStyle(workbook,
                                     sheet = LATEST_LANG_SUMMARY_SHEET_NAME,
                                     row = textRowError,
                                     col = 1:LANG_SUMMARY_COLUMN_LENGTH,
                                     cellstyle = SUMMARY_CELL_STYLE_ERROR)
                }
        }
        latestTranslation <- loadTranslation(latestLangFile, latestMainFile, latestResultHandler)
        
        #################### Functions, Constants ####################
        
        # skip first 3 rows since they contain no keys / only header infos
        ROW_INDEX_FIRST_KEY <- 4

        # the columns to be populated and checked
        COLUMN_NAME_ORIGINAL_TEXT <- 'OriginalText'
        COLUMN_NAME_TEXT <- 'Text'
        COLUMN_NAME_ID <- 'ID'
        COLUMN_NAME_KEY <- 'Key'
        COLUMN_NAME_DESCRIPTION <- 'Description'
        columnList <- list(COLUMN_NAME_ORIGINAL_TEXT, COLUMN_NAME_TEXT)
        
        # summary sheet
        SUMMARY_SHEET_NAME <- 'Summary Sheets'
        SUMMARY_COLUMN_NAME_SHEET <- 'Sheet'
        SUMMARY_COLUMN_NAME_STATUS <- 'Status'
        SUMMARY_COLUMN_NAME_DESCRIPTION <- 'Description'
        SUMMARY_COLUMN_LENGTH <- 3
        # indicates position for next summary
        summaryRowIndex <- 1
        # the summary to be filled into the sheet
        summary <- data.frame(
                Sheet = character(),
                Status = character(),
                stringsAsFactors = FALSE
                )
        # create sheet only when it does not exists yet
        if (!SUMMARY_SHEET_NAME %in% sheetNames) {
                createSheet(workbook, name = SUMMARY_SHEET_NAME)
        }
        
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
                # position of id
                colIndexId <- which(columnNames == 'ID')
                # position of key
                colIndexKey <- which(columnNames == 'Key')
                
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
                
                # set column widths
                # 1st column description width is 10 characters long
                setColumnWidth(workbook, sheet = sheetName, column = which(columnNames == COLUMN_NAME_DESCRIPTION), width = 10 * 256)
                for (i in 2:colLength) {
                        # auto-size for id and key column
                        # width 20 chars long for all other columns
                        columnWidth <- if (i == colIndexId || i == colIndexKey) -1 else 20 * 256
                        setColumnWidth(workbook, sheet = sheetName, column = i, width = columnWidth)
                }
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
        # seems not to work: set active sheet...
        setActiveSheet(workbook, sheet = SUMMARY_SHEET_NAME)
        # save result
        saveWorkbook(workbook, paste('new_', excelFile))
}