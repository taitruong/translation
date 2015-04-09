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
        SHEET_NAME_SUMMARY = 'Summary'
        COLUMN_NAME_SHEET = 'Sheet'
        COLUMN_NAME_STATUS = 'Status'
        summaryRowIndex = 1
        summary <- data.frame(
                Sheet = character(),
                Status = character(),
                stringsAsFactors = FALSE
                )
        if (!SHEET_NAME_SUMMARY %in% sheetNames) {
                createSheet(workbook, name = SHEET_NAME_SUMMARY)
        }
        
        # define cell style and color for Excel for highlighting:
        # - changes in a row: color for complete row
        # - changes in original text: color for a cell in that row
        # - changes in text: color for a cell in that row
        
        # define cell style and color for header row
        cellStyleHeader <- createCellStyle(workbook)
        setFillPattern(cellStyleHeader, fill = XLC$FILL.SOLID_FOREGROUND)
        setFillForegroundColor(cellStyleHeader, color = XLC$COLOR.LIGHT_BLUE)
        
        # define cell style and color for row changes
        cellStyleRowChanged <- createCellStyle(workbook)
        setWrapText(cellStyleRowChanged, wrap = T)
        setFillPattern(cellStyleRowChanged, fill = XLC$FILL.SOLID_FOREGROUND)
        setFillForegroundColor(cellStyleRowChanged, color = XLC$COLOR.LIGHT_YELLOW)
        
        # define cell style and color for original text and text changes in a single cell
        cellStyleTextChanged <- createCellStyle(workbook)
        setWrapText(cellStyleTextChanged, wrap = T)
        setFillPattern(cellStyleTextChanged, fill = XLC$FILL.SOLID_FOREGROUND)
        setFillForegroundColor(cellStyleTextChanged, color = XLC$COLOR.LIGHT_ORANGE)
        
        # define cell style for wrapping text
        cellStyleWrapText <- createCellStyle(workbook)
        setWrapText(cellStyleWrapText, wrap = T)
        
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
                                             cellstyle = cellStyleRowChanged)
                                
                                # cell style for changed original text
                                if (rowNumber %in% changedOriginalRows) {
                                        setCellStyle(workbook, 
                                                     sheet = sheetName, 
                                                     row = rowNumber, 
                                                     col = colIndexOriginalText, 
                                                     cellstyle = cellStyleTextChanged)
                                }
                                
                                # cell style for changed text
                                if (rowNumber %in% changedTextRows) {
                                        setCellStyle(workbook, 
                                                     sheet = sheetName, 
                                                     row = rowNumber, 
                                                     col = colIndexText, 
                                                     cellstyle = cellStyleTextChanged)
                                }
                        } else {
                                # cell style for wrapping text on complete sheet 
                                setCellStyle(workbook,
                                             sheet = sheetName,
                                             row = rowNumber,
                                             col = 1:colLength,
                                             cellstyle = cellStyleWrapText)
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
                
                rowNumbers <- nrow(currentSheet)
                summary[summaryRowIndex, COLUMN_NAME_SHEET] <- sheetName
                # is there any data?
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
                                        stop(c(result,
                                               ' results found. Could not find key ', 
                                               keyValue,
                                               ' in translation file. Error in sheet ',
                                               sheetName,
                                               ', row ',
                                               rowNumber+1,
                                               '. Translation:\n',
                                               translationRow))
                                }
                        }
                        
                        # first write the sheet back into the workbook
                        # before doing further changes e.g. cell styles
                        writeWorksheet(workbook, currentSheet, sheet = sheetName)

                        updateSheetStyles()
                        
                        print(paste(length(changedRows), "row(s) updated."))
                } else {
                        print(paste('Skip empty sheet', sheetName))
                }
                for (columnName in columnList) {
                        if (columnName == COLUMN_NAME_TEXT) {
                                summary[summaryRowIndex, columnName] <- paste(length(changedTextRows), ' changes.')
                        } else if (columnName == COLUMN_NAME_ORIGINAL_TEXT) {
                                summary[summaryRowIndex, columnName] <- paste(length(changedOriginalRows), ' changes.')
                        }
                }
                
                # set header cell style for each sheet
                setCellStyle(workbook, sheet = sheetName, row = 1, col = 1:colLength, cellstyle = cellStyleHeader)
                summaryRowIndex <- summaryRowIndex + 1
        }
        writeWorksheet(workbook, summary, sheet = SHEET_NAME_SUMMARY)
        # save result
        saveWorkbook(workbook, paste('new_', excelFile))
}