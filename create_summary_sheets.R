# problem loading rJava, http://stackoverflow.com/a/9120712
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk1.7.0_75\\jre')
require(XLConnect)
LoadTranslationAndCreateSummarySheet <- function(language.file, main.file, workbook, summarySheetName) {
	source('loadTranslation.R')
	CreateLanguageSummarySheetHandler <- function(summarySheetName) {
		kColumnNameDescription <- 'Description'
		kColumnNameOutput <- 'Output'
		kColumnLength <- 2
		startRow <- 1
		# the summary to be filled into the sheet
		summaryDf <- data.frame(
			Description = character(),
			Output = character(),
			stringsAsFactors = FALSE
		)
		summaryDf[startRow, kColumnNameDescription] <- 'Language file:'
		summaryDf[startRow, kColumnNameOutput] <- language.file
		startRow <- startRow + 1
		summaryDf[startRow, kColumnNameDescription] <- 'Main file:'
		summaryDf[startRow, kColumnNameOutput] <- main.file
		startRow <- startRow + 1
		
		currentResultHandler <- function(result, langDf, mainDf) {
			summaryDf[startRow, kColumnNameDescription] <- 'Number of translations in lang file:'
			summaryDf[startRow, kColumnNameOutput] <- nrow(langDf)
			startRow <- startRow + 1
			
			summaryDf[startRow, kColumnNameDescription] <- 'Number of translations in main file:'
			summaryDf[startRow, kColumnNameOutput] <- nrow(mainDf)
			startRow <- startRow + 1
			
			# check whether there are missing IDs
			langDfNotInMainDf <- langDf[!langDf$ID %in% mainDf$ID,]
			mainDfNotInLangDf <- mainDf[!mainDf$ID %in% langDf$ID,]
			
			langRowNumber = nrow(langDfNotInMainDf)
			mainRowNumber = nrow(mainDfNotInLangDf)
			mainRowError <- if (langRowNumber > 0) {
				summaryDf[startRow, kColumnNameDescription] <- paste(langRowNumber ,'IDs in Language file but not in Main file:')
				summaryDf[startRow, kColumnNameOutput] <- toString(langDfNotInMainDf$ID)
				startRow <- startRow + 1
				startRow
			} else {
				-1
			}
			langRowError <- if (mainRowNumber > 0) {
				summaryDf[startRow, kColumnNameDescription] <-  paste(mainRowNumber ,'IDs in Main file but not in Language file:')
				summaryDf[startRow, kColumnNameOutput] <- toString(mainDfNotInLangDf$ID)
				startRow <- startRow + 1
				startRow
			} else {
				-1
			}
			# check for escape errors e.g. for '<' it should be '&lt;' and not '&lt;lt;'
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
				summaryDf[startRow, kColumnNameDescription] <- paste(length(textEscapeList), 'escape errors in attribute OriginalText with IDs:')
				summaryDf[startRow, kColumnNameOutput] <- toString(originalTextEscapeList)
				startRow <- startRow + 1
				startRow
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
				summaryDf[startRow, kColumnNameDescription] <- paste(length(textEscapeList), 'escape errors in attribute Text with IDs:')
				summaryDf[startRow, kColumnNameOutput] <- toString(textEscapeList)
				startRow <- startRow + 1
				startRow
			} else {
				-1
			}
			
			# create sheet only when it does not exists yet
			if (!summarySheetName %in% getSheets(workbook)) {
				createSheet(workbook, name = summarySheetName)
			}
			writeWorksheet(workbook, summaryDf, sheet = summarySheetName)
			# output column description width is 100 characters long
			setColumnWidth(workbook, sheet = summarySheetName, column = which(colnames(summaryDf) == kColumnNameOutput), width = 100 * 256)
			# auto-size for description column
			setColumnWidth(workbook, sheet = summarySheetName, column = which(colnames(summaryDf) == kColumnNameDescription), width = 30 * 256)
			
			# cell styles for the sheet
			setCellStyle(workbook,
									 sheet = summarySheetName,
									 row = 1,
									 col = 1:kColumnLength,
									 cellstyle = CELL_STYLE_HEADER)
			if (mainRowError != -1) {
				setCellStyle(workbook,
										 sheet = summarySheetName,
										 row = mainRowError,
										 col = 1:kColumnLength,
										 cellstyle = SUMMARY_CELL_STYLE_ERROR)
			}
			if (langRowError != -1) {
				setCellStyle(workbook,
										 sheet = summarySheetName,
										 row = langRowError,
										 col = 1:kColumnLength,
										 cellstyle = SUMMARY_CELL_STYLE_ERROR)
			}
			if (originalTextRowError != -1) {
				setCellStyle(workbook,
										 sheet = summarySheetName,
										 row = originalTextRowError,
										 col = 1:kColumnLength,
										 cellstyle = SUMMARY_CELL_STYLE_ERROR)
			}
			if (textRowError != -1) {
				setCellStyle(workbook,
										 sheet = summarySheetName,
										 row = textRowError,
										 col = 1:kColumnLength,
										 cellstyle = SUMMARY_CELL_STYLE_ERROR)
			}
		}
	}
	# load and return data frame
	loadTranslation(language.file, main.file, CreateLanguageSummarySheetHandler(summarySheetName))
}
