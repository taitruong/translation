# problem loading rJava, http://stackoverflow.com/a/9120712
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk1.7.0_75\\jre')
require(XLConnect)
# needed for cpos functions
require(cwhmisc)

LoadTranslationAndCreateSummarySheet <- function(main.file, 
																								 english.file,
																								 language.file, 
																								 workbook, 
																								 summary.sheet.name) {
	source('load_translation.R')

	#################### initialisations, variables, constants, and function definitions ####################
	
	# define cell style and color for Excel for highlighting
	## cell styles for translation sheets
	### define cell style and color for header row
	kCellStyleHeader <- createCellStyle(workbook)
	setFillPattern(kCellStyleHeader, fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleHeader, color = XLC$COLOR.LIGHT_BLUE)

	### cell style for row ERROR
	kCellStyleSummaryError <- createCellStyle(workbook)
	setFillPattern(kCellStyleSummaryError, fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleSummaryError, color = XLC$COLOR.RED)

	CreateLanguageSummarySheetHandler <- function(summary.sheet.name) {
		kColumnNameDescription <- 'Description'
		kColumnNameOutput <- 'Output'
		kColumnLength <- 2
		start.row <- 1
		# the summary to be filled into the sheet
		summary.df <- data.frame(
			Description = character(),
			Output = character(),
			stringsAsFactors = FALSE
		)
		summary.df[start.row, kColumnNameDescription] <- 'Language file:'
		summary.df[start.row, kColumnNameOutput] <- language.file
		start.row <- start.row + 1
		summary.df[start.row, kColumnNameDescription] <- 'Main file:'
		summary.df[start.row, kColumnNameOutput] <- main.file
		start.row <- start.row + 1
		
		translation.handler <- function(result, langDf, mainDf) {
			summary.df[start.row, kColumnNameDescription] <- 'Number of translations in lang file:'
			summary.df[start.row, kColumnNameOutput] <- nrow(langDf)
			start.row <- start.row + 1
			
			summary.df[start.row, kColumnNameDescription] <- 'Number of translations in main file:'
			summary.df[start.row, kColumnNameOutput] <- nrow(mainDf)
			start.row <- start.row + 1
			
			# check whether there are missing IDs
			lang.df.not.in.main.df <- langDf[!langDf$ID %in% mainDf$ID,]
			main.df.not.in.lang.df <- mainDf[!mainDf$ID %in% langDf$ID,]
			
			lang.row.number = nrow(lang.df.not.in.main.df)
			main.row.number = nrow(main.df.not.in.lang.df)
			main.row.error <- if (lang.row.number > 0) {
				summary.df[start.row, kColumnNameDescription] <- paste(lang.row.number ,'IDs in Language file but not in Main file:')
				summary.df[start.row, kColumnNameOutput] <- toString(lang.df.not.in.main.df$ID)
				start.row <- start.row + 1
				start.row
			} else {
				-1
			}
			lang.row.error <- if (main.row.number > 0) {
				summary.df[start.row, kColumnNameDescription] <-  paste(main.row.number ,'IDs in Main file but not in Language file:')
				summary.df[start.row, kColumnNameOutput] <- toString(main.df.not.in.lang.df$ID)
				start.row <- start.row + 1
				start.row
			} else {
				-1
			}
			# check for escape errors e.g. for '<' it should be '&lt;' and not '&lt;lt;'
			original.text.escape.list <- list()
			for (i in 1:nrow(result)) {
				resultRow <- result[i,]
				if (!is.na(cpos(resultRow$OriginalText, 'amp;amp;')) 
						|| !is.na(cpos(resultRow$OriginalText, 'lt;lt;'))
						|| !is.na(cpos(resultRow$OriginalText, 'gt;gt;'))
						|| !is.na(cpos(resultRow$OriginalText, 'apos;apos;'))
						|| !is.na(cpos(resultRow$OriginalText, 'quot;quot;'))) {
					original.text.escape.list <- c(original.text.escape.list, resultRow$ID)
				}
			}
			original.text.row.error <- if (length(original.text.escape.list) > 0) {
				summary.df[start.row, kColumnNameDescription] <- paste(length(text.escape.list), 'escape errors in attribute OriginalText with IDs:')
				summary.df[start.row, kColumnNameOutput] <- toString(original.text.escape.list)
				start.row <- start.row + 1
				start.row
			} else {
				-1
			}
			text.escape.list <- list()
			for (i in 1:nrow(result)) {
				resultRow <- result[i,]
				if (!is.na(cpos(resultRow$Text, 'amp;amp;')) 
						|| !is.na(cpos(resultRow$Text, 'lt;lt;'))
						|| !is.na(cpos(resultRow$Text, 'gt;gt;'))
						|| !is.na(cpos(resultRow$Text, 'apos;apos;'))
						|| !is.na(cpos(resultRow$Text, 'quot;quot;'))) {
					text.escape.list <- c(text.escape.list, resultRow$ID)
				}
			}
			text.row.error <- if (length(text.escape.list) > 0) {
				summary.df[start.row, kColumnNameDescription] <- paste(length(text.escape.list), 'escape errors in attribute Text with IDs:')
				summary.df[start.row, kColumnNameOutput] <- toString(text.escape.list)
				start.row <- start.row + 1
				start.row
			} else {
				-1
			}
			
			# create sheet only when it does not exists yet
			if (!summary.sheet.name %in% getSheets(workbook)) {
				createSheet(workbook, name = summary.sheet.name)
			}
			writeWorksheet(workbook, summary.df, sheet = summary.sheet.name)
			# output column description width is 100 characters long
			setColumnWidth(workbook, sheet = summary.sheet.name, column = which(colnames(summary.df) == kColumnNameOutput), width = 100 * 256)
			# auto-size for description column
			setColumnWidth(workbook, sheet = summary.sheet.name, column = which(colnames(summary.df) == kColumnNameDescription), width = 30 * 256)
			
			# cell styles for the sheet
			setCellStyle(workbook,
									 sheet = summary.sheet.name,
									 row = 1,
									 col = 1:kColumnLength,
									 cellstyle = kCellStyleHeader)
			if (main.row.error != -1) {
				setCellStyle(workbook,
										 sheet = summary.sheet.name,
										 row = main.row.error,
										 col = 1:kColumnLength,
										 cellstyle = kCellStyleSummaryError)
			}
			if (lang.row.error != -1) {
				setCellStyle(workbook,
										 sheet = summary.sheet.name,
										 row = lang.row.error,
										 col = 1:kColumnLength,
										 cellstyle = kCellStyleSummaryError)
			}
			if (original.text.row.error != -1) {
				setCellStyle(workbook,
										 sheet = summary.sheet.name,
										 row = original.text.row.error,
										 col = 1:kColumnLength,
										 cellstyle = kCellStyleSummaryError)
			}
			if (text.row.error != -1) {
				setCellStyle(workbook,
										 sheet = summary.sheet.name,
										 row = text.row.error,
										 col = 1:kColumnLength,
										 cellstyle = kCellStyleSummaryError)
			}
		}
	}
	# load and return data frame
	LoadTranslation(main.file, english.file, language.file, CreateLanguageSummarySheetHandler(summary.sheet.name))
}
