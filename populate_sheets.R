PopulateSheets <- function(workbook, current.translation, latest.translation, sheet.names) {
	sheets <- readWorksheet(workbook, sheet.names)
	
	#################### initialisations, variables, constants, and function definitions ####################
	
	# define cell style and color for Excel for highlighting
	## cell styles for translation sheets
	### define cell style and color for header row
	kCellStyleHeader <- createCellStyle(workbook)
	setFillPattern(kCellStyleHeader, fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleHeader, color = XLC$COLOR.LIGHT_BLUE)
	### define cell style and color for row changes
	kCellStyleRowChanged <- createCellStyle(workbook)
	setWrapText(kCellStyleRowChanged, wrap = T)
	setFillPattern(kCellStyleRowChanged, fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleRowChanged, color = XLC$COLOR.LIGHT_YELLOW)
	### define cell style and color for original text and text changes in a single cell
	kCellStyleTextChanged <- createCellStyle(workbook)
	setWrapText(kCellStyleTextChanged, wrap = T)
	setFillPattern(kCellStyleTextChanged, fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleTextChanged, color = XLC$COLOR.LIGHT_ORANGE)
	### define cell style for wrapping text
	kCellStyleWrapText <- createCellStyle(workbook)
	setWrapText(kCellStyleWrapText, wrap = T)
	## cell styles for summary sheets
	### cell style for row OK
	kCellStyleSummaryOk <<- createCellStyle(workbook)
	setFillPattern(kCellStyleSummaryOk, fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleSummaryOk, color = XLC$COLOR.LIGHT_GREEN)
	### cell style for row ERROR
	kCellStyleSummaryError <<- createCellStyle(workbook)
	setFillPattern(kCellStyleSummaryError, fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleSummaryError, color = XLC$COLOR.RED)
	### cell style for row details ERROR
	kCellStyleSummaryErrorDetails <<- createCellStyle(workbook)
	setFillPattern(kCellStyleSummaryErrorDetails, fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleSummaryErrorDetails, color = XLC$COLOR.LIGHT_ORANGE)
	
	# skip first 3 rows since they contain no keys / only header infos
	kRowIndexFirstKey <- 4
	
	# the columns to be populated and checked
	kColumnNameOriginalText <- 'OriginalText'
	kColumnNameText <- 'Text'
	kColumnNameId <- 'ID'
	kColumnNameKey <- 'Key'
	kColumnNameDescription <- 'Description'
	populate.columns.list <- list(kColumnNameOriginalText, kColumnNameText)
	
	# summary sheet
	kSummarySheetName <- 'Summary Sheets'
	kSummaryColumnNameSheet <- 'Sheet'
	kSummaryColumnNameStatus <- 'Status'
	kSummaryColumnNameStatus <- 'Description'
	kSummaryColumnLength <- 3
	# indicates position for next summary
	summary.row.index <- 1
	# the summary to be filled into the sheet
	summary <- data.frame(
		Sheet = character(),
		Status = character(),
		stringsAsFactors = FALSE
	)
	# create sheet only when it does not exists yet
	if (!kSummarySheetName %in% sheet.names) {
		createSheet(workbook, name = kSummarySheetName)
	}
	
	# function to populate a sheet's cell
	# returns TRUE if it has changed else FALSE
	ValueChanged <- function(cellValue, translationText) {
		# cell has changed based on these rules:
		# - empty/NA/NULL: no translation yet
		# - value differs from data frame
		if (is.null(cellValue) || is.na(cellValue) || cellValue != translationText) {
			TRUE
		} else {
			FALSE
		}
	}
	
	# the lists containing row indices with status ok or error
	summary.row.list.sheet.status.ok <- list()
	summary.row.list.sheet.status.error <- list()
	summary.row.list.sheet.details.error <- list()
	
	#################### END ####################
	
	# process for each sheet:
	# read key in each row
	# based on the key read the translation data frames and populate cells
	for (sheet.name in sheet.names) {
		current.sheet <- readWorksheet(workbook, sheet = sheet.name)
		
		# get col positions required to set cell styles below
		column.names <- colnames(current.sheet)
		# length - last position of last column (cell style for whole row)
		column.length <- length(column.names)
		# position of original text (cell style when it has changed)
		column.index.original.text <- which(column.names == kColumnNameOriginalText)
		# position of text (cell style when it has changed)
		column.index.text <- which(column.names == 'Text')
		# position of id
		column.index.id <- which(column.names == 'ID')
		# position of key
		column.index.key <- which(column.names == 'Key')
		
		# initialize 3 lists holding changed rows:
		# - row list for any cell changes (union of the other two change lists)
		changed.rows <- list()
		# - row list for changes in 'Original' cell
		changed.original.rows <- list()
		# - row list for changes in 'Text' cell
		changed.text.rows <- list()
		
		# summary of processing
		overall.status <- 'OK'
		# write sheet name in summary column 'Sheet'
		summary[summary.row.index, kSummaryColumnNameSheet] <- sheet.name
		# row position of sheet name in summary
		summary.sheet.name.index <- summary.row.index
		# all other infos goes to next line
		summary.row.index <- summary.row.index + 1
		
		# is there any data?
		row.numbers <- nrow(current.sheet)
		if (row.numbers >= kRowIndexFirstKey) {
			print(paste('Populating sheet', sheet.name))
			
			# iterate through each row and start filling the sheet
			# using our translation data frames
			for (row.number in kRowIndexFirstKey:row.numbers) {
				#print(c('Process row', row.number+1))
				
				# read key from sheet
				cell.value <- current.sheet[row.number, 'Key']
				
				# skip if Keycell is NULL or NA
				if (is.null(cell.value) || is.na(cell.value)) {
					next
				}
				
				# get translation for this key
				translation.row <- current.translation[current.translation$Key == cell.value,]
				# there must be exactly one translation
				result <- nrow(translation.row)
				if (result == 1) {
					# set ID in sheet
					current.sheet[row.number, 'ID'] <- translation.row$ID
					
					# boolean marker to indicate some has changed in this sheet's row
					changed.columns <- list()
					
					# fill columns
					for (column.name in populate.columns.list) {
						if (ValueChanged(current.sheet[row.number, column.name], translation.row[column.name])) {
							#print(paste('> change ', column.name, ' \'', current.sheet[row.number, column.name], '\' to \'', translation.row[column.name], '\'', sep = ''))
							
							# set cell value
							current.sheet[row.number, column.name] <- translation.row[column.name]
							
							# store row number in list
							changed.columns <- c(changed.columns, column.name)
						}
					}
					if (length(changed.columns) > 0) {
						#print(paste('Changed row', row.number + 1))
						changed.rows <- c(changed.rows, row.number + 1)
						if (kColumnNameText %in% changed.columns) {
							changed.text.rows <- c(changed.text.rows, row.number + 1)
						}
						if (kColumnNameOriginalText %in% changed.columns) {
							changed.original.rows <- c(changed.original.rows, row.number + 1)
						}
					}
				} else {
					overall.status <- 'Error'
					# fill text in column status
					summary[summary.row.index, kSummaryColumnNameStatus] <- 
						paste('Unknown key \'', 
									cell.value, '\'', 
									sep = '')
					# store row position of error details in list
					summary.row.list.sheet.details.error <- c(summary.row.list.sheet.details.error, summary.row.index + 1)
					# fill text in column description
					summary[summary.row.index, kSummaryColumnNameStatus] <- 
						paste(nrow(translation.row), 
									' translations found in Sheet \'', 
									sheet.name, 
									'\', row ', 
									row.number+1, 
									'.', 
									sep = '')
					summary.row.index <- summary.row.index + 1
					
				}
			}
			
			print('> update worksheet')
			# first write the sheet back into the workbook
			# before doing further changes e.g. cell styles
			writeWorksheet(workbook, current.sheet, sheet = sheet.name)
			
			print('> update styles in sheet')
			# update sheet styles
			# something has changed?
			# highlight the rows with our defined cell styles above
			for (row.number in kRowIndexFirstKey:row.numbers) {
				if (row.number %in% changed.rows) {
					# cell style for complete row has changed
					setCellStyle(workbook,
											 sheet = sheet.name,
											 row = row.number,
											 col = 1:column.length,
											 cellstyle = kCellStyleRowChanged)
					
					# cell style for changed original text
					if (row.number %in% changed.original.rows) {
						setCellStyle(workbook, 
												 sheet = sheet.name, 
												 row = row.number, 
												 col = column.index.original.text, 
												 cellstyle = kCellStyleTextChanged)
					}
					
					# cell style for changed text
					if (row.number %in% changed.text.rows) {
						setCellStyle(workbook, 
												 sheet = sheet.name, 
												 row = row.number, 
												 col = column.index.text, 
												 cellstyle = kCellStyleTextChanged)
					}
				} else {
					# cell style for wrapping text on complete sheet 
					setCellStyle(workbook,
											 sheet = sheet.name,
											 row = row.number,
											 col = 1:column.length,
											 cellstyle = kCellStyleWrapText)
				}
			}
			print('> done populating sheet')
		}
		
		# write output for number of sheet changes in each column
		summary.column.status.text <- ''
		for (column.name in populate.columns.list) {
			if (column.name == kColumnNameText) {
				summary.column.status.text <- paste(summary.column.status.text, ' ', length(changed.text.rows), ' changes in ', column.name, '.', sep = '')
			} else if (column.name == kColumnNameOriginalText) {
				summary.column.status.text <- paste(summary.column.status.text, ' ', length(changed.original.rows), ' changes in ', column.name, '.', sep = '')
			}
		}
		print(paste('>',summary.column.status.text))
		summary[summary.row.index, kSummaryColumnNameStatus] <- summary.column.status.text
		summary[summary.row.index, kSummaryColumnNameStatus] <- paste(length(changed.original.rows) + length(changed.text.rows), ' total changes.')
		
		# set header cell style for each sheet
		setCellStyle(workbook, sheet = sheet.name, row = 1, col = 1:column.length, cellstyle = kCellStyleHeader)
		summary.row.index <- summary.row.index + 1
		if (overall.status == 'OK') {
			summary.row.list.sheet.status.ok <- c(summary.row.list.sheet.status.ok, summary.sheet.name.index + 1)
		} else {
			summary.row.list.sheet.status.error <- c(summary.row.list.sheet.status.error, summary.sheet.name.index + 1)
		}
		summary[summary.sheet.name.index, kSummaryColumnNameStatus] <- overall.status
		
		# set column widths
		# 1st column description width is 10 characters long
		setColumnWidth(workbook, sheet = sheet.name, column = which(column.names == kColumnNameDescription), width = 10 * 256)
		for (i in 2:column.length) {
			# auto-size for id and key column
			# width 20 chars long for all other columns
			columnWidth <- if (i == column.index.id || i == column.index.key) -1 else 20 * 256
			setColumnWidth(workbook, sheet = sheet.name, column = i, width = columnWidth)
		}
	} # end of processing sheets in for-loop 
	
	# first write data before updating style sheet
	writeWorksheet(workbook, summary, sheet = kSummarySheetName)
	setCellStyle(workbook,
							 sheet = kSummarySheetName,
							 row = 1,
							 col = 1:kSummaryColumnLength,
							 cellstyle = kCellStyleHeader)
	for (i in 2:nrow(summary)) {
		# set style for status ok or error
		if (i %in% summary.row.list.sheet.status.ok) {
			setCellStyle(workbook,
									 sheet = kSummarySheetName,
									 row = i,
									 col = 1:kSummaryColumnLength,
									 cellstyle = kCellStyleSummaryOk)
		} else if (i %in% summary.row.list.sheet.status.error) {
			setCellStyle(workbook,
									 sheet = kSummarySheetName,
									 row = i,
									 col = 1:kSummaryColumnLength,
									 cellstyle = kCellStyleSummaryError)
		} else if (i %in% summary.row.list.sheet.details.error) {
			setCellStyle(workbook,
									 sheet = kSummarySheetName,
									 row = i,
									 col = 2:kSummaryColumnLength,
									 cellstyle = kCellStyleSummaryErrorDetails)
		}
	}
	setColumnWidth(workbook, sheet = kSummarySheetName, column = 1:3, width = -1)
	# seems not to work: set active sheet...
	setActiveSheet(workbook, sheet = kSummarySheetName)
}