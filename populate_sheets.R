PopulateSheets <- function(workbook, 
													 current.translation, 
													 latest.translation, 
													 language.file.suffix, # e.g. 'chi' for chinese translation file
													 sheet.names, 
													 column.version.current = 'Current',
													 column.version.latest = 'Latest',
													 start.at.row = 1) {
	source('constants.R')
	source('sort_sheet_columns.R')
	sheets <- readWorksheet(workbook, sheet.names)
	
	#################### initialisations, variables, constants, and function definitions ####################
	
	# define cell style and color for Excel for highlighting
	## cell styles for translation sheets
	### define cell style and color for header row
	kCellStyleHeader <- createCellStyle(workbook)
	setFillPattern(kCellStyleHeader, 
								 fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleHeader, 
												 color = XLC$COLOR.LIGHT_BLUE)
	### define cell style and color for row error
	kCellStyleRowError <- createCellStyle(workbook)
	setWrapText(kCellStyleRowError, 
							wrap = T)
	setFillPattern(kCellStyleRowError, 
								 fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleRowError, 
												 color = XLC$COLOR.ROSE)
	### define cell style and color for row changes
	kCellStyleRowChanged <- createCellStyle(workbook)
	setWrapText(kCellStyleRowChanged, 
							wrap = T)
	setFillPattern(kCellStyleRowChanged, 
								 fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleRowChanged, 
												 color = XLC$COLOR.LIGHT_YELLOW)
	### define cell style and color for cell error
	kCellStyleCellError <- createCellStyle(workbook)
	setWrapText(kCellStyleCellError, 
							wrap = T)
	setFillPattern(kCellStyleCellError, 
								 fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleCellError, 
												 color = XLC$COLOR.RED)
	### define cell style and color for cell changes
	kCellStyleCellChanged <- createCellStyle(workbook)
	setWrapText(kCellStyleCellChanged, 
							wrap = T)
	setFillPattern(kCellStyleCellChanged, 
								 fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleCellChanged, 
												 color = XLC$COLOR.LIGHT_ORANGE)
	### define cell style for wrapping text
	kCellStyleWrapText <- createCellStyle(workbook)
	setWrapText(kCellStyleWrapText, wrap = T)
	## cell styles for summary sheets
	### cell style for row OK
	kCellStyleSummaryOk <- createCellStyle(workbook)
	setFillPattern(kCellStyleSummaryOk, 
								 fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleSummaryOk, 
												 color = XLC$COLOR.LIGHT_GREEN)
	### cell style for row CHANGED (Old/New keys)
	kCellStyleSummaryChanged <- createCellStyle(workbook)
	setFillPattern(kCellStyleSummaryChanged, 
								 fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleSummaryChanged, 
												 color = XLC$COLOR.YELLOW)
	### cell style for row details CHANGED (Old/New keys)
	kCellStyleSummaryChangedDetails <- createCellStyle(workbook)
	setFillPattern(kCellStyleSummaryChangedDetails, 
								 fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleSummaryChangedDetails, 
												 color = XLC$COLOR.LIGHT_YELLOW)
	### cell style for row ERROR
	kCellStyleSummaryError <- createCellStyle(workbook)
	setFillPattern(kCellStyleSummaryError, 
								 fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleSummaryError, 
												 color = XLC$COLOR.RED)
	### cell style for row details ERROR
	kCellStyleSummaryErrorDetails <- createCellStyle(workbook)
	setFillPattern(kCellStyleSummaryErrorDetails, 
								 fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleSummaryErrorDetails, 
												 color = XLC$COLOR.LIGHT_ORANGE)
	
	# summary sheet
	kSummarySheetName <- 'Summary Sheets'
	kSummaryColumnNameSheet <- 'Sheet'
	kSummaryColumnNameStatus <- 'Status'
	kSummaryColumnNameDescription <- 'Description'
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
	
	# the lists containing row indices with status ok or error
	summary.row.list.sheet.status.ok <- list()
	summary.row.list.sheet.status.oldnew <- list()
	summary.row.list.sheet.details.oldnew <- list()
	summary.row.list.sheet.status.error <- list()
	summary.row.list.sheet.details.error <- list()
	
	#################### END ####################
	
	# process for each sheet:
	# read key in each row
	# based on the key read the translation data frames and populate cells
	for (sheet.name in sheet.names) {
		current.sheet <- readWorksheet(workbook, sheet = sheet.name)
		current.sheet <- SortSheetColumns(current.sheet, language.file.suffix, column.version.current, column.version.latest)
		
		# initialize cell style list holding changed and error rows
		cell.style.rows.df <- 
			read.table(text=paste(Translation$Internal.Style.Columns.All, collapse=' '), 
								 header=T)
		
		# summary of processing
		overall.status <- Translation$Xls.Sheet.Summary.StatusCode.Ok
		# write sheet name in summary column 'Sheet'
		summary[summary.row.index, kSummaryColumnNameSheet] <- sheet.name
		# row position of sheet name in summary
		summary.sheet.name.index <- summary.row.index
		# all other infos goes to next line
		summary.row.index <- summary.row.index + 1
		
		xls.text.columns <- list(Translation$Xml.File.Suffix.Main1, 
														 Translation$Xml.File.Suffix.Main2,
														 language.file.suffix)
		# is there any data?
		row.numbers <- nrow(current.sheet)
		if (row.numbers >= start.at.row) {
			print(paste('Populating sheet', sheet.name))
			
			# iterate through each row and start filling the sheet
			# using our translation data frames
			for (row.number in start.at.row:row.numbers) {
				#print(c('Process row', row.number+1))
				
				# read key from sheet
				cell.value <- current.sheet[row.number, Translation$Xml.Attribute.Key]
				
				# skip if Keycell is NULL or NA
				if (is.null(cell.value) || is.na(cell.value)) {
					next
				}
				
				# get translation for this key
				current.translation.row <- 
					current.translation[current.translation$Key == cell.value,]
				latest.translation.row <- 
					latest.translation[latest.translation$Key == cell.value,]
				# there must be exactly one translation
				current.nrow <- nrow(current.translation.row)
				latest.nrow <- nrow(latest.translation.row)
				# flag to make sure it appears only once in the summary sheet
				# and not 3 times because of 3 text columns
				oldnew.summary <- FALSE
				if (current.nrow == 1 || latest.nrow == 1) {
					# set ID in sheet
					current.sheet[row.number, 'ID'] <- if (current.nrow == 1) {
						current.translation.row$ID
					} else {
						latest.translation.row$ID
					}
					
					# fill columns
					for (column.name in xls.text.columns) {
						# set cell value for current translation
						current.sheet[row.number, paste(column.name, column.version.current)] <- 
							if (current.nrow == 1) {
								current.translation.row[column.name]
							} else {
								paste('[NEW key added in release ', column.version.latest, ', but not existent in', column.version.current, '!]')
							}
						# set cell value for latest translation
						current.sheet[row.number, paste(column.name, column.version.latest)] <- 
							if (latest.nrow == 1) {
								latest.translation.row[column.name]
							} else {
								paste('[OLD key removed in release ', column.version.latest, ', but existent in', column.version.current, '!]')
							}
						
						# key exists only in current or latest?
						key.in.current.latest <- current.nrow == 1 & latest.nrow == 1
						if (!oldnew.summary && !key.in.current.latest) {
							oldnew.summary <- TRUE
							# fill text in column status
							summary[summary.row.index, kSummaryColumnNameStatus] <- 
								paste('Old/New key \'', 
											cell.value, '\'', 
											sep = '')
							# fill text in column description
							summary[summary.row.index, kSummaryColumnNameDescription] <- 
								if (current.nrow != 1) {
									paste('NEW key added in release ', 
												column.version.latest, 
												', but not existent in', 
												column.version.current,
												', sheet row', row.number+1)
								} else {
									paste('OLD key in release ',
												column.version.latest, 
												', but removed in', 
												column.version.current,
												', sheet row', row.number+1)
								}
							# store row position of old/new key details in list
							summary.row.list.sheet.details.oldnew <- 
								c(summary.row.list.sheet.details.oldnew, summary.row.index + 1)
							summary.row.index <- summary.row.index + 1
						}
						# cell style in case:
						## - key does not exist either in current or latest version
						## - text is different between current and latest
						if (!key.in.current.latest || # no key
									latest.translation.row[column.name] != # text differences
								  current.translation.row[column.name]) {
							#if (!key.in.current.latest) print(paste('> old/new key ', column.name, ' in row ', row.number))
							#print(paste('> change ', column.name, ' in row ', row.number, ' \'', current.translation.row[column.name], '\' to \'', latest.translation.row[column.name], '\'', sep = ''))
							
							# store style for changed cell
							cell.style.rows.df[nrow(cell.style.rows.df) + 1, ] <- 
								c(column.name, row.number + 1, Translation$Internal.Style.Df.Style.Type.Cell)
						}
					}
				} else {
					overall.status <- Translation$Xls.Sheet.Summary.StatusCode.Error
					# fill text in column status
					summary[summary.row.index, kSummaryColumnNameStatus] <- 
						paste('Unknown key \'', 
									cell.value, '\'', 
									sep = '')
					# store row position of error details in list
					summary.row.list.sheet.details.error <- 
						c(summary.row.list.sheet.details.error, summary.row.index + 1)
					# fill text in column description
					summary[summary.row.index, kSummaryColumnNameDescription] <- 
						paste(current.nrow, ' current and ', latest.nrow, 
									' latest translations found in sheet, row ', 
									row.number+1, 
									'.', 
									sep = '')
					summary.row.index <- summary.row.index + 1
					# store style for row error
					cell.style.rows.df[nrow(cell.style.rows.df) + 1, ] <- 
						# add any value (e.g. key) in first column and not an empty ''
						# otherwise NA is added and an error occurs
						c(Translation$Xml.Attribute.Key, row.number + 1, Translation$Internal.Style.Df.Style.Type.Row)
				}
			}
			
			print('> update worksheet')
			# first write the sheet back into the workbook
			# before doing further changes e.g. cell styles
			writeWorksheet(workbook, current.sheet, sheet = sheet.name)
			
			# TODO: data frame elements are for some reason a list with a single value
			# here we need to transform into new data frame by
			# extracting each value and defining type (as.numeric, etc.)
			cell.style.rows.df <- transform(cell.style.rows.df,
																			column.name = as.character(column.name),
																			row.number = as.numeric(row.number),
																			style.type = as.numeric(style.type))
			
			print('> update styles in sheet')
			# update sheet styles
			# something has changed?
			# highlight the rows with our defined cell styles above
			for (row.number in start.at.row: (row.numbers + 1)) {
				# row has extra style?
				# nb: there can be more rows for a row number since
				# several columns can be changed
				# first check if there are rows (nrow) because
				# calling which on empty causes an error
				style.row.indices <- if (nrow(cell.style.rows.df) == 0) {
					integer(0)
				} else {
					which(row.number == cell.style.rows.df[Translation$Internal.Style.Df.Row.Number])
				}
				if (length(style.row.indices) > 0) {
					# cell style for complete row has changed
					setCellStyle(workbook,
											 sheet = sheet.name,
											 row = row.number,
											 col = 1:length(colnames(current.sheet)),
											 cellstyle = kCellStyleRowChanged)
					# set styles for some cells this row
					for (style.row.index in style.row.indices) {
						# get row style for each cell
						style.row <- cell.style.rows.df[style.row.index,]
						if (style.row[Translation$Internal.Style.Df.Style.Type] == Translation$Internal.Style.Df.Style.Type.Cell) {
							#print(paste('> cell style for column', paste(style.row[[Translation$Internal.Style.Df.Column.Name]],column.version.current), ' in row', row.number))
							
							# cell style for current translation release column
							setCellStyle(workbook, 
													 sheet = sheet.name, 
													 row = row.number,
													 col = which(colnames(current.sheet) ==
													 							paste(style.row[[Translation$Internal.Style.Df.Column.Name]],
													 											 column.version.current)),
													 cellstyle = kCellStyleCellChanged)

							#print(paste('> cell style for column', paste(style.row[[Translation$Internal.Style.Df.Column.Name]],column.version.latest), ' in row', row.number))
							# cell style for latest translation release column
							setCellStyle(workbook, 
													 sheet = sheet.name, 
													 row = row.number,
													 col = which(colnames(current.sheet) ==
													 							paste(style.row[[Translation$Internal.Style.Df.Column.Name]],
													 										 column.version.latest)),
													 cellstyle = kCellStyleCellChanged)
						} else {
							# cell style for complete row error
							setCellStyle(workbook,
													 sheet = sheet.name,
													 row = style.row[[Translation$Internal.Style.Df.Row.Number]],
													 col = 1:length(colnames(current.sheet)),
													 cellstyle = kCellStyleRowError)
							# cell style for error in Key column
							setCellStyle(workbook,
													 sheet = sheet.name,
													 row = style.row[[Translation$Internal.Style.Df.Row.Number]],
													 col = which(colnames(current.sheet) == 
													 							style.row[[Translation$Internal.Style.Df.Column.Name]]),
													 cellstyle = kCellStyleCellError)
							#stop for loop since complete row is marked
							break;
						}
					}
				} else {
					# ATTENTION: this slows down the perfomance!!!
					# cell style for wrapping text on complete sheet 
					setCellStyle(workbook,
											 sheet = sheet.name,
											 row = row.number,
											 col = 1:length(colnames(current.sheet)),
											 cellstyle = kCellStyleWrapText)
				}
			}
			print('> done populating sheet')
		} else {
			# save unchanged worksheet anyways since we have re-ordered columns
			writeWorksheet(workbook, current.sheet, sheet = sheet.name)
		}
		
		# write output for number of sheet changes in each column
		total.changes <- 0
		summary.column.status.text <- ''
		for (column.name in xls.text.columns) {
			# first check if there are rows (nrow) because
			# calling which on empty causes an error
			changes <- if (nrow(cell.style.rows.df) == 0) {
				0
			} else {
				length(which(cell.style.rows.df[Translation$Internal.Style.Df.Column.Name] == column.name))
			}
			total.changes <- total.changes + changes
			summary.column.status.text <- paste(summary.column.status.text,
																					' ', changes,
																					' change(s) in ',
																					column.name, '.',
																					sep = '')
		}
		print(paste('>',summary.column.status.text))
		summary[summary.row.index, kSummaryColumnNameDescription] <- 
			summary.column.status.text
		summary[summary.row.index, kSummaryColumnNameStatus] <- 
			paste(total.changes,
						' total changes.')
		
		# set header cell style for each sheet
		setCellStyle(workbook,
								 sheet = sheet.name,
								 row = 1,
								 col = 1:length(colnames(current.sheet)),
								 cellstyle = kCellStyleHeader)
		summary.row.index <- summary.row.index + 1
		if (overall.status == Translation$Xls.Sheet.Summary.StatusCode.Ok) {
			summary.row.list.sheet.status.ok <-
				c(summary.row.list.sheet.status.ok,
					summary.sheet.name.index + 1)
		} else {
			summary.row.list.sheet.status.error <- 
				c(summary.row.list.sheet.status.error,
					summary.sheet.name.index + 1)
		}
		summary[summary.sheet.name.index, kSummaryColumnNameStatus] <- overall.status
		
		# set column widths
		# 1st column description width is 10 characters long
		setColumnWidth(workbook, 
									 sheet = sheet.name, 
									 column = 
									 	which(colnames(current.sheet) == Translation$Xls.Column.Other.Description), 
									 width = 10 * 256)
		for (i in 2:length(colnames(current.sheet))) {
			# auto-size for id and key column
			# width 20 chars long for all other columns
			columnWidth <- 
				if (i == which(colnames(current.sheet) == Translation$Xml.Attribute.Id)) {
					10 * 256
				} else if (i == which(colnames(current.sheet) == Translation$Xml.Attribute.Key)) {
					15 * 256
				} else {
					20 * 256
				}
			setColumnWidth(workbook, 
										 sheet = sheet.name, 
										 column = i, 
										 width = columnWidth)
		}
	} # end of processing sheets in for-loop 
	
	# first write data before updating style sheet
	writeWorksheet(workbook, 
								 summary, 
								 sheet = kSummarySheetName)
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
		} else if (i %in% summary.row.list.sheet.status.oldnew) {
			setCellStyle(workbook,
									 sheet = kSummarySheetName,
									 row = i,
									 col = 1:kSummaryColumnLength,
									 cellstyle = kCellStyleSummaryChanged)
		} else if (i %in% summary.row.list.sheet.details.oldnew) {
			setCellStyle(workbook,
									 sheet = kSummarySheetName,
									 row = i,
									 col = 2:kSummaryColumnLength,
									 cellstyle = kCellStyleSummaryChangedDetails)
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
	# sets the active sheet
	# (though the tab itself is not focused and highlighted...)
	setActiveSheet(workbook, sheet = kSummarySheetName)
}