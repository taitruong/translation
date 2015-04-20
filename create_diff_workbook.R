# problem loading rJava, http://stackoverflow.com/a/9120712
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk1.7.0_75\\jre')
require(XLConnect)

CreateDiffWorkbook <- function(current.main.file,
															 current.english.file,
															 current.language.file,
															 latest.main.file,
															 latest.english.file,
															 latest.language.file,
															 outputFilename = 'DiffOutput.xlsx') {
	source('constants.R')
	source('sort_sheet_columns.R')
	source('create_summary_sheets.R')
	
	print(paste('Loading template workbook', Translation$Xls.Diff.Template.File))
	workbook <- loadWorkbook(Translation$Xls.Diff.Template.File)
	# do not use styles e.g. when calling writeWorksheet()
	setStyleAction(workbook, XLC$STYLE_ACTION.NONE)
	
	# define cell styles
	## define cell style and color for header row
	kCellStyleHeader <- createCellStyle(workbook)
	setFillPattern(kCellStyleHeader, 
								 fill = XLC$FILL.SOLID_FOREGROUND)
	setFillForegroundColor(kCellStyleHeader, 
												 color = XLC$COLOR.LIGHT_BLUE)
	### define cell style and color for wrap text
	kCellStyleWrap <- createCellStyle(workbook)
	setWrapText(kCellStyleWrap, 
							wrap = T)
	
	print('Reading current translation files')
	current.translation <- 
		LoadTranslationAndCreateSummarySheet(current.main.file, 
																				 current.english.file, 
																				 current.language.file, 
																				 workbook, 
																				 Translation$Xls.Sheet.Summary.Current)
	
	print('Reading latest translation files')
	latest.translation <- 
		LoadTranslationAndCreateSummarySheet(latest.main.file, 
																				 latest.english.file, 
																				 latest.language.file, 
																				 workbook, 
																				 Translation$Xls.Sheet.Summary.Latest)
	# 
	print(paste('Create diff sheet:', Translation$Xls.Diff.Sheet.Name))
	
	# create sheet in case it does not exist yet
	if (!Translation$Xls.Diff.Sheet.Name %in% getSheets(workbook)) {
		createSheet(workbook, Translation$Xls.Diff.Sheet.Name)
	}
	diff.sheet <- readWorksheet(workbook, Translation$Xls.Diff.Sheet.Name)
	diff.sheet <- Sort.Sheet.Columns(diff.sheet)
	
	print('Checking differences between current and latest translations')
	# merge by id
	# here we don't need keys, but we keep key in one df for output
	latest.translation[Translation$Xml.Attribute.Key] <- NULL
	current.latest <- merge(current.translation,
													latest.translation,
													by = Translation$Xml.Attribute.Id,
													suffix = c('', Translation$Xls.Column.Suffix.Latest))
	diff.row <- 1
	for (row in 1:nrow(current.latest)) {
		added <- FALSE
		for (column.current in Translation$Xls.Text.Columns) {
			column.latest <- paste(column.current, Translation$Xls.Column.Suffix.Latest, sep='')
			# diff check
			if ((is.na(current.latest[row, column.current]) && !is.na(current.latest[row, column.latest]))
					|| (!is.na(current.latest[row, column.current]) && is.na(current.latest[row, column.latest]))
					|| current.latest[row, column.current] != current.latest[row, column.latest]) {
				# add id and key
				if (!added) {
					diff.sheet[diff.row, Translation$Xls.Column.Other.Id] <- 
						current.latest[row, Translation$Xml.Attribute.Id]
					diff.sheet[diff.row, Translation$Xls.Column.Other.Key] <- 
						current.latest[row, Translation$Xml.Attribute.Key]
					added <- TRUE
				}
				
				diff.sheet[diff.row, column.current] <- 
					current.latest[row, column.current]
				diff.sheet[diff.row, column.latest] <- 
					current.latest[row, column.latest]
			}
		}
		if (added) {
			diff.row <- diff.row + 1
		}
	}

	# remove column description - we don't need it
	print(paste('Removing column', Translation$Xls.Column.Other.Description))
	diff.sheet[Translation$Xls.Column.Other.Description] <- NULL

	# first write the sheet back into the workbook
	# before doing further changes e.g. cell styles
	print('Write differences in worksheet')
	writeWorksheet(workbook, diff.sheet, sheet = Translation$Xls.Diff.Sheet.Name)

	# cell styles
	# set header cell style
	print('Set cell styles')
	setCellStyle(workbook,
							 sheet = Translation$Xls.Diff.Sheet.Name,
							 row = 1,
							 col = 1:length(colnames(diff.sheet)),
							 cellstyle = kCellStyleHeader)
	# wrap text cell style for all text columns below haea
	for (diff.row in 1: (nrow(diff.sheet) - 1)) {
		setCellStyle(workbook,
								 sheet = Translation$Xls.Diff.Sheet.Name,
								 row = diff.row + 1,
								 col = length(Translation$Xls.Column.Other.All):length(colnames(diff.sheet)),
								 cellstyle = kCellStyleWrap)
	}
	
	# set column widths
	print('Set column widths')
	for (column.name in colnames(diff.sheet)) {
		if (column.name == Translation$Xls.Column.Other.Key) {
			# set key column widths 30 chars long
			setColumnWidth(workbook, 
										 sheet = Translation$Xls.Diff.Sheet.Name, 
										 column = 
										 	which(colnames(diff.sheet) == column.name), 
										 width = 30 * 256)
		} else if (column.name %in% Translation$Xls.Column.Other.All) {
			# auto-size for description, id, and key
			setColumnWidth(workbook, 
										 sheet = Translation$Xls.Diff.Sheet.Name, 
										 column = 
										 	which(colnames(diff.sheet) == column.name), 
										 width = -1)
		} else {
			# set text column widths 20 chars long
			setColumnWidth(workbook, 
										 sheet = Translation$Xls.Diff.Sheet.Name, 
										 column = 
										 	which(colnames(diff.sheet) == column.name), 
										 width = 20 * 256)
			
		}
	}
	
	# sets the active sheet
	# (though the tab itself is not focused and highlighted...)
	print(paste(Translation$Xls.Sheet.Summary.Current, ': make active sheet'))
	setActiveSheet(workbook, sheet = Translation$Xls.Sheet.Summary.Current)
	
	print(paste('Saving workbook to', outputFilename))
	saveWorkbook(workbook, outputFilename)
}