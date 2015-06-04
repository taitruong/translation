# problem loading rJava, http://stackoverflow.com/a/9120712
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk1.7.0_75\\jre')
require(XLConnect)

CreateQaTranslationWorkbook <- function(excel.file, 
																				current.file.dir,
																				latest.file.dir,
																				language.file.suffix, # e.g. 'chi' for chinese translation file
																				column.version.current = 'Current',
																				column.version.latest = 'Latest',
																				module.file.prefix = 'recruitingappTranslation' # e.g. 'recruiting' or 'employee'
																				) {
	source('create_summary_sheets.R')
	source('populate_sheets.R')

	# read excel workbook and its sheets
	print(paste('Reading workbook', excel.file))
	workbook <- loadWorkbook(excel.file)
	# get the sheet names before creating the summary sheets because
	# we do not want to populate these
	sheet.names <- getSheets(workbook)
	# do not use styles e.g. when calling writeWorksheet()
	setStyleAction(workbook, XLC$STYLE_ACTION.NONE)
	print('Reading current translation files')
	current.translation <- 
		LoadTranslationAndCreateSummarySheet(current.file.dir, 
																				 language.file.suffix, 
																				 module.file.prefix, 
																				 workbook, 
																				 Translation$Xls.Sheet.Summary.Current)
	
	print('Reading latest translation files')
	latest.translation <- 
		LoadTranslationAndCreateSummarySheet(latest.file.dir, 
																				 language.file.suffix, 
																				 module.file.prefix, 
																				 workbook, 
																				 Translation$Xls.Sheet.Summary.Latest)
	
	PopulateSheets(workbook, 
								 current.translation, 
								 latest.translation, 
								 language.file.suffix,
								 sheet.names,
								 column.version.current,
								 column.version.latest,
								 start.at.row = 4) 	# skip first 3 rows since they contain no keys / only header infos
	
	# save result
	result.filename <- paste('new_', excel.file, sep='')
	print(paste('Saving', result.filename))
	saveWorkbook(workbook, result.filename)
	
}