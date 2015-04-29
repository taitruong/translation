# problem loading rJava, http://stackoverflow.com/a/9120712
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk1.7.0_75\\jre')
require(XLConnect)

CreateQaTranslationWorkbook <- function(excel.file, 
																	current.main.file,
																	current.english.file,
																	current.language.file,
																	latest.main.file,
																	latest.english.file,
																	latest.language.file) {
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
	
	PopulateSheets(workbook, 
								 current.translation, 
								 latest.translation, 
								 sheet.names,
								 start.at.row = 4) 	# skip first 3 rows since they contain no keys / only header infos
	
	# save result
	result.filename <- paste('new_', excel.file)
	print(paste('Saving', result.filename))
	saveWorkbook(workbook, result.filename)
	
}