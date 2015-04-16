# problem loading rJava, http://stackoverflow.com/a/9120712
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk1.7.0_75\\jre')
require(XLConnect)

FillTranslationWorkbook <- function(excel.file, 
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
	current.translation <- LoadTranslationAndCreateSummarySheet(current.main.file, current.english.file, current.language.file, workbook, 'Summary Current Translation')
	
	print('Reading latest translation files')
	latest.translation <- LoadTranslationAndCreateSummarySheet(latest.main.file, latest.english.file, latest.language.file, workbook, 'Summary Latest Translation')
	
	PopulateSheets(workbook, current.translation, latest.translation, sheet.names)
	
	# save result
	saveWorkbook(workbook, paste('new_', excel.file))
	
}