SortSheetColumns <- function(translation.df,
														 language.file.suffix, # e.g. 'chi' for chinese translation file
														 column.version.current = 'Current',
														 column.version.latest = 'Latest'
														 ) {
	column.names <- colnames(translation.df)
	known.columns <- c(
		Translation$Xls.Column.Other.All,
		list(
			paste(Translation$Xml.File.Suffix.Main1, column.version.current),
			paste(Translation$Xml.File.Suffix.Main2, column.version.current),
			paste(language.file.suffix, column.version.current)),
		list(
			paste(Translation$Xml.File.Suffix.Main1, column.version.latest),
			paste(Translation$Xml.File.Suffix.Main2, column.version.latest),
			paste(language.file.suffix, column.version.latest))
	)
	
	unknown.columns <- setdiff(column.names, known.columns)
	if (nrow(translation.df) > 0) {
		for (column in known.columns) {
			# first add missing columns
			if (!column %in% column.names) {
				translation.df[column] <- NA
			}
		}
		# now re-order columns
		translation.df[as.character(c(known.columns, unknown.columns))]
	}	else {
		column.names <- colnames(translation.df)
		read.table(text=as.character(c(known.columns, unknown.columns)), header=T)
	}
}