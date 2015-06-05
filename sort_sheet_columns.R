SortSheetColumns <- function(translation.df,
														 language.file.suffix, # e.g. 'chi' for chinese translation file
														 column.version.current = 'Current',
														 column.version.latest = 'Latest'
														 ) {
	column.names <- colnames(translation.df)
	known.columns <- c(
		Translation$Xls.Column.Other.All,
		list(
			paste(Translation$Xml.File.Suffix.Main1, Translation$Xls.Column.Placeholder.Current, sep='.'),
			paste(Translation$Xml.File.Suffix.Main2, Translation$Xls.Column.Placeholder.Current, sep='.'),
			paste(language.file.suffix, Translation$Xls.Column.Placeholder.Current, sep='.')),
		list(
			paste(Translation$Xml.File.Suffix.Main1, Translation$Xls.Column.Placeholder.Latest, sep='.'),
			paste(Translation$Xml.File.Suffix.Main2, Translation$Xls.Column.Placeholder.Latest, sep='.'),
			paste(language.file.suffix, Translation$Xls.Column.Placeholder.Latest, sep='.'))
	)
	
	unknown.columns <- setdiff(column.names, known.columns)
	df <- if (nrow(translation.df) > 0) {
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
	
	#rename known columns by adding version
	known.columns.with.version <- c(
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
	colnames(df) <- c(known.columns.with.version, unknown.columns)
	df
}