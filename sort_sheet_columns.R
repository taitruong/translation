Sort.Sheet.Columns <- function(translation.df) {
	column.names <- colnames(translation.df)
	unknown.columns <- setdiff(column.names,Translation$Xls.Column.All)
	if (nrow(translation.df) > 0) {
		for (column in Translation$Xls.Column.All) {
			# first add missing columns
			if (!column %in% column.names) {
				translation.df[column] <- NA
			}
		}
		# now re-order columns
		translation.df[as.character(c(Translation$Xls.Column.All, unknown.columns))]
	}	else {
		column.names <- colnames(translation.df)
		read.table(text=as.character(c(Translation$Xls.Column.All, unknown.columns)), header=T)
	}
}