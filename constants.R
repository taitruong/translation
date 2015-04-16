Translation <- list()

Translation$"Xml.Attribute.Original.Text" <- 'OriginalText'
Translation$"Xml.Attribute.English.Text" <- 'EnglishText'
Translation$"Xml.Attribute.Text" <- 'Text'
Translation$"Xml.Attribute.Id" <- 'ID'
Translation$"Xml.Attribute.Key" <- 'Key'

Translation$"Xls.Column.Current.All" <- list(Translation$Xml.Attribute.Original.Text, 
																						 Translation$Xml.Attribute.English.Text, 
																						 Translation$Xml.Attribute.Text)

Translation$"Xls.Column.Suffix.Latest" <- 'Latest'
Translation$"Xls.Column.Latest.All" <- 
	sapply(Translation$Xls.Column.Current.All,
				 function(column) {
				 	paste(column, Translation$Xls.Column.Suffix.Latest, sep = '')
				 })

Translation$"Xls.Column.Other.Description" <- 'Description' 
Translation$"Xls.Column.Other.Id" <- Translation$Xml.Attribute.Id
Translation$"Xls.Column.Other.Key" <- Translation$Xml.Attribute.Key
Translation$"Xls.Column.Other.All" <- list(Translation$Xls.Column.Other.Description, 
																						 Translation$Xls.Column.Other.Id, 
																						 Translation$Xls.Column.Other.Key)

# keep this order!
Translation$"Xls.Column.All" <- c(Translation$Xls.Column.Other.All,
																	Translation$Xls.Column.Current.All, 
																	Translation$Xls.Column.Latest.All)

Translation$"Internal.Style.Df.Column.Name" <- 'column.name'
Translation$"Internal.Style.Df.Row.Number" <- 'row.number'
Translation$"Internal.Style.Df.Style.Type" <- 'style.type'
Translation$"Internal.Style.Columns.All" <- list(Translation$Internal.Style.Df.Column.Name,
																								 Translation$Internal.Style.Df.Row.Number,
																								 Translation$Internal.Style.Df.Style.Type)

Translation$"Internal.Style.Df.Style.Type.Row" <- 1
Translation$"Internal.Style.Df.Style.Type.Cell" <- 2
