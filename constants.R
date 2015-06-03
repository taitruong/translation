Translation <- list()

Translation$Xml.File.Suffix.Main1 <- 'Main'
Translation$Xml.File.Suffix.Main2 <- 'eng'

Translation$Xml.Attribute.Original.Text <- 'OriginalText'
Translation$Xml.Attribute.Text <- 'Text'
Translation$Xml.Attribute.Id <- 'ID'
Translation$Xml.Attribute.Key <- 'Key'

Translation$Df.Attribute.Main1 <- Translation$Xml.File.Suffix.Main1
Translation$Df.Attribute.Main2 <- Translation$Xml.File.Suffix.Main2

Translation$Xls.Column.Suffix.Latest <- 'Latest'
Translation$Xls.Latest.Text.Columns <- 
	sapply(Translation$Xls.Text.Columns,
				 function(column) {
				 	paste(column, Translation$Xls.Column.Suffix.Latest, sep = '')
				 })

Translation$Xls.Column.Other.Description <- 'Description' 
Translation$Xls.Column.Other.Id <- Translation$Xml.Attribute.Id
Translation$Xls.Column.Other.Key <- Translation$Xml.Attribute.Key
Translation$Xls.Column.Other.All <- list(Translation$Xls.Column.Other.Description, 
																				 Translation$Xls.Column.Other.Id, 
																				 Translation$Xls.Column.Other.Key)

# keep this order!
Translation$Xls.Column.All <- c(Translation$Xls.Column.Other.All,
																Translation$Xls.Text.Columns, 
																Translation$Xls.Latest.Text.Columns)

Translation$Xls.Sheet.Summary.Current <- 'Summary Current Translation'
Translation$Xls.Sheet.Summary.Latest <- 'Summary Latest Translation'
Translation$Xls.Sheet.Summary.StatusCode.Ok <- 'OK'
Translation$Xls.Sheet.Summary.StatusCode.Error <- 'Error'

Translation$Internal.Style.Df.Column.Name <- 'column.name'
Translation$Internal.Style.Df.Row.Number <- 'row.number'
Translation$Internal.Style.Df.Style.Type <- 'style.type'
Translation$Internal.Style.Columns.All <- list(Translation$Internal.Style.Df.Column.Name,
																							 Translation$Internal.Style.Df.Row.Number,
																							 Translation$Internal.Style.Df.Style.Type)

Translation$Internal.Style.Df.Style.Type.Row <- 1
Translation$Internal.Style.Df.Style.Type.Cell <- 2

Translation$Xls.Diff.Template.File <- 'TranslationTemplate.xlsx'
Translation$Xls.Diff.Sheet.Name <- 'Diff'
