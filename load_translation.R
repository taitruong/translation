# problem loading rJava, http://stackoverflow.com/a/9120712
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk1.7.0_75\\jre')
require(XML)

# Creates a data frame with the columns ID, Key, and Text
LoadTranslation <- function(main.file,
														english.file,
														language.file,
														translation.handler = NULL # handler params:
														# result data frame: contains merged df all 3 files
														# main data frame
														# english data frame
														# language data frame
) {
	
	# extracts the attribute using toString() on translationTag and 
	# substrings the value from the part between attribute='value'
	Get.Raw.Attribute <- function(raw.string, attribute) {
		# extract text after 'attribute="'
		attribute.start <- paste(attribute, '=\"', sep='')
		raw.string <- sub(attribute.start, 
								'', 
								substring(raw.string, regexpr(attribute.start, raw.string)))
		# extract text until '='
		substring(raw.string, 1, regexpr('\"', raw.string) - 1)
	}
	
	# Base class loading XML file with Translation tags
	Load.Xml <- function(filename, getAttributes) {
		doc <- xmlParse(filename)
		root <- xmlRoot(doc)
		# read all Translation tags
		df <- as.data.frame(t(xpathSApply(root, "//*/Translation", getAttributes)), 
												stringsAsFactors = FALSE)
	}
	
	# Creates a data frame with the columns ID, Key, and OriginalText
	Load.Main <- function(filename) {
		df <- Load.Xml(filename, getAttributes <- function(translationTag){
			# get attributes
			attributes <- xmlAttrs(translationTag)
			attribute.names <- names(attributes)
			
			# get raw string of whole node
			raw.string <- toString.XMLNode(translationTag)

			# get attributes
			if (!Translation$Xml.Attribute.Id %in% attribute.names) 
				stop(c(filename, 
							 ': Missing attribute ', 
							 Translation$Xml.Attribute.Id, 
							 ' in:', 
							 raw.string))
			# TODO @tt cannot convert ID to numeric (as.numeric(xmlGetAttr('ID')))
			# since in the code below when converted to a data frame it is a list - check later
			id <- as.numeric(Get.Raw.Attribute(raw.string, Translation$Xml.Attribute.Id)) 
			
			if (!Translation$Xml.Attribute.Key %in% attribute.names) 
				stop(c(filename, 
							 ': Missing attribute Key in:', 
							 raw.string))
			# workaround retrieving attribute via 'raw' toString function 
			# for some reason retrieving value via attributes[['SomeAttribute']]
			# does not return special characters (e.g. ä,ö,ü) correctly - though XML is in UTF-8!
			# one nice side effect: escape characters like '&lt;', '&amp;' are not converted to '<', '&', etc.
			key <- Get.Raw.Attribute(raw.string, Translation$Xml.Attribute.Key)
			
			if (!Translation$Xml.Attribute.Original.Text %in% attribute.names) 
				stop(c(filename, 
							 ': Missing attribute ', 
							 Translation$Xml.Attribute.Original.Text,
							 ' in:', 
							 toString.XMLNode(translationTag)))
			original.text <- Get.Raw.Attribute(raw.string, Translation$Xml.Attribute.Original.Text)
			
			# unusable attributes' xmlAttrs function
			# since special characters are not treated correctly
			#original.text <- attributes[['OriginalText']]
			
			# not needed since we don't use attributes' xmlAttrs function
			# attention: text may have escape characters and unfortunately
			# they are converted (probably by Java) e.g. from '&amp;' to '&'
			# therefore we have to convert it back to escape characters
			#original.text <- gsub(pattern='&', x=original.text, '&amp;', fixed=T)
			#original.text <- gsub(pattern='<', x=original.text, '&lt;', fixed=T)
			#original.text <- gsub(pattern='>', x=original.text, '&gt;', fixed=T)
			#original.text <- gsub(pattern="'", x=original.text, '&apos;', fixed=T)
			#original.text <- gsub(pattern='"', x=original.text, '&quot;', fixed=T)
			
			c(ID = id, Key = key, OriginalText = original.text)
			# no need to use sapply c(sapply(xmlChildren(translationTag), xmlValue), ID = id, Text = text)
		})
		
		# TODO: data frame elements are for some reason a list with a single value
		# here we need to transform into new data frame by
		# extracting each value and defining type (as.numeric, etc.)
		df <- transform(df, 
										ID = as.numeric(ID), 
										Key = as.character(Key), 
										OriginalText = as.character(OriginalText))
		
		# sort by first column ID
		df[order(df[,1]),]
	}
	
	# Creates a data frame with the columns ID and attribute.name
	Load.Language <- function(filename, attribute.name, column.name) {
		Load.Xml(filename, getAttributes <- function(translationTag){
			# get raw string of whole node
			raw.string <- toString.XMLNode(translationTag)

			# get attributes
			attributes <- xmlAttrs(translationTag)
			attribute.names <- names(attributes)
			
			if (!Translation$Xml.Attribute.Id %in% attribute.names) 
				stop(c(filename, 
							 ': Missing attribute ', 
							 Translation$Xml.Attribute.Id, 
							 ' in:', 
							 raw.string))
			# TODO @tt cannot convert ID to numeric (as.numeric(xmlGetAttr('ID')))
			# since in the code below when converted to a data frame it is a list - check later
			id <- as.numeric(Get.Raw.Attribute(raw.string, Translation$Xml.Attribute.Id))
			
			if (!attribute.name %in% attribute.names) 
				stop(c(filename, 
							 ': Missing attribute ', 
							 attribute.name, 
							 ' in:', 
							 raw.string))
			# workaround retrieving attribute via 'raw' toString function 
			# for some reason retrieving value via attributes[['SomeAttribute']]
			# does not return special characters (e.g. ä,ö,ü) correctly - though XML is in UTF-8!
			# one nice side effect: escape characters like '&lt;', '&amp;' are not converted to '<', '&', etc.
			text <- Get.Raw.Attribute(raw.string, attribute.name)
			
			# unusable attributes' xmlAttrs function - since special characters are not treated correctly
			#text <- attributes[['Text']]
			
			# not needed since we don't use attributes' xmlAttrs function
			# attention: text may have escape characters and unfortunately
			# they are converted (probably by Java) e.g. from '&amp;' to '&'
			# therefore we have to convert it back to escape characters
			#text <- gsub(pattern='&', x=text, '&amp;', fixed=T)
			#text <- gsub(pattern='<', x=text, '&lt;', fixed=T)
			#text <- gsub(pattern='>', x=text, '&gt;', fixed=T)
			#text <- gsub(pattern="'", x=text, '&apos;', fixed=T)
			#text <- gsub(pattern='"', x=text, '&quot;', fixed=T)
			
			result <- c(id, text)
			names(result) <- c('ID', column.name)
			result
			# no need to use sapply c(sapply(xmlChildren(translationTag), xmlValue), ID = id, Text = text)
		})
	}
	
	print(paste('Loading main file', main.file))
	main.df <- Load.Main(main.file)
	
	print(paste('Loading english file', english.file))
	english.df <- Load.Language(english.file, Translation$Xml.Attribute.Text, Translation$Xml.Attribute.English.Text)
	# TODO: data frame elements are for some reason a list with a single value
	# here we need to transform into new data frame by
	# extracting each value and defining type (as.numeric, etc.)
	english.df <- transform(english.df, 
													ID = as.numeric(ID), 
													EnglishText = as.character(EnglishText))
	
	# sort by first column ID
	english.df[order(english.df[,1]),]
	
	print(paste('Loading lang file', language.file))
	language.df <- Load.Language(language.file, Translation$Xml.Attribute.Text, Translation$Xml.Attribute.Text)
	# TODO: data frame elements are for some reason a list with a single value
	# here we need to transform into new data frame by
	# extracting each value and defining type (as.numeric, etc.)
	language.df <- transform(language.df, 
													 ID = as.numeric(ID), 
													 Text = as.character(Text))
	
	# sort by first column ID
	language.df[order(language.df[,1]),]
	
	print('Checking whether IDs in lang file exist in main file')
	
	# checking whether all IDs from lang file does also exist in main file
	print('>Result: merging main and english data frames')
	result <- merge(main.df, english.df, by = 'ID')
	print('>Result: merging result with language data frame')
	result <- merge(result, language.df, by = 'ID')

	print('Pass result to translation handler (summary creation)')
	# result handling
	if (!is.null(translation.handler)) {
		translation.handler(result, main.df, english.df, language.df)
	}
	
	result
}