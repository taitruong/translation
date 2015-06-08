# problem loading rJava, http://stackoverflow.com/a/9120712
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk1.7.0_75\\jre')
require(XML)

# Creates a data frame with the columns ID, Key, and Text
LoadTranslation <- function(main1.file,
														main2.file,
														language.file,
														language.file.suffix,
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
	
	# Creates a data frame with the columns ID and attribute.name
	Load.Language.As.Data.Frame <- function(filename, attribute.name, column.name, consider.key = FALSE) {
		Load.Xml(filename, getAttributes <- function(translationTag){
			# get attributes
			attributes <- xmlAttrs(translationTag)
			attribute.names <- names(attributes)
			
			# get raw string of whole node
			#raw.string <- toString.XMLNode(translationTag) # does not work for special characters like kyrilian, japanese, chinese letters...
			raw.string <- as(translationTag, 'character')
			
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
			
			if (consider.key && !Translation$Xml.Attribute.Key %in% attribute.names) 
				stop(c(filename, 
							 ': Missing attribute Key in:', 
							 raw.string))
			# workaround retrieving attribute via 'raw' toString function 
			# for some reason retrieving value via attributes[['SomeAttribute']]
			# does not return special characters (e.g. ä,ö,ü) correctly - though XML is in UTF-8!
			# one nice side effect: escape characters like '&lt;', '&amp;' are not converted to '<', '&', etc.
			key <- if (consider.key) {
				Get.Raw.Attribute(raw.string, Translation$Xml.Attribute.Key)
			} else {
				'UNUSED' # won't be used below
			}
			
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
			# since special characters are not treated correctly
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
			
			if (consider.key) {
				result <- c(id, key, text)
				names(result) <- c('ID', 'Key', column.name)
			} else {
				result <- c(id, text)
				names(result) <- c('ID', column.name)
			}
			result
			# no need to use sapply c(sapply(xmlChildren(translationTag), xmlValue), ID = id, Text = text)
		})
	}
	
	print(paste('Loading main file', main1.file))
	main1.df <- Load.Language.As.Data.Frame(main1.file, Translation$Xml.Attribute.Original.Text, Translation$Df.Attribute.Main1, T)
	# TODO: data frame elements are for some reason a list with a single value
	# here we need to transform into new data frame by
	# extracting each value and defining type (as.numeric, etc.)
	## first rename column name since transform()-call requires a tag name we cannot pass a variable like SomeText = as.character(...)
	names(main1.df) <- c(Translation$Xml.Attribute.Id, Translation$Xml.Attribute.Key, 'Text')
	## transform
	main1.df <- transform(main1.df, 
									ID = as.numeric(ID), 
									Key = as.character(Key), 
									Text = as.character(Text))
	## re-rename columns
	names(main1.df) <- c(Translation$Xml.Attribute.Id, Translation$Xml.Attribute.Key, Translation$Df.Attribute.Main1)
	
	# sort by first column ID
	main1.df[order(main1.df[,1]),]
	
	print(paste('Loading english file', main2.file))
	main2.df <- Load.Language.As.Data.Frame(main2.file, Translation$Xml.Attribute.Text, Translation$Df.Attribute.Main2, F)
	# TODO: data frame elements are for some reason a list with a single value
	# here we need to transform into new data frame by
	# extracting each value and defining type (as.numeric, etc.)
	## first rename column name since transform()-call requires a tag name we cannot pass a variable like SomeText = as.character(...)
	names(main2.df) <- c(Translation$Xml.Attribute.Id, 'Text')
	## transform
	main2.df <- transform(main2.df, 
													ID = as.numeric(ID), 
													Text = as.character(Text))
	## re-rename columns
	names(main2.df) <- c(Translation$Xml.Attribute.Id, Translation$Df.Attribute.Main2)
	
	# sort by first column ID
	main2.df[order(main2.df[,1]),]
	
	print(paste('Loading lang file', language.file))
	language.df <- Load.Language.As.Data.Frame(language.file, Translation$Xml.Attribute.Text, language.file.suffix, F)
	# TODO: data frame elements are for some reason a list with a single value
	# here we need to transform into new data frame by
	# extracting each value and defining type (as.numeric, etc.)
	## first rename column name since transform()-call requires a tag name we cannot pass a variable like SomeText = as.character(...)
	names(language.df) <- c(Translation$Xml.Attribute.Id, 'Text')
	## transform
	language.df <- transform(language.df, 
													 ID = as.numeric(ID), 
													 Text = as.character(Text))
	## re-rename columns
	names(language.df) <- c(Translation$Xml.Attribute.Id, language.file.suffix)
	
	# sort by first column ID
	language.df[order(language.df[,1]),]
	
	print('Checking whether IDs in lang file exist in main file')
	
	# checking whether all IDs from lang file does also exist in main file
	print('>Result: merging main and english data frames')
	result <- merge(main1.df, main2.df, by = 'ID')
	print('>Result: merging result with language data frames')
	result <- merge(result, language.df, by = 'ID')

	print('Pass result to translation handler (summary creation)')
	# result handling
	if (!is.null(translation.handler)) {
		translation.handler(result, main1.df, main2.df, language.df)
	}
	
	result
}