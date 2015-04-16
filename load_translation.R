# problem loading rJava, http://stackoverflow.com/a/9120712
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk1.7.0_75\\jre')
require(XML)

# Creates a data frame with the columns ID, Key, and Text
LoadTranslation <- function(main.file,
														english.file,
														language.file,
														translation.handler = NULL # handler must accept these params:
														# result data frame: containing the merged df of both translation files
														# language data frame
														# main data frame
) {
	# Base class loading XML file with Translation tags
	LoadXml <- function(filename, getAttributes) {
		doc <- xmlParse(filename)
		root <- xmlRoot(doc)
		# read all Translation tags
		df <- as.data.frame(t(xpathSApply(root, "//*/Translation", getAttributes)), stringsAsFactors = FALSE)
	}
	
	# Creates a data frame with the columns ID, Key, and OriginalText
	LoadMain <- function(filename) {
		df <- LoadXml(filename, getAttributes <- function(translationTag){
			# get attributes
			attributes <- xmlAttrs(translationTag)
			attribute.names <- names(attributes)
			
			# get attributes
			if (!'ID' %in% attribute.names) stop(c(filename, ': Missing attribute ID in:', toString.XMLNode(translationTag)))
			id <- as.numeric(attributes[['ID']]) # TODO @tt cannot convert ID to numeric (as.numeric(xmlGetAttr('ID'))) since in the code below when converted to a data frame it is a list - check later
			
			if (!'Key' %in% attribute.names) stop(c(filename, ': Missing attribute Key in:', toString.XMLNode(translationTag)))
			key <- attributes[['Key']]
			
			if (!'OriginalText' %in% attribute.names) stop(c(filename, ': Missing attribute OriginalText in:', toString.XMLNode(translationTag)))
			# workaround retrieving attribute via 'raw' toString function 
			# for some reason retrieving value via attributes[['SomeAttribute']]
			# does not return special characters (e.g. ä,ö,ü) correctly - though XML is in UTF-8!
			# one nice side effect: escape characters like '&lt;', '&amp;' are not converted to '<', '&', etc.
			original.text <- toString.XMLNode(translationTag)
			original.text <- sub('OriginalText=\"', '', substring(original.text, regexpr('OriginalText=\"', original.text)))
			original.text <- substring(original.text, 1, regexpr('\"', original.text) - 1)
			
			# unusable attributes' xmlAttrs function - since special characters are not treated correctly
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
		df <- transform(df, ID = as.numeric(ID), Key = as.character(Key), OriginalText = as.character(OriginalText))
		
		# sort by first column ID
		df[order(df[,1]),]
	}
	
	# Creates a data frame with the columns ID and attribute.name
	LoadLanguage <- function(filename, attribute.name, column.name) {
		LoadXml(filename, getAttributes <- function(translationTag){
			# get attributes
			attributes <- xmlAttrs(translationTag)
			attribute.names <- names(attributes)
			
			if (!'ID' %in% attribute.names) stop(c(filename, ': Missing attribute ID in:', toString.XMLNode(translationTag)))
			id <- as.numeric(attributes[['ID']]) # TODO @tt cannot convert ID to numeric (as.numeric(xmlGetAttr('ID'))) since in the code below when converted to a data frame it is a list - check later
			
			if (!attribute.name %in% attribute.names) stop(c(filename, ': Missing attribute ', attribute.name, ' in:', toString.XMLNode(translationTag)))
			# workaround retrieving attribute via 'raw' toString function 
			# for some reason retrieving value via attributes[['SomeAttribute']]
			# does not return special characters (e.g. ä,ö,ü) correctly - though XML is in UTF-8!
			# one nice side effect: escape characters like '&lt;', '&amp;' are not converted to '<', '&', etc.
			text <- toString.XMLNode(translationTag)
			attribute.start <- paste(attribute.name, '=\"', sep='')
			text <- sub(attribute.start, '', substring(text, regexpr(attribute.start, text)))
			text <- substring(text, 1, regexpr('\"', text) - 1)
			
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
	main.df <- LoadMain(main.file)
	
	print(paste('Loading english file', english.file))
	english.df <- LoadLanguage(english.file, 'Text', 'EnglishText')
	# TODO: data frame elements are for some reason a list with a single value
	# here we need to transform into new data frame by
	# extracting each value and defining type (as.numeric, etc.)
	english.df <- transform(english.df, ID = as.numeric(ID), EnglishText = as.character(EnglishText))
	
	# sort by first column ID
	english.df[order(english.df[,1]),]
	
	print(paste('Loading lang file', language.file))
	language.df <- LoadLanguage(language.file, 'Text', 'Text')
	# TODO: data frame elements are for some reason a list with a single value
	# here we need to transform into new data frame by
	# extracting each value and defining type (as.numeric, etc.)
	language.df <- transform(language.df, ID = as.numeric(ID), Text = as.character(Text))
	
	# sort by first column ID
	language.df[order(language.df[,1]),]
	
	print('Checking whether IDs in lang file exist in main file')
	
	# checking whether all IDs from lang file does also exist in main file
	print('Result: merging main and english data frames')
	result <- merge(main.df, english.df, by = 'ID')
	print('Merging result and language data frames')
	result <- merge(result, language.df, by = 'ID')
	
	# result handling
	if (!is.null(translation.handler)) {
		translation.handler(result, language.df, main.df)
	}
	
	result
}