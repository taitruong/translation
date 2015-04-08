# problem loading rJava, http://stackoverflow.com/a/9120712
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk1.7.0_75\\jre')
require(XML)

# Creates a data frame with the columns ID, Key, and Text
loadTranslation <- function(langFile, mainFile) {
        # Base class loading XML file with Translation tags
        loadXml <- function(filename, getAttributes) {
                doc <- xmlParse(filename)
                root <- xmlRoot(doc)
                # read all Translation tags
                df <- as.data.frame(t(xpathSApply(root, "//*/Translation", getAttributes)), stringsAsFactors = FALSE)
                
        }
        
        # Creates a data frame with the columns ID and Key
        loadMain <- function(filename) {
                df <- loadXml(filename, getAttributes <- function(translationTag){
                        # get attributes
                        attributes <- xmlAttrs(translationTag)
                        attrNames <- names(attributes)

                        # get attributes
                        if (!'ID' %in% attrNames) stop(c(filename, ': Missing attribute ID in:', toString.XMLNode(translationTag)))
                        id <- as.numeric(attributes[['ID']]) # TODO @tt cannot convert ID to numeric (as.numeric(xmlGetAttr('ID'))) since in the code below when converted to a data frame it is a list - check later
                        
                        if (!'Key' %in% attrNames) stop(c(filename, ': Missing attribute Key in:', toString.XMLNode(translationTag)))
                        key <- attributes[['Key']]
                        
                        if (!'OriginalText' %in% attrNames) stop(c(filename, ': Missing attribute OriginalText in:', toString.XMLNode(translationTag)))
                        # workaround retrieving attribute via 'raw' toString function 
                        # for some reason retrieving value via attributes[['SomeAttribute']]
                        # does not return special characters (e.g. ä,ö,ü) correctly - though XML is in UTF-8!
                        # one nice side effect: escape characters like '&lt;', '&amp;' are not converted to '<', '&', etc.
                        originalText <- toString.XMLNode(translationTag)
                        originalText <- sub('OriginalText=\"', '', substring(originalText, regexpr('OriginalText=\"', originalText)))
                        originalText <- substring(originalText, 1, regexpr('\"', originalText) - 1)
                        
                        # unusable attributes' xmlAttrs function - since special characters are not treated correctly
                        #originalText <- attributes[['OriginalText']]

                        # not needed since we don't use attributes' xmlAttrs function
                        # attention: text may have escape characters and unfortunately
                        # they are converted (probably by Java) e.g. from '&amp;' to '&'
                        # therefore we have to convert it back to escape characters
                        #originalText <- gsub(pattern='&', x=originalText, '&amp;', fixed=T)
                        #originalText <- gsub(pattern='<', x=originalText, '&lt;', fixed=T)
                        #originalText <- gsub(pattern='>', x=originalText, '&gt;', fixed=T)
                        #originalText <- gsub(pattern="'", x=originalText, '&apos;', fixed=T)
                        #originalText <- gsub(pattern='"', x=originalText, '&quot;', fixed=T)
                        
                        c(ID = id, Key = key, OriginalText = originalText)
                        # no need to use sapply c(sapply(xmlChildren(translationTag), xmlValue), ID = id, Text = text)
                })
                
                # TODO: data frame elements are for some reason a list with a single value, so here we need to transform
                # transform into new data frame by extracting value from singleton list and defining type (as.numeric, etc.)
                # make ID numeric, Text as character
                df <- transform(df, ID = as.numeric(ID), Key = as.character(Key), OriginalText = as.character(OriginalText))
                
                # sort by first column ID
                df[order(df[,1]),]
        }

        print(paste('Loading lang file', langFile))
        langDf <- loadXml(langFile, getAttributes <- function(translationTag){
                # get attributes
                attributes <- xmlAttrs(translationTag)
                attrNames <- names(attributes)

                if (!'ID' %in% attrNames) stop(c(langFile, ': Missing attribute ID in:', toString.XMLNode(translationTag)))
                id <- as.numeric(attributes[['ID']]) # TODO @tt cannot convert ID to numeric (as.numeric(xmlGetAttr('ID'))) since in the code below when converted to a data frame it is a list - check later
                
                if (!'Text' %in% attrNames) stop(c(langFile, ': Missing attribute Text in:', toString.XMLNode(translationTag)))
                # workaround retrieving attribute via 'raw' toString function 
                # for some reason retrieving value via attributes[['SomeAttribute']]
                # does not return special characters (e.g. ä,ö,ü) correctly - though XML is in UTF-8!
                # one nice side effect: escape characters like '&lt;', '&amp;' are not converted to '<', '&', etc.
                text <- toString.XMLNode(translationTag)
                text <- sub('Text=\"', '', substring(text, regexpr('Text=\"', text)))
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
                
                c(ID = id, Text = text)
                # no need to use sapply c(sapply(xmlChildren(translationTag), xmlValue), ID = id, Text = text)
        })
        
        # TODO: data frame elements are for some reason a list with a single value, so here we need to transform
        # transform into new data frame by extracting value from singleton list and defining type (as.numeric, etc.)
        # make ID numeric, Text as character
        langDf <- transform(langDf, ID = as.numeric(ID), Text = as.character(Text))
        
        # sort by first column ID
        langDf[order(langDf[,1]),]

        print(paste('Loading main file', mainFile))
        mainDf <- loadMain(mainFile)
        
        print('Checking whether IDs in lang file exist in main file')
        
        # checking whether all IDs from lang file does also exist in main file
        langDfNotInMainDf <- langDf[!langDf$ID %in% mainDf$ID,]
        if (nrow(langDfNotInMainDf) > 0) {
                stop(c('The following IDs from the lang file are NOT in the main file: ', toString(langDfNotInMainDf$ID)))
        }
        
        # makre sure both data frame has same row size
        if (nrow(langDf) != nrow(mainDf)) {
                stop(c('Lang file size (', nrow(langDf), ') is not equal main file size (', nrow(mainDf), ')'))
        }
        
        print('Merging main and lang data frames...')
        merge(mainDf, langDf)
}