# problem loading rJava, http://stackoverflow.com/a/9120712
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk1.7.0_75\\jre')
require(XML)

# Creates a data frame with the columns ID, Key, and Text
loadTranslation <- function(langFile, mainFile) {
        # Base class loading XML file with Translation tags
        loadXml <- function(filename, getAttributes) {
                doc <- xmlParse(filename, getDTD = F)
                root <- xmlRoot(doc)
                # read all Translation tags
                df <- as.data.frame(t(xpathSApply(root, "//*/Translation", getAttributes)), stringsAsFactors = FALSE)
                
        }
        
        # Creates a data frame with the columns ID and Key
        loadMain <- function(filename) {
                df <- loadXml(filename, getAttributes <- function(translationTag){
                        # get attributes
                        id <- xmlGetAttr(translationTag, 'ID') # TODO @tt cannot convert ID to numeric (as.numeric(xmlGetAttr('ID'))) since in the code below when converted to a data frame it is a list - check later
                        if (is.null(id)) stop(c(filename, ': Missing attribute ID in:', toString.XMLNode(translationTag)))
                        
                        key <- xmlGetAttr(translationTag, 'Key')
                        if (is.null(key)) stop(c(filename, ': Missing attribute Key in:', toString.XMLNode(translationTag)))
                        
                        originalText <- xmlGetAttr(translationTag, 'OriginalText')
                        if (is.null(originalText)) stop(c(filename, ': Missing attribute OriginalText in:', toString.XMLNode(translationTag)))
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

        print('loading lang file...')
        langDf <- loadXml(langFile, getAttributes <- function(translationTag){
                # get attributes
                attributes <- xmlAttrs(translationTag)
                attrNames <- names(attributes)

                if (!'ID' %in% attrNames) stop(c(langFile, ': Missing attribute ID in:', toString.XMLNode(translationTag)))
                id <- as.numeric(attributes[['ID']]) # TODO @tt cannot convert ID to numeric (as.numeric(xmlGetAttr('ID'))) since in the code below when converted to a data frame it is a list - check later
                
                if (!'Text' %in% attrNames) stop(c(langFile, ': Missing attribute Text in:', toString.XMLNode(translationTag)))
                text <- attributes[['Text']]

                # attention: text may have escape characters and unfortunately
                # they are converted (probably by Java) e.g. from '&amp;' to '&'
                # therefore we have to convert it back to escape characters
                text <- gsub(pattern='&', x=text, '&amp;', fixed=T)
                text <- gsub(pattern='<', x=text, '&lt;', fixed=T)
                text <- gsub(pattern='>', x=text, '&gt;', fixed=T)
                text <- gsub(pattern="'", x=text, '&apos;', fixed=T)
                text <- gsub(pattern='"', x=text, '&quot;', fixed=T)
                
                c(ID = id, Text = text)
                # no need to use sapply c(sapply(xmlChildren(translationTag), xmlValue), ID = id, Text = text)
        })
        
        # TODO: data frame elements are for some reason a list with a single value, so here we need to transform
        # transform into new data frame by extracting value from singleton list and defining type (as.numeric, etc.)
        # make ID numeric, Text as character
        langDf <- transform(langDf, ID = as.numeric(ID), Text = as.character(Text))
        
        # sort by first column ID
        langDf[order(langDf[,1]),]

        print('loading main file...')
        mainDf <- loadMain(mainFile)
        
        print('checking whether IDs in lang file exist in main file')
        
        # checking whether all IDs from lang file does also exist in main file
        langDfNotInMainDf <- langDf[!langDf$ID %in% mainDf$ID,]
        if (nrow(langDfNotInMainDf) > 0) {
                stop(c('The following IDs from the lang file are NOT in the main file: ', toString(langDfNotInMainDf$ID)))
        }
        
        # makre sure both data frame has same row size
        if (nrow(langDf) != nrow(mainDf)) {
                stop(c('Lang file size (', nrow(langDf), ') is not equal main file size (', nrow(mainDf), ')'))
        }
        
        print('merging main and lang data frames...')
        merge(mainDf, langDf)
}