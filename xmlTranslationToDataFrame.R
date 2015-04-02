# problem loading rJava, http://stackoverflow.com/a/9120712
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jdk1.7.0_75\\jre')
# source: http://stackoverflow.com/a/16805244
require(XML)
xmlTranslationToDataFrame <- function(xdata){
        dumFun <- function(x){
                # get attributes
                xattrs <- xmlAttrs(x)
                #doesn't make sense to convert to numeric since below
                # when converted to a data frame it is a list, don't ask me why...
                #c(sapply(xmlChildren(x), xmlValue), ID = as.numeric(xattrs[['ID']]), Text = as.character(xattrs[['Text']]))
                
                #convert back to escape characters
                text <- gsub(pattern='&', x=xattrs[['Text']], '&amp;', fixed=T)
                text <- gsub(pattern='<', x=text, '&lt;', fixed=T)
                text <- gsub(pattern='>', x=text, '&gt;', fixed=T)
                text <- gsub(pattern="'", x=text, '&apos;', fixed=T)
                text <- gsub(pattern='"', x=text, '&quot;', fixed=T)
                
                c(sapply(xmlChildren(x), xmlValue), ID = xattrs[['ID']], Text = text)
        }
        doc <- xmlParse(xdata, getDTD = F)
        root <- xmlRoot(doc)
        
        df <- as.data.frame(t(xpathSApply(root, "//*/Translation", dumFun)), stringsAsFactors = FALSE)

        # transform
        # attributes are somewhoe a list with a single element
        # make ID numeric, Text as character
        df <- transform(df, ID = as.numeric(ID), Text = as.character(Text))

        # sort by first column ID
        df[order(df[,1]),]
}