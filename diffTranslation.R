diffTranslation <- function(current, latest, colText = 'Text') {
        # merge current and latest data frame by id and
        # add suffix for all redundant columns in both data frames with '.current' and '.latest'
        m <- merge(current, latest, by = 'ID', suffixes=c('', '.latest'))
        
        result <- data.frame()
        count <- 1
        for (i in 1:nrow(m)) {
                # colText != colText + '.latest' ???
                # if different then add to result
                if (m[[i, colText]] != m[[i, paste(colText, '.latest', sep ='')]]) {
                        #print(m[i,])
                        result <- rbind(result, m[i,])
                        #result[count,] <- m[i,]
                        #count <- count + 1
                }
        }
        result
}