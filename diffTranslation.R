diffTranslation <- function(current, latest, colText = 'Text') {
        m <- merge(current, latest, by = 'ID', suffixes=c('.current', '.latest'))
        
        result <- data.frame(
                ID=numeric(),
                Text.current=character(),
                Text.latest=character())
        count <- 1
        for (i in 1:nrow(m)) {
                current <- m[[i, paste(colText, '.current', sep ='')]]
                latest <- m[[i, paste(colText, '.latest', sep ='')]]
                #print(class(current))
                if (current != latest) {
                        #print(m[i,])
                        result <- rbind(result, m[i,])
                        #result[count,] <- m[i,]
                        #count <- count + 1
                }
        }
        result
}