rm(list=ls())

setwd("C:/Users/ThunderDom/Documents/R/R code")

#load libraries
library(stringr)
library(tm)
library(wordcloud)
library(slam)
library(sentimentr)

#Load comments/text
library(xlsx)
post = read.xlsx("Deals Chats.xls", sheetIndex=1)


#renaming the variable as comments
colnames(post$NA..4)="comments"
str(post)
post$NA..4=as.character(post$NA..4)

#extracting the variable 7 and converting it into a data frame
data=data.frame(post[,7])
str(data)                
names(data)="comments"
data$comments=as.character(data$comments)

#to delete the missing values
data=data[complete.cases(data),]

#convert the lines containing * into NA
data=gsub("[*]",NA,data)

#convert the lines containing - into NA
data=gsub("[-]",NA,data)


#converting the vector into the character vector
data=as.character(data)


#converting the vector into the data frame
data=data.frame(c(data))

data[data==""]<-NA
data<-data[complete.cases(data),]
data=as.character(data)
data=data.frame(c(data))

names(data)="comments"

postCorpus = Corpus(VectorSource(data$comments))

#applyong the tolower function to corpus
postCorpus=tm_map(postCorpus,tolower)

writeLines(as.character(postCorpus[[14]]))

#applying the remove punctation function to corpus
postCorpus=tm_map(postCorpus,removePunctuation)

#applying the remove numbers function to corpus
postCorpus=tm_map(postCorpus,removeNumbers)

#stripping the white spaces 
postCorpus=tm_map(postCorpus,stripWhitespace)

#stemming the corpus
postCorpus=tm_map(postCorpus,stemDocument)

#converting into white document
postCorpus=tm_map(postCorpus,PlainTextDocument)

#recreate corpus
postCorpus = Corpus(VectorSource(postCorpus))

#applying the stop words with custom stopwords 
postCorpus = tm_map(postCorpus, removeWords, c('i','its','it','us','use','used','using','will','yes','say','can','take','one',
                                               stopwords('english')))

#converting into term document matrix
tdm=TermDocumentMatrix(postCorpus)

inspect(tdm)

#extracting the patterns
words_freq = rollup(tdm, 2, na.rm=TRUE, FUN = sum)
words_freq = as.matrix(words_freq)
words_freq = data.frame(words_freq)
words_freq$words = row.names(words_freq)
row.names(words_freq) = NULL
words_freq = words_freq[,c(2,1)]
names(words_freq) = c("Words", "Frequency")
wordcloud(words_freq$Words, words_freq$Frequency,scale=c(1, .8), colors=brewer.pal(10, "Dark2"))

#Most frequent terms which appears in atleast 300 times
findFreqTerms(tdm, 300)

#creating the word cloud 
graphics.off()
wordcloud(postCorpus, max.words = 500, scale=c(6, .5), colors=brewer.pal(300, "Dark2"))  
 
#applying the sentiment analysis
library(RSentiment)
data$comments=as.character(data$comments)
data = calculate_sentiment(data$comments) 

save(data,file="data.rda")
save(post,file="post.rda")
save(postCorpus,file="postCorpus.rda")
save(words_freq,file="word frequency.rda")
save(tdm,file="tdm.rda")
