
###########################################################################
###########################################################################
###                                                                     ###
###                   SEARCHING WORDS AND MORE IN CVS                   ###
###                                                                     ###
###########################################################################
###########################################################################



# Libraries
library(pdftools) # read pdf
library(stringr) # Search for words
library(stringi) # remove accents
library(data.table) # data.table
library(writexl) # Write excel


# Path
path<-"J:/Trabajo jaime/CVs"
setwd(path)

# All folders inside path
subdir<-c(list.dirs(path = path, full.names = TRUE, recursive = TRUE))


# Word to be searched
word<-"ingeniero"



#################################################################
##                     Searching for words                     ##
#################################################################


results <- data.frame()
for (folder in subdir){
  # Search in all folders
  setwd(folder)
  
  # Pdfs
  files <- list.files(pattern = "pdf$")
  
  if (length(files)==0){
    next
  } # If there aren't any pdfs, skip
  
  # Pdf text
  text <- lapply(files, pdf_text)
  
  for (n in 1:length(text)){
    # Search for each CV
    cv<-text[[n]]
    # Search for each page
    for (page in 1:length(cv)){
      # Lower case
      page_text<-tolower(cv[page])
      word<-tolower(word)
      # remove accents
      page_text<-iconv(page_text, from = 'UTF-8', to = 'ASCII//TRANSLIT')
      # Check if word is in page
      if(str_detect(page_text, tolower(word), negate = FALSE)==TRUE){
        #print(paste0("El documento ",files[n]," contiene la palabra ",word," en la página ",page))
        # Save in dataset
        # Extract words before and after word of interest
        wordsbefore<-wordsafter<-10
        sentence<- str_extract(page_text, paste0("(( \\S+){",wordsbefore,"} ",word,"[[:punct:]\\s]*( \\S+){",wordsafter,"})"))
        while (is.na(sentence)){
          wordsbefore<-wordsbefore-1
          wordsafter<-wordsafter-1
          if (wordsafter==0|wordsbefore==0){
            sentence<-"ERROR"
            break
          }
          sentence<- str_extract(page_text, paste0("(( \\S+){",wordsbefore,"} ",word,"[[:punct:]\\s]*( \\S+){",wordsafter,"})"))
        }
        results<-rbind(results,data.frame(paste0(files[n]),paste0(page),word,sentence,n,folder) )
        
      }
      # If doesn't find then next
      else{
        next
      }
    }
    
  }
}
  
  
 

# Names for dataset
names <- c("Document","Page","Word","Sentence","PDF number","Folder")
colnames(results) <- names


#################################################################
##                Adding experience in dataset                 ##
#################################################################


results$experience<-NA

pdfs<-unique(results$`PDF number`)
for (n in pdfs){
  # Search for each CV
    cv<-text[[n]]
    # Search for each page
    for (page in 1:length(cv)){
      # Lower case
      page_text<-tolower(cv[page])
      word<-tolower(word)
      # remove accents
      page_text<-iconv(page_text, from = 'UTF-8', to = 'ASCII//TRANSLIT')
      word<-iconv(word, from = 'UTF-8', to = 'ASCII//TRANSLIT')
      # Check if word is in page
        if(str_detect(page_text, tolower(word), negate = FALSE)==TRUE){
          experience<- str_extract(page_text, paste0("(( \\S+){",1,"} ","anos","[[:punct:]\\s]*( \\S+){",0,"})"))
          results[results$Document==paste0(files[n]),]$experience<-experience
        } else{
          next
            }
    }
}
        
        
# Save results
write_xlsx(results,"Results.xlsx")












##################################################################
##                           OLD CODE                           ##
##################################################################








install.packages("iconv")
library(iconv)
page_text<-iconv(page_text, from = 'UTF-8', to = 'ASCII//TRANSLIT')



base <- data.table(terme = c("Millésime", 
                             "boulangère", 
                             "üéâäàåçêëèïîì"))

base[, terme := stri_trans_general(str = terme, 
                                   id = "Latin-ASCII")]



# sentence<-str_extract(page_text,paste0("([^\\s]+\\s+){",wordsbefore,"}",word,"(\\s+[^\\s]+){",wordsafter,"}"))
for (m in 1:length(max(wordsbefore,wordsafter))){
  sentence<- str_extract(page_text, paste0("(( \\S+){",wordsbefore,"} ",word,"[[:punct:]\\s]*( \\S+){",wordsafter,"})"))
  
  if (is.na(sentence) ){
    
  }else{next
  }
}

count<-1
for (cv in text){
  cv<-tolower(cv)
  if(str_detect(cv, word, negate = FALSE)==TRUE){
    print(paste0("El documento ",files[count]," contiene la palabra ",word," en la página ",count2))
  }else{
    next
  }
  
}


for (cv in 1:text){
  count2<-1
  for (page in cv){
    if(str_detect(tolower(page), word, negate = FALSE)==TRUE){
      print(paste0("El documento ",files[count]," contiene la palabra ",word," en la página ",count2))
    }else{
      next
    }
    count2<- count2+count2
  }
  count<-count+count
  
}
cv
