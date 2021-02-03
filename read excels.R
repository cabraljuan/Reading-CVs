
###########################################################################
###########################################################################
###                                                                     ###
###                      SEARCHING DATA IN EXCELS                       ###
###                                                                     ###
###########################################################################
###########################################################################



# Libraries
library(data.table) # data.table
library(writexl) # Write excel
library(readxl) # read excel
library(plyr) # manage dataframes

# Path
path<-"J:/Trabajo jaime/excels"
setwd(path)

# All folders inside path
subdir<-c(list.dirs(path = path, full.names = TRUE, recursive = TRUE))




#################################################################
##                     Searching for data in excels            ##
#################################################################



results <- data.frame()


for (folder in subdir){
  # Search in all folders
  setwd(folder)
  
  # Excels
  formats<-c("xlsm$","xlsx$","xls$")
  files<-c()
  for (patterns in formats){
    files <- c(files,list.files(pattern=patterns))
    
  }
  
  if (length(files)==0){
    next
  } # If there aren't any excels, skip
  
  
  for (excel in files){ # Looking at each excel in a folder
    print(paste0("Loading: ",folder,"/",excel))
    # Search for each excel
    
    # If sheets doesn't exist then try another name, using trycatch to capture any error associated with a non existent sheet
    error <- tryCatch(read_excel(excel,sheet="CD"), error = function(e) e)
    if (!( class(error)[2] =="error") ){ # IF using CD doesn't give errors then use CD
      excelfile<-read_excel(excel,sheet="CD")
    } else{ # IF using CD gives error, then try "COSTO DIRECTO"
      error <- tryCatch(read_excel(excel,sheet="COSTO DIRECTO"), error = function(e) e)
      if (!( class(error)[2] =="error") ){ # If using "COSTO DIRECTO" doesn't give errors then use COSTO DIRECTO
        excelfile<-read_excel(excel,sheet="COSTO DIRECTO")
        
      } else{next} # If "COSTO DIRECTO" gives error then skip this file
    }
   
    # Search for specific data
    excelfile$'Datos Importantes :'<-ifelse(is.na(excelfile$'Datos Importantes :'),"empty",excelfile$'Datos Importantes :') # Change NA cells to "empty" cells
    colnumber<-which( colnames(excelfile)=='Datos Importantes :' )
    wage<-excelfile[excelfile$'Datos Importantes :'=="SUELDO LIQUIDO INICIAL SEGÚN OFERTA",colnumber+2]
    # If wage>0 then save category
    
    if (wage>0){
      # Number of row 
      rownumber<-which(grepl("SUELDO LIQUIDO INICIAL SEGÚN OFERTA", excelfile$'Datos Importantes :'))
      # Extract data from row above
      category<-excelfile[rownumber-1,]$`43416`
    }
    results<-rbind(results,data.frame(wage=wage,category=category,folder=folder,excel=excel))
    
   excel<-NA
   
  }
  wage<-NA
  category<-NA
}



# Names for dataset
names <- c("Wage","Category","Folder","File name")
colnames(results) <- names


# Save results
write_xlsx(results,paste0(path,"/","Results.xlsx") )




