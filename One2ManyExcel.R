# This code reads in an excel or csv file, filters it based on criteria and
# generates multiple excel files each with one of the values in the criteria

# Packages to be installed include dplyr and xlsx

install.packages('dplyr')
library(dplyr)
install.packages('xlsx')
library(xlsx)
# Read in the file - excel or csv
#df=read.xlsx('Example_Excel_File.xlsx',sheetIndex = 1,startRow = 2, header=T)
df2<-read.csv('Example_Csv_File.csv',na.strings = "abcdef", header=T)

# Create a variable to capture unique values of the Center_Gender column
# For other columns, replace Center_Gender with the column_name
SC<-unique(df2$Center_Gender)
SClist<-list()

# for plain writing to excel file, without excel formatting
# 
# for (i in 1:length(SC)) {
  # SClist[[i]]<-dplyr::filter(df2, Center_Gender==SC[i])
  # write.xlsx(SClist[[i]], file = paste(SC[i], ".xlsx"))
# }

# Using the dplyr filter function to create dataframes 
# each pertaining to a unique value of Center_Gender
# 
# for creating excel formats by creating files and adding to it
# update ColIndex to how many ever columns you need formatting for

for (i in 1:length(SC)) {
  SClist[[i]]<-dplyr::filter(df2, Center_Gender==SC[i])
  wb1<-createWorkbook()
  sheet1<-createSheet(wb1)
  ColStyle<- CellStyle(wb1) + Font(wb1, heightInPoints = 14, isBold = T)
    addDataFrame(SClist[[i]], sheet1, colnamesStyle = ColStyle)
  autoSizeColumn(sheet=sheet1, colIndex = c(1:9))
  saveWorkbook(wb1, file = paste(SC[i], ".xlsx"))
} 
