# Excel to Csv Converter

 ConvertExcelToCsv.jar, converts the Excel Spreadsheet into CSV files. Both .xls and .xlsx file formats are supported.
 Each Tab in the spreadsheet will be converted to a new csv file with name, same as the Sheet name. 
 1st arg is Excel file name. 2nd arg is Sheet name to be translated to csv file. 
 If not specified, it converts all sheets to separate csv files.
  
 # Usage:
  
   1) Clone/Download the repository into your local environment. Copy your source Excel file in the 'target' folder where ConvertExcelToCsv.jar is residing.
   
   2) Run below command, to convert all Tabs of 'sample-xlsx-file.xlsx' spreadsheet to separate csv files. Csv Files are named after the Tab/Sheet names.
  
   > java -jar ConvertExcelToCsv.jar ./sample-xlsx-file.xlsx
     
  3) To convert the only 1 Sheet (with name 'Sample_Sheet2') on 'sample-xlsx-file.xlsx' spreadsheet to Sample_Sheet2.csv file,
  
   > java -jar ConvertExcelToCsv.jar ./sample-xlsx-file.xlsx Sample_Sheet2 
    
  4) Output csv files will be generated in 'target' folder.
    
 # Usage: (For local testing)
 
 1) Clone this git repo into your eclipse general project
 2) Convert the existing Java Project to Maven Project
 3) Copy your input Excel file to root project folder
 3) Build and Run Excel2Csv.java
 4) Output csv files should be generated in project root folder
 
 To package jars
 1) Select pom.xml, Run as 'Maven install'
 
 
 
 