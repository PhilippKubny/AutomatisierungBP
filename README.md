Ignore -schlagwortOptionen (don't need)

excel.py has everything to do with excel. Extracting information and writing back information. 
PDFScanner.py is for scanning downloaded AD files (only supports AD files so far)
PDFdump.py is only for debugging and figuring out the structure of a PDF to then implement a scanner in PDFScanner
PlayHandelsregister.py is the main that automates everything. First, it's inputting extracted information into the search form, then it performs the search, downloads the document, scans the document and lastly writes back to the excel file. 

2 Options to run the PlayHandelsregister
1. batch running with Excel table. 
    i.e python PlayHandelsregister.py -d --download-ad --excel "C:\Users\User\Downloads\TestBP.xlsx" --sheet "Tabelle1" --start 25 --end 30 -postal --outdir "C:\Users\User\Downloads\BP"
   -d is for debugging (more information)
   --download-ad is for downloading the AD file (if not mentioned, it will only perform the search without downloading the document)
   --excel "<Excelpath>" tells the Excel path (required in batch running mode (it's the keyword to run the batch function in the first place))
   --sheet "<Table>" is to specify the table you want to use in the excel file
   --start <int> is to specify the start row to beginn performing the automation (it correlates 1 to 1 with the excel rows, that means, i.e 3 will start the automation from row 3 in the excel file, default is 3 since the first 2 rows are headers in the excel file)
   --end <int> same as --start (start and end are included -> start 3 end 5 are 3 rows of scanning), this is required in the batch running mode
   -postal is optional if wanting to run the search with the postal code included (potentially higher hit rate if the register number is not given in the excel file)
   -outdir "" specifies the directory where all the AD PDF are downloaded to (if not given it will automatically create a folder in Downloads named "BP"


2. run in singel shot mode, this is complementory to the batch mode for singel entries
   i.e python PlayHandelsregister.py -s "THYSSENKRUPP SCHULTE GMBH"  --register-number "26718" -d --download-ad -postal --outdir "C:\Users\User\Downloads\BP" -sap "2203241" -row "352"
   in single-shot mode (occurs automatically if -excel is not given) -s, -sap, -row are mandatory
   -s stands for schlagwoerter and correlates to the company name
   --regist-number is optional if you know the register number in the handelsregisterbuch
   -sap is the number we give internally in MTU can be read of the excel sheet
   -row correlates to the companies row in the excel sheet (so that the information can be updated in the excel sheet)
