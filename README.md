Idee: 
	We imitate a human who's doing everything by hand. There is no API for handelsregister.de at this moment.
 	In the following you can see the startpage of the website https://www.handelsregister.de/rp_web/welcome.xhtml. On there we click on the 'Advanced Search option' on the right side. 
	<img width="1552" height="815" alt="image" src="https://github.com/user-attachments/assets/bf638fee-9e7c-4694-a3c7-42b1c53af508" />
 	
  After clicking we'll get the search form:
  	<img width="1888" height="879" alt="image" src="https://github.com/user-attachments/assets/187c6a1d-77c1-4d06-90b7-793d9c721757" /><img width="1368" height="870" alt="image" src="https://github.com/user-attachments/assets/b223b4f7-c043-48a7-97fe-c69798511366" />
	Depending on how you run the python script (more explained below, i.e -excel, -registernumber, -postal etc.) the respective fields will be filled. Mandatory to fill is the "Company or keywords" field. 
 Other than that, the script also allows to fill the "register number" and "postal code" field. In theory you could also change the "Search for records that" field by modifying the -mode parameter (I wouldn't recommend using it though, because the functions don't seem to be implemented very well on the website)

Lets try running a search (i.e Schwille Elektronik Produktions- und Vertriebs GmbH). After clicking on the "Find" button a list of Matching results will show up:
<img width="1891" height="925" alt="image" src="https://github.com/user-attachments/assets/0fa0d58d-37c5-4bf5-a477-fa0cf77f4233" />


 If the company name combined with the optional parameters is unique we will only receive one result (will always be unique if we have the register number of the company). The python programm will only handle unique results.
 If there are multiple results it will be logged in the column "name1-4" (column 'T') how many results where found. 0 means there was no matching comany (either the company name was misspelled or is not registered in the handelsregister).

If the result is unique the script will download the AD file and save it to the BP folder in Downloads with the following format <internalSAPNumber_ComapanyName_DateOfDownload>:
<img width="516" height="322" alt="image" src="https://github.com/user-attachments/assets/dffa6cf2-ce8c-4564-a8f0-d6190b53115d" />

There are 2 types of AD files in the Handelsregister:
<img width="1188" height="901" alt="image" src="https://github.com/user-attachments/assets/e90d0952-e1b1-48ff-8794-c20756a89cda" /><img width="1128" height="895" alt="image" src="https://github.com/user-attachments/assets/0cc079eb-5b0c-4590-8d69-6dbf310e7a3d" />

Depending on which format was used the script will perform a PDF scan to extract all the relevant information (Company name, register number, address). 
If the AD file had an unexpected format it will show error in the excel column "name1-4"/'T' "Unexpected format"

If everything worked it will write back all the information that was exctracted back to the excel file and continune with the next row. 









































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
   


3. run in singel shot mode, this is complementory to the batch mode for singel entries
   
   i.e python PlayHandelsregister.py -s "THYSSENKRUPP SCHULTE GMBH"  --register-number "26718" -d --download-ad -postal --outdir "C:\Users\User\Downloads\BP" -sap "2203241" -row "352"
   
   in single-shot mode (occurs automatically if -excel is not given) -s, -sap, -row are mandatory
   
   -s stands for schlagwoerter and correlates to the company name
   
   --regist-number is optional if you know the register number in the handelsregisterbuch
   
   -sap is the number we give internally in MTU can be read of the excel sheet
   
   -row correlates to the companies row in the excel sheet (so that the information can be updated in the excel sheet)



