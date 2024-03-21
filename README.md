
# MAE Maybank Credit and Debit Card Statement Processing

I've created a program for the processing of Maybank bank statements for Credit and Debit PDF bank statements, converting PDFs to Excel for better tracking. 

Current solutions are scarce, so I've developed a completely free Python program with a user-friendly graphical interface.

To ensure safety, as the program reads and creates files locally, the source code is openly shared to address any concerns about malicious intent. You're welcome to review the code or use ChatGPT for a security check.

For queries, reach out at ongaunter@gmail.com.



## How the pdf processing program looks like

![alt text](/image/program_screenshot.png)



## How to organize your bank statements into a folder

![alt text](/image/bank_statement_folder.png)



## Before pressing run, this is how your program interface should look like

![alt text](/image/program_look.png)

1. ***Select Folder with PDFs***: After placing all of your Bank Statements into a single folder. Click **Browse** and navigate onto the folder where you stored all your bank statements
2. ***Select Export Path***: This is where you want to save the excel file on your computer. You can specify where you want to save it on your local computer. Or you could also save it in the same folder with the PDFs.
3. ***Enter Excel Filename (without extentsion)***:  There's no need to specify whetehr its ".xlsx" or ".csv", it will automatically save the file as CSV.
4.  ***Select Processing Mode***: The PDF structure of **Debit** and **Credit** statement are different, this is why there is a drop down option for you to select whether you wan to process Debit or Credit.

**REMINDER: DEBIT AND CREDIT BANK STATEMENTS MUST BE IN A DIFFERENT FOLDER**


## Finally, process the file by clicking on "Process Files and Export to Excel" 

1. It will process all of your pdfs that you put into the folder and convert them into Excel format for you to perform further analysis and detailed tracking of your income and expenses. 






