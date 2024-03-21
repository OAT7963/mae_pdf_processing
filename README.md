
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

1. ***Select Folder with PDFs:*** Place all your bank statements in a single folder. Click "Browse" to locate and select this folder.
2. ***Select Export Path:*** Choose where you want to save the exported Excel file. You can save it either on your local computer or in the same folder as the PDFs.
3. ***Enter Excel Filename:*** Input the desired name for your Excel file without adding an extension like ".xlsx" or ".csv"; the file will automatically be saved as CSV.
4. ***Select Processing Mode:*** Since the structure of debit and credit statements varies, use the dropdown menu to specify whether you want to process a debit or credit statement.

**REMINDER: DEBIT AND CREDIT BANK STATEMENTS MUST BE IN A DIFFERENT FOLDER**


## Finally, process the file by clicking on "Process Files and Export to Excel" 

1. It will process all of your pdfs that you put into the folder and convert them into Excel format for you to perform further analysis and detailed tracking of your income and expenses.


## OUTPUT - How the Excel File output looks like

The below is for Credit Card Bank Statment. As for Debit card, you will see additional columns like "Money in, Money Out" and "Statement Balance"

![alt text](/image/output_cc_example.png)






