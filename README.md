The project encompasses the following key components:
Folder Access and Data Filtering:
The Python script will access a designated folder containing PDF invoices. It will filter the documents based on a specified month or timeline, ensuring that only relevant files are processed.
PDF Data Extraction:
Using the PyPDF2 library, the script will open each PDF and extract the required financial data. The focus will be on obtaining the "Inspection and Maintenance Total" from each invoice.
Libraries Utilized:
PyPDF2: For reading PDF files and extracting text.
datetime: To manage and filter dates, ensuring the correct timeframe is selected.
time: To handle timing functions, such as processing delays if needed.
os: For interacting with the file system, such as listing files in the specified directory.
pandas: To manage and manipulate data in a structured format, enabling easy integration with Excel.
Excel Automation:
The extracted amounts will be organized into a new sheet within a designated Excel workbook, specifically for Accruals. This sheet will provide a clear view of all relevant financial figures.
Data Comparison and Validation:
The script will also implement an automated comparison between the new data and existing data within the original Excel sheet. Using pandas, the script will check for discrepancies and update the respective columns as necessary.
Additionally, it will generate notes highlighting any differences in values, providing context for discrepancies and ensuring transparency in the reporting process.
Output and Reporting:
The final output will include an updated Excel workbook containing the new Accruals sheet, with highlighted changes and discrepancies noted. This will aid stakeholders in understanding the financial data and making informed decisions.
Steps:
The expected timeline for project completion is as follows:
Step 1: Initial setup of the Python script and testing folder access.
Step 2: Development and testing of the data extraction functionality, including PDF parsing.
Step 4: Implementation of Excel automation, data comparison, and final reporting.
Deliverables
By the end of this project, the following deliverables will be provided:
A fully functional Python script that automates the extraction and reporting process.
Office scripts that make the process easy on excel workbook to deliver.
Conclusion
This project aims to make the accrual process easy and time saving. By employing Python scripts and relevant libraries, we can reduce manual effort, minimize errors, and improve the overall data collecting process for customer invoices.