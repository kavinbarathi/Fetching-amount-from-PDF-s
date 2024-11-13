import datetime

import time

import os

from PyPDF2 import PdfReader

import pandas as pd
 
path = str(input("Enter the File Path :"))
 
print("Processing..............!","Please wait for sometime",sep="\n\n")
 
file_list = os.listdir(path)
 
ls = []

date = str(datetime.date.today()).split("-")

current_month = int(date[1])

date[1]=str(int(date[1])-1)
 
month = datetime.date(int(date[0]),int(date[1]),int(date[2])).strftime("%b")
 
pdf_file_list=[]
 
for pdf in file_list:

    if pdf.endswith(".pdf"):

        modified_timestamp = os.path.getmtime(path+"\\"+pdf)

        modified_date = str(datetime.datetime.fromtimestamp(modified_timestamp)).split()[0]

        modified_month = int(modified_date.split("-")[1])

        if modified_month in [current_month,current_month-1]:

            pdf_file_list.append(pdf)
 
file_count = len(pdf_file_list)
 
for file in pdf_file_list:

    reader = PdfReader(path+'\\'+file)

    pages = reader.pages

    for page in pages:

        data = page.extract_text()

        splitted_data = data.split('\n')

        modified_timestamp = os.path.getmtime(path+"\\"+file)

        modified_date = str(datetime.datetime.fromtimestamp(modified_timestamp)).split()[0]

        for i in range(len(splitted_data)):

            if "Inspection & Maintenance  Total $" in splitted_data[i]:

                job_number = int("".join([i for i in file.split('_')[0] if i.isdigit()]))

                amount = (splitted_data[i+3].replace("$","")).replace(",","")

                ls.append([job_number,amount,modified_date])            
 
Total = pd.DataFrame(ls,columns=["Job Number","Amount As Per PDF", "Last Modified Date"],index=None)
 
Total[["Modified Year","Modified Month","Modified Date"]] = Total["Last Modified Date"].str.split("-",expand=True)
 
Total = Total.astype({"Job Number":int,"Amount As Per PDF":float,"Last Modified Date":str,"Modified Year":int,"Modified Month":int,"Modified Date":int})
 
Total = Total.sort_values(["Modified Year","Modified Month","Modified Date"],ascending=False)
 
Total = Total[~Total[["Job Number"]].duplicated()]
 
Total = Total.drop(columns=["Modified Year","Modified Month","Modified Date"])
 
if "Accruals for "+month+".xlsx" in os.listdir():

        os.remove("Accruals for "+month+".xlsx")
 
print("The Data is Ready")
 
target_file_path = str(input("Enter the target Excel file's path : "))
 
target_file_path = target_file_path.replace('"','')
 
with pd.ExcelWriter(target_file_path, mode="a" , engine="openpyxl") as writer:
 
    Total.to_excel(writer, sheet_name="Accruals",index = False)
 