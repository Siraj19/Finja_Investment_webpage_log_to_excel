from bs4 import BeautifulSoup
import os
import xlsxwriter

#finding text file for parsing
source_dir=os.getcwd()
for item in os.listdir():
    if item.endswith('txt'):
        text_file = open(item, "r")
        break
#Loading text file into variable
data = text_file.read()
text_file.close()
#Cleaning used text file
file = open(item,"r+")
file.truncate(0)
file.close()
#Converting into parsable data
soup = BeautifulSoup(data, "html.parser")
#Picking up desired data class
investments=soup.find_all('div',{"class":"row align-items-center"})
#Creating an excel file
workbook = xlsxwriter.Workbook('finja_investment_record.xlsx')
worksheet = workbook.add_worksheet()
#Initializing excel coordinates
row=0
col=0
#Writing relevant data to the excel file by loading data from choosen class
worksheet.write(row, col, "Due Date")
worksheet.write(row, col+1, "Amount")
worksheet.write(row, col+2, "Profit")
worksheet.write(row, col+3, "Days Left")
row+=1
for investment in investments[1:]:
    temp=""
    temp=str(investment.find('h2',{"id":"due_date"}))
    worksheet.write(row, col,temp[18:len(temp)-5])
    worksheet.write(row, col+3,"=DATE("+temp[18:len(temp)-5].replace('-',',')+")-TODAY()")
    temp=str(investment.find('h2',{"id":"amount"}))
    worksheet.write(row, col+1, "="+temp[20:len(temp)-5])
    temp=str(investment.find('h2',{"id":"profit"}))
    worksheet.write(row, col+2, "="+temp[20:len(temp)-5])
    row+=1

row+=1
worksheet.write(row, col+1, "Amount")
worksheet.write(row, col+2, "Profit")
row+=1
worksheet.write(row, col+1, "=SUM(B2:B"+str(row-2)+")")
worksheet.write(row, col+2, "=SUM(C2:C"+str(row-2)+")")
workbook.close()