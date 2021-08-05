# -*- coding: utf-8 -*-
"""
Created on Wed Dec  9 12:16:24 2020

@author: Sayam
"""

from bs4 import BeautifulSoup
import requests
import xlwt 
from xlwt import Workbook 
wb = Workbook() 
  
# add_sheet is used to create sheet. 
sheet1 = wb.add_sheet('Sheet 1')
roww = 1
col = 0
for i in range(330,341): 
    url = "https://zhanglab.ccmb.med.umich.edu/BioLiP/qsearch_pdb.cgi?page=" + str(i) + "&para=lig3+%3D+%27NUC%27++order+by+id+%3D%27%27%2C+id"
    html_content = requests.get(url).text
    soup = BeautifulSoup(html_content, "lxml")

    tab=soup.findAll("table")
    
    rows = tab[0].findAll("tr")[1:]
    
    for row in rows:
         col=0;
         cells = row.findChildren('td')
         cell1 = cells[0].findChildren('a')
         cell1Text = cells[0].text
         sheet1.write(roww, col, cell1Text) 
         col+=1;
         
         cell2 = cells[1].findChildren('a')
         cell2Text = cells[1].text
         sheet1.write(roww, col, cell2Text) 
         col+=1;
         
         
                 
             
         cell3Text = cells[2].text
         sheet1.write(roww, col, cell3Text) 
         col+=1;
         
         cell4Text = cells[3].text
         sheet1.write(roww, col, cell4Text) 
         col+=1;
         
         cell5Text = cells[4].text
         sheet1.write(roww, col, cell5Text) 
         col+=1;
         
         cell6Text = cells[5].text
         sheet1.write(roww, col, cell6Text) 
         col+=1;
         
         cell7Text = cells[6].text
         sheet1.write(roww, col, cell7Text) 
         col+=1;
         
         
         cell7 = cells[6].findChildren('a')
         if len(cell7) != 0:
             url2 = cell7[0]['href']
             html_content2 = requests.get(url2).text
             soup2 = BeautifulSoup(html_content2, "lxml")
             SeqText = soup2.find("pre",{"class": "sequence"})
             if SeqText is not None:
                 SeqText = SeqText.text
                 result = ''.join([i for i in SeqText if not i.isdigit()])
                 result = result.replace(" ","")
             else:
                 result = "N/A"
         else:
             result = "N/A"
         
         cell8Text = cells[7].text
         sheet1.write(roww, col, cell8Text) 
         col+=1;
         
         cell9Text = cells[8].text
         sheet1.write(roww, col, cell9Text) 
         col+=1;
         
         
         cell3span = cells[2].find('span')['title']
         temp = ""
         for i in range(0,len(cell3span)):
             
             if(cell3span[i] != " "):
                 temp+=cell3span[i]
             elif(cell3span[i] == ' '):
                 sheet1.write(roww, col, temp)
                 col+=1
                 temp="";
                 sheet1.write(roww, col, result)
                 col-=1
                 roww+=1
         sheet1.write(roww, col, temp)
         col+=1
         sheet1.write(roww, col, result)
         col-=1
         roww+=1        



wb.save('xlwt example.xls') 