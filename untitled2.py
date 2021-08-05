# -*- coding: utf-8 -*-
"""
Created on Sun Dec 13 13:38:55 2020

@author: Sayam
"""

import pandas as pd
df = pd.read_excel(r'C:\desktop Items\output5.xlsx', index_col=0)
#print('vfv')
#df.drop(df['Site # Detail'=="R77A"].index)
#df = df.iloc[: , [0, 1, 2, 3, 4, 5, 6, 7,8, 9]].copy()
errors = []
df['New Combination'] = ""
for j in range(0,971960):

    siteNo = df.iloc[j][12]
    if not type(int) or not pd.isna(siteNo):
        if siteNo != "Break":
            if siteNo != "After Break":
                try:
                    if not pd.isna(siteNo):
                            
                        if(siteNo[0] == ' '):
                            siteNo = siteNo[1:]
     
                    sequence = df.iloc[j][9]
                    if not pd.isna(sequence):
                        if(sequence[0] == ' '):
                            sequence = sequence[1:]
                    if not pd.isna(sequence):
                        print(df.iloc[j][8][0])
                        character=df.iloc[j][8][0]
                        temp = siteNo[1]
                        try:
                            if siteNo[2] is not None:
                                temp+=siteNo[2]
                                
                        except:
                            print('exception')
                        
                        try:
                            if siteNo[3] is not None:
                                temp+=siteNo[3]
                        except:
                            print('exception')
                        try:
                            if siteNo[4] is not None:
                                temp+=siteNo[4]
    
                        except:
                            print('exception')
                            
                            start = 0
                            end = int(temp)+20;
                            if int(temp)-21 < 0:
                                start = 0;
                            else:
                                start = int(temp)-21
                                
                            if int(temp)+20 > len(sequence):
                                end = len(sequence);
                            else:
                                end = int(temp)+20
                            
                            combination = '';
                            for i in range(start,end):
                                combination+=sequence[i];           
                            print(combination)
                            print(j)
                            df.iloc[j, df.columns.get_loc('New Combination')] = combination
                except:
                    errors.append([siteNo,j+2])

df.to_excel("C:\desktop Items\output7.xlsx")  

errordf = pd.DataFrame(errors,columns=['Error', 'Row'])
errordf.to_excel("C:\desktop Items\errors7.xlsx")
