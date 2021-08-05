# -*- coding: utf-8 -*-
"""
Created on Fri Jan  1 21:36:46 2021

@author: Sayam
"""

import pandas as pd
df = pd.read_excel(r'C:\desktop Items\output3.xlsx', index_col=0)
df['Fill Combination']=""
errors = []
#df['Combination'] = ""
for j in range(0,971960):

    siteNo = df.iloc[j][8]
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
                        a=0
                        b=0
                        combination=""
                        start = 0
                        end = int(temp)+20;
                        if int(temp)-21 < 0:
                            a=(int(temp)-21)*-1
                            for k in range(0,a):
                                combination+="X"
                            start = 0;
                        else:
                            start = int(temp)-21
                            
                        if int(temp)+20 > len(sequence):
                            b=(int(temp)+20)-len(sequence)
                            end = len(sequence);
                        else:
                            end = int(temp)+20
                        
                        
                        for i in range(start,end):
                            combination+=sequence[i];
                        for k in range(0,b):
                            combination+="X"
                        print(combination)
                        print(j)
                        df.iloc[j, df.columns.get_loc('Fill Combination')] = combination
                except:
                    errors.append([siteNo,j+2])

df.to_excel("C:\desktop Items\output4.xlsx")  

errordf = pd.DataFrame(errors,columns=['Error', 'Row'])
errordf.to_excel("C:\desktop Items\errors2.xlsx")