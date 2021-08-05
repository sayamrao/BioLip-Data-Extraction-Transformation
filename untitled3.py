# -*- coding: utf-8 -*-
"""
Created on Mon Jan  4 17:47:18 2021

@author: Sayam
"""

import pandas as pd
df = pd.read_excel(r'C:\desktop Items\output4.xlsx', index_col=0)
df["NewSite"] = ""
errors = []
i = 971929
while i < 971948:
    uniprot= df.iloc[i][5]
    if not pd.isna(uniprot):
        try:
            siteDetails =  []
            newSite= []
            if not pd.isna(df.iloc[i][8]):
                siteDetails.append(df.iloc[i][8][1:] if df.iloc[i][8][0] ==' ' else df.iloc[i][8])
            
            sequence = df.iloc[i][9]
            j = i+1
        
        
            while pd.isna(df.iloc[j][5]):
                print("HI@")
                if not pd.isna(df.iloc[j][8]):
                    print(df.iloc[j][8][1:])
                    print("I : " + str(i))
                    print("jjjj : " + str(j))
                    siteDetails.append(df.iloc[j][8][1:] if df.iloc[j][8][0] ==' ' else df.iloc[j][8])
                    print("HI")
                j+=1
              
            print(siteDetails[0])
            
            for k in range(0, len(siteDetails)-1):
                a= int(siteDetails[k][1:])
                b= int(siteDetails[k+1][1:])
                while (b-a) > 81:
                    temp = a+20+21;
                    
                    newSite.append(sequence[temp] + str(temp));
                    a= temp
            index = i
            
            for l in range(0,len(newSite)):
                df.iloc[index, df.columns.get_loc('NewSite')] = newSite[l]
                index+=1;
            i=j-1
            print("Value of j " + str(j))
            print("ith value: " + str(i))
            i+=1 
        except:
            errors.append([uniprot,i+2]);
            print("Exception")
            i+=1
            while not pd.isna(df.iloc[i][5]):
                i+=1
    else:
        i+=1
df.to_excel("C:\desktop Items\output6.xlsx")  

errordf = pd.DataFrame(errors,columns=['Error', 'Row'])
errordf.to_excel("C:\desktop Items\errorsNew2.xlsx")

