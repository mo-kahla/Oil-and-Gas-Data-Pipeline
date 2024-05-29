# -*- coding: utf-8 -*-
"""
Created on Thu Mar 21 10:16:24 2024

@author: Tan7080
"""



import pandas as pd
import numpy as np
import dateutil.parser as dparser
import glob
import re

dh = []


for infile in glob.glob("*.xlsm"):
    
        
       
            def contains_well(cell):
                if isinstance(cell, str) or isinstance(cell, bytes):
                    pattern = re.compile(r'W\s*E\s*L\s*L\s*N\s*A\s*M\s*E', re.IGNORECASE)
                    return bool(pattern.search(cell))
                else:
                    return False
      
        
            def contains_total(cell):
                if isinstance(cell, str) or isinstance(cell, bytes):
                    pattern = re.compile(r'T\s*O\s*T\s*A\s*L', re.IGNORECASE)
                    return bool(pattern.search(cell))
                else:
                    return False
                        
                 
     
        
            #injection update
            df6 = pd.read_excel(infile,sheet_name = 'Waterflooding')
            mask1 = df6.applymap(contains_well)
            mask2 = df6.applymap(contains_total)         
                    
            rows1, cols1 = np.where(mask1)
            rows2, cols2 = np.where(mask2)
            
            row1= rows1[0]
            col1 =cols1[0]
            row2=rows2[0]
            col2=cols2[0]
            process_df= df6.iloc[row1:row2,col1:18]
            final_df = process_df.iloc[1:]
            final_df.columns = process_df.iloc[0]
            final_df = final_df.reset_index(drop=True)
            final_df = final_df.iloc[:,[0,1,3,9,14]]
            final_df['WELL NAME']=final_df['WELL NAME'].str.strip()
            final_df['ZONE']=final_df['ZONE'].str.strip()
            
            d = {'(?i)Fer?daus': 'FER',
                 '(?i)Rawda': 'RAWDA',
                 '(?i)central': 'CENTRAL',
                 '(?i)Ganna': 'GAN',
                 '(?i)GAN West': 'GAN-West',
                 '(?i)Abrar': 'ABRAR',
                 '(?i) SoutH': '- S',
                 '(?i) East': '-E',
                 '(?i)AS': 'ABRAR- S',
                 '(?i)Rayan': 'RAYAN',
                 '(?i)SIDRA': 'Sidra',
                 '(?i)^S-': 'Sidra-',
                 '(?i)^A-': 'ARBAR-',
                 '(?i)^r-': 'RAWDA-',
                 '(?i)^g-': 'GAN-',
                 '(?i)Abrar -': 'ABRAR-',
                 '(?i)RAWDA-E-1': 'RAWDA-E',
                 '(?i)Abrar-86 New well': 'ABRAR-86',
                 r'\\([^)]+\\)': '',
                 '(?i)M/L-ARG & U-BAHARYA': 'M-ARG/L-ARG/U-BAHARAYA',
                 '(?i)M/L-ARG': 'M-ARG/L-ARG',
                 '(?i)U-BAHARYA': 'U-BAHARAYA',
                 '(?i)L-ARG': 'L-ARG',
                 '(?i)M-ARG': 'M-ARG',
                 
                 '(?i)ARE': 'ARE',
                 '(?i)ARG/M': 'M-ARG',
                    '(?i)ABRAR- S ':'ABRAR- S-',
                    '(?i)Sidra -':'Sidra-'
                   }
                
            
            
            for key,value in d.items():
                final_df["WELL NAME"].replace(key,value,regex=True,inplace=True)
                final_df["ZONE"].replace(key,value,regex=True,inplace=True)
            
            final_df['unique_id']=final_df['WELL NAME']+"_"+"WI"+"_"+final_df['ZONE']
            final_df = final_df.iloc[:,[5,3,4]]
            final_df.rename(columns={final_df.columns[1]: 'whp'}, inplace=True)
            final_df.rename(columns={final_df.columns[2]: 'act_inj'}, inplace=True)
            final_df.insert(loc=1, column='inj_hrs', value="")
            
                 
            
          
            
            sheet_date = dparser.parse(infile,fuzzy=True, dayfirst=True).strftime("%Y-%m-%d")
            
            final_df.insert(1, "date", sheet_date)
            final_df['date']=pd.to_datetime(final_df['date'])
            final_df['date'] = final_df['date'] - pd.Timedelta(days=1)
            final_df['whp'] = pd.to_numeric(final_df['whp'], errors='coerce')
            final_df['act_inj'] = pd.to_numeric(final_df['act_inj'], errors='coerce')
            

            
            dh.append(final_df)
            
            print('{} done'.format(sheet_date))
            ''' concatinating the dataframes from all sheets and export as csv'''
            
dh=pd.concat(dh) 
       
dh.to_csv("inj.csv")
            
print('finished inj safely')