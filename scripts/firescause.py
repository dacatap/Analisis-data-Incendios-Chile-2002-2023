# -*- coding: utf-8 -*-
"""
Created on Wed Dec  4 01:06:59 2024

@author: darca
"""
import pandas as pd
import numpy as np

def clean_conaf_exc(df, indice):
    try:
        target_row = df.iloc[:,0][df.iloc[:,0]==indice]    

        target_row = df.iloc[:,0][df.iloc[:,0]==indice]    
        
        # Create column names with the desired structure
        top_row = df.iloc[target_row.index[0]].tolist()
        second_row = df.iloc[target_row.index[0]+1].tolist()
        
        # Clean up NaN values and create meaningful column names
        cleaned_columns = []
        for top, bottom in zip(top_row, second_row):
            if pd.isna(top) and not pd.isna(bottom):
                cleaned_columns.append(bottom)
            elif not pd.isna(top) and pd.isna(bottom):
                cleaned_columns.append(top)
            elif not pd.isna(top) and not pd.isna(bottom):
                # Check if the bottom value looks like a region (e.g., 'III', 'IV')
                if isinstance(bottom, str) and bottom.strip().isalnum():
                    cleaned_columns.append(bottom)
                else:
                    cleaned_columns.append(f"{top} {bottom}".strip())
            else:
                cleaned_columns.append('')        
                
        # Drop the header rows and reset the DataFrame
        df = df.iloc[target_row.index[0]+2:].copy()
        df.columns = cleaned_columns
        
        # Set the index
        df = df.set_index(indice)
        
        if df.index[-1] is np.nan or pd.isna(df.index[-1]):
            df.index.values[-1] = 'TOTAL'
        return df
        
    except Exception as e:
        print(e)
        
        
        
#MAIN
#Add relative path to file, meanwhile the name of the file used is here, but in case of wanting to try the code yourself, you'll have to specify the route
file_path = '8.- Distribución Nacional de la Ocurrencia (N°) de Incendios Forestales según Causalidad, 2003 - 2024_octubre.xls'

dfdictionary = pd.read_excel(file_path, sheet_name=None)

#Normalizing sheet names
sheetnames = [i for i in list(dfdictionary.keys()) if "_" in i]
for i in sheetnames:
    dfdictionary[i.replace("_", " ")] = dfdictionary.pop(i)

for i in dfdictionary:
    dfdictionary[i] = clean_conaf_exc(dfdictionary[i], 'Causas general')      

with pd.ExcelWriter("CausasIncendiosClean.xlsx") as writer:
    for i in range(2023, 2002, -1):
        indexyear = "Causas " + str(i)
        df = dfdictionary[indexyear]
        df.to_excel(writer, sheet_name=indexyear, index= 'Causas general')