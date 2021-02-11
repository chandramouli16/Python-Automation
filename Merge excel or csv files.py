#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import pandas as pd

from tkinter import Tk    
from tkinter.filedialog import askdirectory


# In[2]:


print("All files should have same column names and same number of rows to skip\n")

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
path = askdirectory(title="Select the folder which contains all excel files") # show an "Open" dialog box and return the path to the selected file
print("\nFolder path is : ",path)


# In[3]:


df_final = pd.DataFrame()
skip_rows = int(input("\nNumber of rows to skip : "))


# In[4]:


try:
    for file in os.listdir(path):
        ext = file.split(".")[1]
        if ext in ["xlsx","xls"]:
            df_file=pd.read_excel(os.path.join(path,file), skiprows=skip_rows, index_col=[0])
        elif ext == "csv":
            df_file=pd.read_csv(os.path.join(path,file), skiprows=skip_rows, index_col=[0])
        else:
            print(ext," Extension not Supported")
        df_final = pd.concat([df_final,df_file],ignore_index=True)

    new_filename = input("\nEnter new merged file name : ")+".xlsx"

    df_final.to_excel(new_filename,index=False)
    print("\nFile Created")
    
except:
    print("Error Merging files")

