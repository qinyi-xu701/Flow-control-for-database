#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import datetime as dt
import os
import shutil
import pandas as pd
import pyodbc

# # connect SQL server
conn = pyodbc.connect('Driver={SQL Server Native Client 11.0}; Server=g7w11206g.inc.hpicorp.net; Database=CSI; Trusted_Connection=Yes;')
cursor = conn.cursor()


# In[ ]:


df = pd.read_excel(os.path.join(os.path.expanduser('~'),'desktop','deleteFromSQL.xlsx'))


# In[ ]:


for index,row in df.iterrows():
    delete_query = "DELETE FROM CSI.OPS.GPS_tbl_ops_shortage_ext WHERE ReportDate =? AND ODM=? AND Item=? AND Commodity=? AND FV=? AND Platform=? "
    cursor.execute(delete_query,(row['ReportDate'],row['ODM'],row['Item'],row['Commodity'],row['FV'],row['Platform']))
    if cursor.rowcount:
        # print(row)
        print(f"{cursor.rowcount} rows deleted from CSI.OPS.GPS_tbl_ops_shortage_ext")

conn.commit()
conn.close()

