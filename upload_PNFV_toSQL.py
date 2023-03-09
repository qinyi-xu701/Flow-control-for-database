#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pyodbc
import pandas as pd
from pathlib import Path
import glob
import datetime
import re
import os
import numpy as np
from pandas.api.types import is_string_dtype
from pandas.api.types import is_numeric_dtype
from datetime import date
import math
import time


# In[ ]:


# home and time
home = Path.home()
todaystr = date.today().strftime('%Y-%m-%d')
PNFV_alternative = pd.read_excel(Path(home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','PNFV', 'alternative.xlsx'))
PNFV = pd.read_excel(Path(home, 'HP Inc','GPSTW SOP - 2021 日新','PN FV description mapping table_ALL.xlsx'))


# In[ ]:


PNFV = pd.merge(PNFV, PNFV_alternative, on = 'Descr', how = 'left')


# In[ ]:


PNFV


# In[ ]:


# PNFV.to_excel(Path(home, 'HP Inc','GPSTW SOP - 2021 日新','PN FV description mapping table_ALL.xls'), index = False)


# In[ ]:


start_time = time.time()
conn = pyodbc.connect('Driver={SQL Server Native Client 11.0}; Server=g7w11206g.inc.hpicorp.net; Database=CSI; Trusted_Connection=Yes;')
cursor = conn.cursor()

cursor.execute(f"SELECT COUNT(*) FROM OPS.GPS_tbl_ops_PN_FV")
conn.commit()

cursor.execute(f"DELETE FROM OPS.GPS_tbl_ops_PN_FV")
conn.commit()
print("%s seconds ---" % (time.time() - start_time))

cursor.execute(f"SELECT COUNT(*) FROM OPS.GPS_tbl_ops_PN_FV")
conn.commit()

for index, row in PNFV.iterrows():

    Commodity = str(row['Commodity'])
    Supplier = str(row['Supplier'])
    PN = str(row['PN'])
    Descr = str(row['Descr'])
    alternative = str(row['alternative part flag'])

    cursor.execute(f"INSERT INTO CSI.OPS.GPS_tbl_ops_PN_FV ( Commodity, Supplier, PN, Descr, [alternative part flag] )\
                    VALUES('{Commodity}','{Supplier}','{PN}','{Descr}','{alternative}')")
    
    print("%s seconds ---" % (time.time() - start_time))
conn.commit()
conn.close()
print("%s seconds ---" % (time.time() - start_time))

