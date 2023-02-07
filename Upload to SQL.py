#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pyodbc
import pandas as pd
from pathlib import Path
import glob
import datetime
import numpy as np
from datetime import date
import math
import time


# In[2]:


# home and time
home = Path.home()
todaystr = date.today().strftime('%Y-%m-%d')


# In[3]:


conn = pyodbc.connect('Driver={SQL Server Native Client 11.0}; Server=g7w11206g.inc.hpicorp.net; Database=CSI; Trusted_Connection=Yes;')
cursor = conn.cursor()


# In[4]:


FD_path = Path( home, "desktop" , 'FD_all.xlsx')
shortage_path = Path( home, "desktop" , 'Shortage_all.xlsx')
PNbasedDetail_path = Path( home, "desktop" , 'PNbasedDetail_all.xlsx')


# In[5]:


FD = pd.read_excel( FD_path, sheet_name='Sheet1' )
shortage = pd.read_excel( shortage_path, sheet_name='Sheet1' )
PNbasedDetail = pd.read_excel( PNbasedDetail_path, sheet_name='Sheet1' )


# In[6]:


for index, row in FD.iterrows():
    f_ODM = row['ODM']
    f_Item = row['Item']
    f_Commodity = row['Commodity']
    f_FV = row['FV']
    f_Platform = row['Platform']
    f_HP_PN = row['HP_PN']
    f_Supplier = row['Supplier']
    f_PN = row['HP PN']
    f_ReportDate = row['ReportDate']
    f_FDdate = row['FDdate']
    f_FDQty = row['FDQty']
    f_BuyerName = row['BuyerName']
    
    cursor.execute(f"INSERT INTO CSI.GPS.GPS_tbl_ops_fd ( ODM,Item,Commodity,FV,Platform,Supplier,[HP PN],FDdate,FDQty,Reportdate,BuyerName )\
                    VALUES('{f_ODM}','{f_Item}','{f_Commodity}','{f_FV}','{f_Platform}','{f_Supplier}','{f_PN}','{f_FDdate}','{f_FDQty}','{f_ReportDate}','{f_BuyerName}')".replace("'NaT'", "NULL"))

conn.commit()


# In[7]:


for index, row in shortage.iterrows():
    s_ODM = row['ODM']
    s_Item = row['Item']
    s_Commodity = row['Commodity']
    s_FV = row['FV']
    s_Platform = row['Platform']
    s_P1 = row['P1']
    s_P2 = row['Net P2']
    s_P3 = row['Net P3']
    s_Total = row['Total Shortage Qty']
    s_BT = row['BT shortage']
    s_working = row['Working on upside']
    s_ReportDate = pd.to_datetime(row['ReportDate'])
    s_lastFDdate = pd.to_datetime(row['last FD date'])
    s_BuyerName = row['BuyerName']
    s_PN = row['HP_PN']

    cursor.execute(f"INSERT INTO CSI.GPS.GPS_tbl_ops_shortage_ext ( ODM,Item,Commodity,FV,Platform,P1,[Net P2],[Net P3],[Total Shortage Qty],[BT shortage],[Working on upside],ReportDate,[last FD date],HP_PN,BuyerName )\
                    VALUES('{s_ODM}','{s_Item}','{s_Commodity}','{s_FV}','{s_Platform}','{s_P1}','{s_P2}','{s_P3}','{s_Total}','{s_BT}','{s_working}','{s_ReportDate}','{s_lastFDdate}','{s_PN}','{s_BuyerName}')".replace("'NaT'", "NULL"))

conn.commit()


# In[8]:


for i in ['GPS Remark', 'ODM use column1','ODM use column2','ODM use column3','ODM use column4','ODM use column5']:
    PNbasedDetail[i] = PNbasedDetail[i].fillna("")


# In[9]:


for index, row in PNbasedDetail.iterrows():
    p_ODM = row['ODM']
    p_Item = row['Item']
    p_Commodity = row['Commodity']
    p_PN = row['HP PN']
    p_Remark = row['GPS Remark'].replace("\'", "\'\'")
    p_stock = row['852 stock']
    p_change = row['852 stock change']
    p_over = row['Over pull qty']
    p_ODM1 = str(row['ODM use column1']).replace("\'", "\'\'")
    p_ODM2 = str(row['ODM use column2']).replace("\'", "\'\'")
    p_ODM3 = str(row['ODM use column3']).replace("\'", "\'\'")
    p_ODM4 = str(row['ODM use column4']).replace("\'", "\'\'")
    p_ODM5 = str(row['ODM use column5']).replace("\'", "\'\'")
    p_ReportDate = row['ReportDate']
    p_BuyerName = row['BuyerName']

    cursor.execute(f"INSERT INTO CSI.GPS.GPS_tbl_ops_PNbasedDetail ( ODM,Item,Commodity,[HP PN],[GPS Remark],[852 stock],[852 stock change],[Over pull qty],\
                    [ODM use column1],[ODM use column2],[ODM use column3],[ODM use column4],[ODM use column5],ReportDate,BuyerName )\
                    VALUES('{p_ODM}','{p_Item}','{p_Commodity}','{p_PN}','{p_Remark}','{p_stock}','{p_change}','{p_over}','{p_ODM1}','{p_ODM2}','{p_ODM3}','{p_ODM4}','{p_ODM5}','{p_ReportDate}','{p_BuyerName}')".replace("'NaT'", "NULL"))

conn.commit()
conn.close()


# In[ ]:




