import datetime as dt
import os
import shutil

import pyodbc

# # connect SQL server
conn = pyodbc.connect('Driver={SQL Server Native Client 11.0}; Server=g7w11206g.inc.hpicorp.net; Database=CSI; Trusted_Connection=Yes;')
cursor = conn.cursor()

datelist = [dt.datetime(2022, 11, i) for i in range(29,30)]
# print(datelist)
databaselist = ['GPS_tbl_ops_fd', 'GPS_tbl_ops_shortage_ext', 'GPS_tbl_ops_PNbasedDetail']


for _ in datelist:
    for s in databaselist:
        # delete statement
        delete_query = "DELETE FROM CSI.OPS." + s + " WHERE ReportDate = ?"
        params = (_)
        cursor.execute(delete_query, params)

        # check result
        if cursor.rowcount:
            print(_)
            print(f"{cursor.rowcount} rows deleted from CSI.OPS." + s)



conn.commit()
conn.close()