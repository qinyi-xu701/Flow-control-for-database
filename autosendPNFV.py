# %%
import timeit
start = timeit.default_timer()
import win32com.client
from datetime import date
from pathlib import Path
import time
import pandas as pd
import re

# %%
home = Path.home()
#讀取今天日期
today = date.today().strftime("%Y%m%d")
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# folder = Path(home, 'HP Inc', 'GPSTW SOP - 2021 日新')

folder = Path(home, 'HP Inc', 'GPS TW Innovation - Documents','Users','BSP','Shortage management related (Ri Xin)')
emailReceiver_path = Path(home, 'HP Inc', 'GPS TW Innovation - Documents', 'Project team', 'receiver.xlsx')
emailReceiver = pd.ExcelFile(emailReceiver_path)

# %%
cc_addresses = emailReceiver.parse('cc').iloc[:, 0].dropna().tolist()
bcc_addresses = emailReceiver.parse('bcc').iloc[:, 1].dropna().tolist()

cc_text = ';'.join(cc_addresses)
bcc_text = ';'.join(bcc_addresses)

# %%
#打開新的email草稿
mail = win32com.client.Dispatch("Outlook.Application").CreateItem(0)
#讀取預設簽名檔
mail.Display()
signature = mail.Body

# %%
#新增收件人以及信件主旨
#mail.To = 'louis.lu2@hp.com'
mail.To = ''
mail.CC = cc_text
mail.BCC = bcc_text

# %%
mail.Subject = 'PN FV description mapping table update_' + today

#新增信件內容
mail.HTMLBody = '<h3>This is HTML Body</h3>'
mail.Body = "Hi receiver, \n\nPlease find latest PN FV description mapping table attached. Thanks." + signature
#將剛才最後兩封信做為附件
#mail.Attachments.Add(r'C:\Users\lulo\HP Inc\GPSTW SOP - 2021 日新\PN FV description mapping table_ALL.xlsx')
print(Path(home, folder, 'PN FV description mapping table_ALL.xlsx'))

mail.Attachments.Add(str(Path(folder, 'PN FV description mapping table_ALL.xlsx')))
time.sleep(15)
#寄信
mail.Send()
stop = timeit.default_timer()
print('RunTime: ', stop - start, ' seconds') 


