#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import datetime
import os
import win32com.client
from pathlib import Path
import re


# In[ ]:


path = os.path.expanduser(os.path.join('~', 'HP Inc','GPSTW SOP - 2021 日新','Project team','Upload folder ( for buyer update )'))


# In[ ]:


today = datetime.date.today()
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

outlook.Session.Accounts.Item(2)
#.GetNamespace("MAPI")


# In[ ]:


item = outlook.Folders.Item(1)
item.Name


# In[ ]:


inbox = item.Folders['inbox']
project_folder = inbox.Folders['Newcomen']
target_folder = project_folder.Folders['Processed_Data']
messages = target_folder.Items


# In[ ]:


messages.Sort("[Senton]")


# In[ ]:


for i in messages:
    print(i.Senton)


# In[ ]:


def saveattachemnts(regex = '<(\d{4}-\d{2}-\d{2}) processed data>\[\'\d{8}_[&a-zA-Z ]+_[&a-zA-Z ]+_.*'):
    #regex = '<(\d{4}-\d{2}-\d{2}) processed data>\[\'\d{8}_[a-zA-Z]+_[a-zA-Z]+_.*'
    for message in messages:
        # if message.Subject == subject and message.Unread or message.Senton.date() == today:
        #assert re.match(regex, message.Subject)
        if re.match(regex, message.Subject) and message.Unread or message.Senton.date() == today:
            # body_content = message.body
            print(message)
            attachments = message.Attachments
            attachment = attachments.Item(1)
            for attachment in message.Attachments:
                if '_FD.xlsx' in str(attachment):
                    try:
                        attachment.SaveAsFile(os.path.join(path, 'FD_today', str(attachment)))
                    except Exception as e:
                        print(attachment)
                        print(e)
                    if re.match(regex, message.Subject) and message.Unread:
                        message.Unread = False
                    continue
                elif '_Shortage.xlsx' in str(attachment):
                    try: 
                        attachment.SaveAsFile(os.path.join(path, 'shortage_today', str(attachment)))
                    except Exception as e:
                        print(attachment)
                        print(e)
                    if re.match(regex, message.Subject) and message.Unread:
                        message.Unread = False
                    continue
                elif '_PNbasedDetail.xlsx' in str(attachment):
                    try:
                        attachment.SaveAsFile(os.path.join(path, 'PNbasedDetail_today', str(attachment)))
                    except Exception as e:
                        print(attachment)
                        print(e)
                    if re.match(regex, message.Subject) and message.Unread:
                        message.Unread = False
                    continue
                else:
                    pass
                #break
        else:
            pass
            #print(message)


# In[ ]:


saveattachemnts()


# In[ ]:




