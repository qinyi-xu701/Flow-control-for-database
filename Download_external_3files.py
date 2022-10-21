#!/usr/bin/env python
# coding: utf-8

# In[1]:


import datetime
import os
import win32com.client
from pathlib import Path
import re


# In[2]:


path = os.path.expanduser(os.path.join('~', 'HP Inc','GPSTW SOP - 2021 日新','Project team','Upload folder ( for buyer update )'))


# In[3]:


today = datetime.date.today()
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

outlook.Session.Accounts.Item(2)
#.GetNamespace("MAPI")


# In[5]:


item = outlook.Folders.Item(2)
inbox = item.Folders['inbox']
project_folder = inbox.Folders['Newcomen']
target_folder = project_folder.Folders['Processed_Data']
messages = target_folder.Items


# In[9]:


def saveattachemnts(regex = '<(\d{4}-\d{2}-\d{2}) processed data>\[\'\d{8}_[&a-zA-Z ]+_[&a-zA-Z ]+_.*'):
    #regex = '<(\d{4}-\d{2}-\d{2}) processed data>\[\'\d{8}_[a-zA-Z]+_[a-zA-Z]+_.*'
    for message in messages:
        # if message.Subject == subject and message.Unread or message.Senton.date() == today:
        #assert re.match(regex, message.Subject)
        if re.match(regex, message.Subject):
            # body_content = message.body
            attachments = message.Attachments
            attachment = attachments.Item(1)
            for attachment in message.Attachments:
                if '_FD.xlsx' in str(attachment):
                    attachment.SaveAsFile(os.path.join(path, 'FD', str(attachment)))
                    if re.match(regex, message.Subject) and message.Unread:
                        message.Unread = False
                    continue
                elif '_Shortage.xlsx' in str(attachment):
                    attachment.SaveAsFile(os.path.join(path, 'shortage', str(attachment)))
                    if re.match(regex, message.Subject) and message.Unread:
                        message.Unread = False
                    continue
                elif '_PNbasedDetail.xlsx' in str(attachment):
                    attachment.SaveAsFile(os.path.join(path, 'PNBasedetail', str(attachment)))
                    if re.match(regex, message.Subject) and message.Unread:
                        message.Unread = False
                    continue
                else:
                    pass
                #break
        else:
            print(message)


# In[10]:


saveattachemnts()

