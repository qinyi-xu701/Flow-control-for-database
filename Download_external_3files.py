#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import datetime
import os
import win32com.client
from pathlib import Path
import re


# In[ ]:


#path and date variable
home = os.path.expanduser('~')
path = os.path.join(home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','Upload folder ( for buyer update )')


# In[ ]:


today = datetime.date.today()
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
outlook.Session.Accounts.Item(2)


# In[ ]:


# get into the correct email inbox
item = outlook.Folders.Item(2)
if item.Name != 'gpscommunication@hp.com':
    item = outlook.Folders.Item(1)
else:
    pass


# In[ ]:


#get into inbox and get emails
inbox = item.Folders['inbox']
project_folder = inbox.Folders['Newcomen']
target_folder = project_folder.Folders['Processed_Data']
messages = target_folder.Items


# In[ ]:


# sort by message sent time
messages.Sort("[Senton]")


# In[ ]:


def saveattachemnts(regex = '.*<(\d{4}-\d{2}-\d{2}) processed data>\[\'\d{8}_[&a-zA-Z ]+_[&a-zA-Z ]+_.*'):
    #regex = '<(\d{4}-\d{2}-\d{2}) processed data>\[\'\d{8}_[a-zA-Z]+_[a-zA-Z]+_.*'
    for message in messages:
        # if message.Subject == subject and message.Unread or message.Senton.date() == today:
        #assert re.match(regex, message.Subject)
        #if re.match(regex, message.Subject):
        if (message.Unread or message.Senton.date() == today) and re.match(regex, message.Subject) :
            #body_content = message.body
            #print(message.Sender.GetExchangeUser().PrimarySmtpAddress)
            #attachments = message.Attachments
            #attachment = attachments.Item(1)
            for attachment in message.Attachments:
                #print(attachment)

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
                if '_reason' in str(attachment):
                    try:
                        attachment.SaveAsFile(os.path.join(path, 'error_reason', str(message.Sender.GetExchangeUser().PrimarySmtpAddress).split('@')[0].replace('.', '_') + '_' + str(attachment)))
                        print(message.Sender.GetExchangeUser().PrimarySmtpAddress)
                    except Exception as e:
                        print(attachment)
                        print(e)
                    if re.match(regex, message.Subject) and message.Unread:
                        message.Unread = False
                    continue

                else:
                    try:                        
                        attachment.SaveAsFile(os.path.join(home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','External test destination', 'today',str(attachment)))                    
                    except Exception as e:
                        print(attachment)
                        print(e)
        else:
            pass


# In[ ]:


# saveattachemnts()


# new data --> will save to [today] <br>
# amend datat --> will save to [today] & [amend]

# In[ ]:


def saveattachemnts2(regex = '<(\d{4}-\d{2}-\d{2}) processed data>\[\'\d{8}_[&a-zA-Z ]+_[&a-zA-Z ]+_.*'):
    #regex = '<(\d{4}-\d{2}-\d{2}) processed data>\[\'\d{8}_[a-zA-Z]+_[a-zA-Z]+_.*'
    for message in messages:
        # if message.Subject == subject and message.Unread or message.Senton.date() == today:
        #assert re.match(regex, message.Subject)
        #if re.match(regex, message.Subject):
        if (message.Unread or message.Senton.date() == today) and re.match(regex, message.Subject) :
            #body_content = message.body
            #print(message.Sender.GetExchangeUser().PrimarySmtpAddress)
            #attachments = message.Attachments
            #attachment = attachments.Item(1)
            for attachment in message.Attachments:
                #print(attachment)

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
                if '_reason' in str(attachment):
                    try:
                        attachment.SaveAsFile(os.path.join(path, 'error_reason', str(message.Sender.GetExchangeUser().PrimarySmtpAddress).split('@')[0].replace('.', '_') + '_' + str(attachment)))
                        print(message.Sender.GetExchangeUser().PrimarySmtpAddress)
                    except Exception as e:
                        print(attachment)
                        print(e)
                    if re.match(regex, message.Subject) and message.Unread:
                        message.Unread = False
                    continue

                else:
                    try:                        
                        attachment.SaveAsFile(os.path.join(home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','External test destination', 'today',str(attachment)))                    
                    except Exception as e:
                        print(attachment)
                        print(e)
        elif (message.Unread or message.Senton.date() == today) and re.match('.*<(\d{4}-\d{2}-\d{2}) processed data>\[\'\d{8}_[&a-zA-Z ]+_[&a-zA-Z ]+_.*', message.Subject) :  #amend data
            #body_content = message.body
            #print(message.Sender.GetExchangeUser().PrimarySmtpAddress)
            #attachments = message.Attachments
            #attachment = attachments.Item(1)
            for attachment in message.Attachments:
                #print(attachment)

                if '_FD.xlsx' in str(attachment):
                    try:
                        attachment.SaveAsFile(os.path.join(path, 'FD_today','amend', str(attachment)))
                        attachment.SaveAsFile(os.path.join(path, 'FD_today', str(attachment)))
                    except Exception as e:
                        print(attachment)
                        print(e)
                    if re.match('.*<(\d{4}-\d{2}-\d{2}) processed data>\[\'\d{8}_[&a-zA-Z ]+_[&a-zA-Z ]+_.*', message.Subject) and message.Unread:
                        message.Unread = False
                    continue
                elif '_Shortage.xlsx' in str(attachment):
                    try: 
                        attachment.SaveAsFile(os.path.join(path, 'shortage_today','amend', str(attachment)))
                        attachment.SaveAsFile(os.path.join(path, 'shortage_today', str(attachment)))
                    except Exception as e:
                        print(attachment)
                        print(e)
                    if re.match('.*<(\d{4}-\d{2}-\d{2}) processed data>\[\'\d{8}_[&a-zA-Z ]+_[&a-zA-Z ]+_.*', message.Subject) and message.Unread:
                        message.Unread = False
                    continue
                elif '_PNbasedDetail.xlsx' in str(attachment):
                    try:
                        attachment.SaveAsFile(os.path.join(path, 'PNbasedDetail_today','amend' ,str(attachment)))
                        attachment.SaveAsFile(os.path.join(path, 'PNbasedDetail_today' ,str(attachment)))
                    except Exception as e:
                        print(attachment)
                        print(e)
                    if re.match('.*<(\d{4}-\d{2}-\d{2}) processed data>\[\'\d{8}_[&a-zA-Z ]+_[&a-zA-Z ]+_.*', message.Subject) and message.Unread:
                        message.Unread = False
                    continue
                if '_reason' in str(attachment):
                    try:
                        attachment.SaveAsFile(os.path.join(path, 'error_reason', str(message.Sender.GetExchangeUser().PrimarySmtpAddress).split('@')[0].replace('.', '_') + '_' + str(attachment)))
                        print(message.Sender.GetExchangeUser().PrimarySmtpAddress)
                    except Exception as e:
                        print(attachment)
                        print(e)
                    if re.match('.*<(\d{4}-\d{2}-\d{2}) processed data>\[\'\d{8}_[&a-zA-Z ]+_[&a-zA-Z ]+_.*', message.Subject) and message.Unread:
                        message.Unread = False
                    continue

                else:
                    try:                        
                        attachment.SaveAsFile(os.path.join(home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','External test destination', 'today','amend',str(attachment)))  
                        attachment.SaveAsFile(os.path.join(home, 'HP Inc','GPSTW SOP - 2021 日新','Project team','External test destination', 'today',str(attachment)))                     
                    except Exception as e:
                        print(attachment)
                        print(e)
        
        else:
            pass


# In[ ]:


saveattachemnts2()

