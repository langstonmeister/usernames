#!/usr/bin/env python
# coding: utf-8

# In[1]:


import openpyxl
import os
import re
from openpyxl import Workbook


# In[2]:


path = '/Users/langston.kyle/Downloads'
os.chdir(path)


# In[3]:


file = input('Name of file: ')
wb = openpyxl.load_workbook(file)
source = wb['Sheet1']
s = source['E']
wb.create_sheet(title='usernames')
ws = wb['usernames']
np = []
course = source['A5'].value
first = []
last = []
students = {}


# In[4]:


words = course.split()
course = words[1]
print(course)


# In[5]:


source.delete_rows(1, 6)
np = ['username', 'password', 'first', 'last', 'role', 'email', 'class']

for rowOfCellObjects in source['E1':'E' + str(len(source['E'])-1)]:
    for cellObj in rowOfCellObjects:
        
        st = cellObj.value
        wholeName = cellObj.value
        if st != None:
            if 'Password' in st:
                continue
            #if '/' not in st:
            #    continue
            #wholeName = st
            listNames = re.split('\W', wholeName)
            listNames = [string for string in listNames if string != ""]
            #listNames = list(filter(None, listNames))
            try:
                first = listNames[2]
                last = listNames[1]
            except:
                print('issue here')
                pass
            np = re.split(' / ', st)
            np.append(first.title())
            np.append(last.title())
        np.append('student')
        np.append('')
        np.append(course)
        print(np)
        #print(listNames)
        ws.append(np)
        
        
        


# In[6]:


wb.save(filename = 'MusicFirstUsers.xlsx')


# In[ ]:




