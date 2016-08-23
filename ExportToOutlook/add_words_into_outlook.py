"""Add new words in to Microsoft Outlook.

Add words as tasks into Outlook task list.

Usage: Open Outlook. Run add_words_into_outlook.py
"""
__author__ = "Chen, Cho-Yi (ntu.joey@gmail.com)"
__date__ = "2012-02-07"
__version__ = "0.1.1"

import win32com.client

try:
	oOutlookApp = win32com.client.gencache.EnsureDispatch("Outlook.Application")
except:
	print("MSOutlook: unable to load Outlook")

WORDS_FILE = 'input_word_list.txt'

for i, line in enumerate(open(WORDS_FILE)):
	# get the word & its definition
	word, defi = line.strip().split(' ')
	print word, defi
	
	# add the word into Outlook task list
	newtask = oOutlookApp.CreateItem(win32com.client.constants.olTaskItem)
	newtask.Subject = word
	newtask.Body = defi
	newtask.Save()
	#newtask.Close()

print('Done! Total {} new words added.'.format(i))

