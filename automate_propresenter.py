# coding: utf-8
import os
import pyautogui
import pyperclip
import time
import dropbox
from pathlib import Path
import docx2txt
import re
import subprocess


path = './verse_demo_test.docx'
text = docx2txt.process(path)
paragraph = text.count('\n\n') + 1
result = list(filter(lambda x : x != '', text.split('\n\n')))
file = '/Applications/ProPresenter 6.app'

# pyautogui.hotkey("command", "t")

def write_splited_slide(verse):
    time.sleep(0.5)
    pyautogui.click(x=515, y=339, clicks=2)
    pyautogui.hotkey("left")
    time.sleep(0.3)
    pyperclip.copy(verse.strip()[:-1])
    pyautogui.hotkey("option", "shift", "command", "v")

def split_slide(verse, num_of_slides):
    verse_array = verse.split(",")
    verse_len = len(verse_array)
    char_count = 0
    verse_before_split = ''
    verse_after_split = ''
    for i in range(verse_len):
        char_count += len(verse_array[i])
        if (char_count >= 220):
            write_splited_slide(verse_before_split)
            write_book()
            for j in range(verse_len):
                if (j >= i):
                    verse_after_split += verse_array[j] + ', '
            open_template()
            write_splited_slide(verse_after_split)
            break
        else:
            verse_before_split += verse_array[i] + ', '

def check_len(verse):
    if (len(verse[1]) >= 220):
        split_slide(verse[1], len(verse[1])/220)
    else:
        write_verse(verse)
    
def write_verse(verse):
    time.sleep(0.5)
    pyautogui.click(x=515, y=339, clicks=2)
    pyautogui.hotkey("left")
    time.sleep(0.3)
    pyperclip.copy(verse[1].strip())
    pyautogui.hotkey("option", "shift", "command", "v")

def write_book():
    time.sleep(2)
    pyautogui.click(x=1199, y=687, clicks=2)
    pyperclip.copy(verse[0].strip())
    pyautogui.hotkey("option", "shift", "command", "v")

def click_on_x():
    time.sleep(0.5)
    pyautogui.click(x=346, y=121, clicks=1)

def click_on_editor():
    pyautogui.click(x=261, y=69, clicks=1)
    time.sleep(0.5)  

def propresenter():
    print('pro presetenter running...')

def open_template():  
    pyautogui.click(x=384, y=882, clicks=1)
    time.sleep(0.5)
    pyautogui.click(x=568, y=824, clicks=1)
    time.sleep(0.5)
    pyautogui.click(x=664, y=751, clicks=1)
    time.sleep(0.5)

def open_template_topic():
    pyautogui.click(x=384, y=882, clicks=1)
    time.sleep(0.5)
    pyautogui.click(x=568, y=824, clicks=1)
    time.sleep(0.5)
    pyautogui.click(x=677, y=812, clicks=1)
    time.sleep(0.5)

def write_template(verse):
    time.sleep(0.5)
    pyautogui.click(x=876, y=497, clicks=2)
    verse = verse[0].upper()
    pyperclip.copy(verse.strip())
    pyautogui.hotkey("option", "shift", "command", "v")


#Functions to open the Topic Template
    
open_propresenter = subprocess.call(['open', file])
open_propresenter
time.sleep(0.5)

#Limite de Palavras:

for i in range(len(result)):
#     print(result[i])
    verse = result[i]
    verse = verse.split("\\")
    print(verse)
    if len(verse) == 1:
        open_template_topic()
        write_template(verse)
    else:
        open_template()
        check_len(verse)
        # write_verse(verse)
        write_book()


# verse = "2 Crônicas 15:7-14 #. Todos os filhos de Judá muito se alegraram com esse juramento, porquanto o havia pronunciado de todo o coração. O povo buscou a Deus com a mais sincera disposição de alma, e o SENHOR permitiu que as pessoas da terra o encontrassem e lhes abençoou com paz em suas fronteiras."
# verse = verse.split("#")
# print(verse)
# if len(verse) == 1:
#     print(len(verse))
#     open_template_topic()
#     write_template()
# else:
#     open_template()
#     check_len(' '.join(verse))
#     # write_verse(verse)
#     write_book()
