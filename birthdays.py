# This file will contain a function that will tell you about your family members birthdays but is not complete yet

# Changed to a birthday wisher program

import datetime as dtim
import time
import smtplib
import pandas as pds
import keyboard as kbad
import pywhatkit as pith
import webbrowser as web
import pyautogui as ptui

if __name__ == '__main__':
   def sendmsg(name, to, msg):
      msg = msg+' to '+name
      dt = dtim.datetime.now()
      msgs = f'WhatsApp Message to {name} on {to} with message {msg} has been sent.'
      print(msgs)
      with open('my1logs.txt', 'a') as my1:
         my1.write(f'[{dt}] '+msg+'\n')
      web.open('https://web.whatsapp.com/send?phone=+'+to+'&text='+msg)
      time.sleep(14.5)
      ptui.press('enter')
      time.sleep(6.5)
      kbad.press_and_release('Ctrl+w')
####      pith.sendwhatmsg('+'+str(to), msg, 10, 41)
##   def sendmsg1(to, sub, msg):
##      print(f'Message to {to} with subject {sub} and message {msg} sent')
   today = dtim.datetime.now().strftime('%d-%m')
   gk = pds.read_excel('D:/Py/Birthday list.xlsx', sheet_name=0, nrows=35, index_col=0, usecols=[0, 1, 2, 3, 4, 5, 6])
#### print(gk)
   for index, item in gk.iterrows():
####      print(item)
#### print(index, item['Birthday_date'])
      bday = item['Birthday_date'].strftime('%d-%m')
      print(bday)
      if(today==bday):
         sendmsg(str(item['Name']), str(item['Phone Number']),str(item['Message']))
   kbad.press_and_release('Alt+f4')
####         web.open(item['Name'] ,item['Phone Number'], item['Message'])
