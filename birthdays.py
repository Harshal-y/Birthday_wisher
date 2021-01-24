# A birthday wisher program.

# Sir this porgram sends a whatsapp message but for this you need to activate whatsapp web.
while True:
   try:
      import numpy
      import datetime as dtim
      import time
      import os
#The above three modules come builtin with python
      import pandas as pds
#pandas version 1.0.5 is required
      import keyboard as kbad
#keyboard any version will work. but pls install the keyboard module it is very necessary in this code.
      import webbrowser as web
      import xlrd
      import openpyxl
# this module comes builtin
      import pyautogui as ptui
      break
#pyautogui any version will work
   except ImportError as i_recieved_an_import_error:
      i_recieved_an_import_error = str(i_recieved_an_import_error)
      print('Sorry Sir but I crashed because i faced',i_recieved_an_import_error , 'but i needed that module in order to run the code pls install it or let me help you.\n')
      print('Sir I shall open the command prompt window to help you pls do not close it manually at any point of time we will do it automatically.\n')
      users_input = input('For Help press \'h\' :\n')
      if users_input == 'h':
         os.startfile('cmd.exe')
         try:
            import keyboard
            import_error_1 = i_recieved_an_import_error.find('\'')
            import_error = i_recieved_an_import_error[import_error_1+1:len(i_recieved_an_import_error)-1]
            time.sleep(3)
            if str(import_error) == 'pandas':
               keyboard.write('pip install '+str(import_error)+'==1.0.5')
            elif str(import_error) == 'numpy':
               keyboard.write('pip install '+str(import_error)+'==1.19.3')
            elif str(import_error) == 'xlrd':
               keyboard.write('pip install '+str(import_error)+'==1.2.0')
            else:
               keyboard.write('pip install '+str(import_error))
            keyboard.press_and_release('enter')
            time.sleep(10)
            keyboard.press_and_release('alt+f4')
         except ImportError as an_error:
            print('Sir I am opening command prompt so that i can help you pls come back here when the command prompt window appears.\n')
            time.sleep(12)
            print('Sorry sir but you need to install the keyboard module of python with this command : pip install keyboard so that I can help you further.\n')
      else:
         print('Sir I can\'t understand what u want to say but thanks for using...\n')
         time.sleep(2)
         print('I am closing this program as this porgram is not working and crashing.\n')
         time.sleep(9)
         exit(0)
try:
   def sendmsg(name, to, msg):
      msg = msg+' to '+name
      dt = dtim.datetime.now()
      msgs = f'WhatsApp Message to {name} on {to} with message {msg} has been sent.'
      print(msgs)
      with open('my1logs.txt', 'a') as my1:
         my1.write(f'[{dt}] '+msg+'\n')
      web.open('https://web.whatsapp.com/send?phone=+'+to+'&text='+msg)
      while True:
         ag = ptui.locateCenterOnScreen('web_whatsapp_send_msg_img_2.png')
         if ag is None:
            time.sleep(1)
            pass
         elif ag is not None:
            ptui.moveTo(ag)
            ptui.click()
            break
         else:
            ag = ptui.locateCenterOnScreen('web_whatsapp_send_msg_img_2.png')
      time.sleep(5.5)
      kbad.press_and_release('ctrl+w')
except Exception as i_crashed:
   i_crashed = str(i_crashed)
   print(f'Sir I faced :',i_crashed,'issue so the program crashed')

if __name__ == '__main__':
   today = dtim.datetime.now().strftime('%d-%m')
   yearNow = dtim.datetime.now().strftime('%Y')
   gk = pds.read_excel('Birthday list.xlsx', sheet_name=0, nrows=35, index_col=0, usecols=[0, 1, 2, 3, 4, 5, 6])
   wit = []
   for index, item in gk.iterrows():
      bday = item['Birthday_date'].strftime('%d-%m')
      print(bday)
      if(today==bday) and yearNow not in str(item['Year Wished']):
         sendmsg(str(item['Name']), str(item['Phone Number']),str(item['Message']))
         wit.append(index)
   for i in wit:
      yr = gk.loc[i, 'Year Wished']
      gk.loc[i, 'Year Wished'] = str(yr) + ',' + str(yearNow)
   gk.to_excel('Birthday list.xlsx')
