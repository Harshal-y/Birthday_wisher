# A birthday wisher program.

# Sir this porgram sends a whatsapp message but for this you need to activate whatsapp web on your pc or desktop.
while True:
   try:
      import numpy
# numpy version 1.19.3 is required, Have not tested with 1.19.5 but with numpy 1.19.3 it is working fine.
      import datetime as dtim
      import time
      import os
#The above three modules come builtin with python.
      import pandas as pds
#pandas version 1.0.5 is required
      import keyboard as kbad
#keyboard latest version will work. but pls install the keyboard module it is very necessary in this code.
      import webbrowser as web
# webbrowser module comes built-in.
      import xlrd
      import openpyxl
# openpyxl latest version will satisfy.
# xlrd version 1.2.0 is required.
      import pyautogui as ptui
      break
#pyautogui latest version will work.
   except ImportError as i_encountered_an_import_error:
      # Handling Import error if the user does not have all the required modules and storing the error in i_encountered_an_import_error.
      i_encountered_an_import_error = str(i_encountered_an_import_error)
      # converting i_encountered_an_import_error to a string so that i can use it in print commands.
      print('Sorry Sir but I crashed because i faced', i_encountered_an_import_error , 'but i am trying to handle it.\n')
      # Showing the error to the user
      users_input = input('For Help press \'h\' :\n')
      # asking the user for help
      if users_input == 'h' or users_input == '-h' or users_input == '--h' or users_input == '-help' or users_input == '--help' or users_input == 'help':
         print('Sir I shall open the command prompt window to help you pls do not close it manually at any point of time I will do it automatically.\n')
         # Telling the user that I can help you and install the necessary modules on your behalf but for that I need the keyboard module of python and if you don't have the keyboard module pls install it by typing the following in your command prompt -> 'pip install keyboard'.
         time.sleep(9)
         # Adding a time.sleep() of 9 seconds so that the user gets the time to read the above message.
         os.startfile('cmd.exe')
         # starting the command prompt to help the user by installing the required modules so that the program does not crash or stop.
         try:
            import keyboard
            # Checking if the user has keyboard module installed in his system or not as the program requires it to install other required modules if the user faces an ImportError.
            start_index = i_encountered_an_import_error.find('\'')
            # Using the following method to get the name of the module that could not be imported -> str(start_index:end_index)
            # Here we use the above method to find the start index that is if the error is like -> No module named 'so and so for eg. pandas', so the name of the module starts after "'" so we find the index of "'" and add 1 to it so that our program knows where the name of the module starts from.
            name_of_module_that_could_not_be_imported = i_encountered_an_import_error[start_index+1:len(i_encountered_an_import_error)-1]
            # In the above line we used len() function of python that returns the total characters present in that string. Next we subtract 1 from the length of the total characters present in the string because this returns the second last word of our string. Below we have an example to better explain this line :
            # We did -> i_encountered_an_import_error[start_index+1:len(i_encountered_an_import_error)] but this returned the name of the module that could not be imported plus "'" character.
            time.sleep(3)
            # we used a time.sleep() of 3 seconds so that our program gets time till the command prompt window appears because our program has to write in command prompt to install the required modules.
            if str(name_of_module_that_could_not_be_imported) == 'pandas':
               # we use an if statement to tell our program that if pandas module could not be imported then install it's version 1.0.5
               keyboard.write('pip install '+str(name_of_module_that_could_not_be_imported)+'==1.0.5')
               # Here we write in the command prompt to install the pandas module if not found or could not be imported.
            elif str(name_of_module_that_could_not_be_imported) == 'numpy':
               keyboard.write('pip install '+str(name_of_module_that_could_not_be_imported)+'==1.19.3')
            elif str(name_of_module_that_could_not_be_imported) == 'xlrd':
               keyboard.write('pip install '+str(name_of_module_that_could_not_be_imported)+'==1.2.0')
            else:
               keyboard.write('pip install '+str(name_of_module_that_could_not_be_imported))
            keyboard.press_and_release('enter')
            time.sleep(10)
            keyboard.press_and_release('alt+f4')
         except ImportError as an_error:
            print('Sir I am opening command prompt so that i can help you pls come back here when the command prompt window appears.\n')
            time.sleep(9)
            os.startfile('cmd.exe')
            print('Sorry sir but when I was trying to help you I found that you do not have the keyboard module installed but I required it to run this code could you pls install the keyboard module of python by writing this command in your command prompt : \'pip install keyboard\'\n')
      else:
         print('Sir I can\'t understand what u want to say but thanks for using...\n')
         time.sleep(2)
         print('I am closing this program as this porgram is not working and crashing.\n')
         time.sleep(9)
         exit(0)
   except Exception as caught_an_exception:
      caught_an_exception = str(caught_an_exception)
      print(f'Sorry Sir but the program crashed because {caught_an_exception} pls fix this issue or contact the owner of this repository for help -> \'github.com/Harshal-y\'')
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
   try:
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
   except Exception as exception_occured:
      exception_occured = str(exception_occured)
      print('Sorry Sir but I crashed because :', exception_occured, 'pls fix this issue or conatct the owner of this repository for help - \'github.com/Harshal-y\'')
