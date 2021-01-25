# A birthday wisher program.

# Sir this porgram sends a whatsapp message but for this you need to activate whatsapp web on your pc or desktop.

# Below we are staring an infine while loop and trying to import the required modules.
stop_count = 0
while True:
   try:
# Handling if the program encounters any error during the execution.
      import numpy
# Numpy version 1.19.3 is required, Have not tested with 1.19.5 but with numpy 1.19.3 it is working fine.
      import datetime as dtim
      import time
      import os
# The above three modules come builtin with python.
      import pandas as pds
# Pandas version 1.0.5 is required
      import keyboard as kbad
# Keyboard latest version will work. but pls install the keyboard module it is very necessary in this code.
      import webbrowser as web
# Webbrowser module comes built-in.
      import xlrd
      import openpyxl
# Openpyxl latest version will satisfy.
# Xlrd version 1.2.0 is required.
      import pyautogui as ptui
# Pyautogui latest version will work.
      break
# Breaking the while loop after all the necessary modules get imported.
   except ImportError as i_encountered_an_import_error:
      # Handling Import error if the user does not have all the required modules and storing the error in i_encountered_an_import_error.
      i_encountered_an_import_error = str(i_encountered_an_import_error)
      # Here We are converting i_encountered_an_import_error to a string so that i can use it in print commands.
      print('Sorry Sir but I crashed because i faced', i_encountered_an_import_error , 'but i am trying to handle it on my own.\n')
      # Showing the error to the user and asking him if he could fix it.
      users_input = input('For Help press \'h\' :\n')
      # asking the user for help
      if users_input == 'h' or users_input == '-h' or users_input == '--h' or users_input == '-help' or users_input == '--help' or users_input == 'help' or users_input == \'h\':
         print('Sir I shall open the command prompt window to help you pls do not close it manually at any point of time I will do it automatically.\n')
         # Telling the user that I can help you and install the necessary modules on your behalf but for that I need the keyboard module of python and if you don't have the keyboard module pls install it by typing the following in your command prompt -> 'pip install keyboard'.
         time.sleep(9)
         # Adding a time.sleep() of 9 seconds so that the user gets the time to read the above message.
         os.startfile('cmd.exe')
         # Starting the command prompt to help the user by installing the required modules so that the program does not crash or stop.
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
            # We use a time.sleep() of 3 seconds so that our program gets time till the command prompt window appears because our program has to write in command prompt to install the required modules.
            # We require specific version of some modules so that we can run our program smoothly so we are using if elif and else.
            if str(name_of_module_that_could_not_be_imported) == 'pandas':
               # We use an if statement to tell our program that if the pandas module could not be imported then install it's version 1.0.5.
               keyboard.write('pip install '+str(name_of_module_that_could_not_be_imported)+'==1.0.5')
               # Here we use our keyboard module to write in the command prompt that install pandas module version - 1.0.5.
            elif str(name_of_module_that_could_not_be_imported) == 'numpy':
               # We use an elif statement to tell our program that if the numpy module could not be imported then install it's version 1.19.3.
               keyboard.write('pip install '+str(name_of_module_that_could_not_be_imported)+'==1.19.3')
               # Here we use our keyboard module to write in the command prompt that install numpy module version - 1.19.3.
            elif str(name_of_module_that_could_not_be_imported) == 'xlrd':
               # We use an elif statement to tell our program that if the xlrd module could not be imported then install it's version 1.2.0.
               keyboard.write('pip install '+str(name_of_module_that_could_not_be_imported)+'==1.2.0')
               # Here we use our keyboard module to write in the command prompt that install xlrd module version - 1.2.0.
            else:
               # We use an else statement to tell our program that if the module that could not be imported was not numpy or pandas or xlrd then install its latest version.
               keyboard.write('pip install '+str(name_of_module_that_could_not_be_imported))
               # Here we use our keyboard module to write in the command prompt that install the latest version of the module that could not be imported.
            keyboard.press_and_release('enter')
            # Here we use our keyboard module to hit the enter key on the users keyboard so that the instaling of the module starts.
            time.sleep(10)
            # Here We use a time.sleep() of 10 seconds so that the module gets the time to install and when we next time try to import it, It gets imorted and our program's execution moves forward.
            keyboard.press_and_release('alt+f4')
            # Here We press the 'Alt' and the 'f4' key on the keyboard so that the command prompt window closes as 'Alt' plus 'f4' is shortcut to close the program that is currently opened on the screen.
         except ImportError as an_error:
            # Handling the program if we find that the user does not have the keyboard module installed
            print('Sir I am opening command prompt so that i can help you pls come back here when the command prompt window appears.\n')
            # Telling the user that I was trying to help you install the necessary modules that could not be imported but because the user does not have the keyboard module installed so i can't do anything now.
            time.sleep(9)
            # Giving the user the time to read the message
            os.startfile('cmd.exe')
            # Starting the command prompt window
            print('Sorry sir but when I was trying to help you I found that you do not have the keyboard module installed but I required it to run this code could you pls install the keyboard module of python by writing this command in your command prompt : \'pip install keyboard\'\n')
            # Telling the user to pls install the keyboard module so that the program can help the user.
            time.sleep(20)
            # Waiting for the user to install the keyboard module.
      else:
         print('Sir I can\'t understand what u want to say but thanks for using...\n')
         # Handling the execution of the program if the user declined the help
         time.sleep(2)
         # Waiting for the user to read the command
         print('I am closing this program as this porgram is not working and crashing.\n')
         # Telling the user that the program cannot run due to some error and because the user has declined the help so the program cannot do anything
         time.sleep(9)
         # Again Waiting for the user to read the message
         exit(0)
         # Exiting the program as the help was denied by the user when the program told that it faced some issues
   except Exception as caught_an_exception:
      # Handling if any error other than import error occurs during the execution of the program.
      caught_an_exception = str(caught_an_exception)
      # Changing the error to a string so that we can print it nicely on the screen.
      print(f'Sorry Sir but the program crashed because {caught_an_exception} pls fix this issue or contact the owner of this repository for help -> \'github.com/Harshal-y\'')
      # Telling the user about the error and asking him to fix the error if the user know how and why the error occured or contact the owner of this repository for help on this error.
try:
   # Trying to define a function that will send the whatsapp message for the user.
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
