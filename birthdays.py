# A birthday wisher program.

# Sir this porgram sends a whatsapp message but for this you need to activate whatsapp web on your pc or desktop.

# Below we are staring an infine while loop and trying to import the required modules.
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
      # This function takes three parameter :
      # 1st: name of the person to whom the program has to wish.
      # 2nd: the phone number of the person to whom the message has to be delivered.
      # 3rd: message the user wants to say to the person who's birthday is.
      # Converting the message to a message like -> 'Happy birthday' + to + 'Harry' -> Happy birthday to Harry.
      dt = dtim.datetime.now()
      # getting the current date and time.
      msgs = f'WhatsApp Message to {name} on {to} with message {msg} has been sent.'
      # Telling the user about who's birthday is today and what message is being sent to them.
      print(msgs)
      # displaying the message to the user
      with open('my1logs.txt', 'a') as my1:
         # creating a file where we will save to whom we have sent birthday message and when.
         my1.write(f'[{dt}] '+msg+'\n')
         # writing in a file that on this date this time this message was sent
      web.open('https://web.whatsapp.com/send?phone=+'+to+'&text='+msg)
      # Sending the whatsapp messge through pc by opening web.whatsapp.com on the browser.
      while True:
         # Again Starting an infinite while loop
         ag = ptui.locateCenterOnScreen('web_whatsapp_send_msg_img_2.png')
         # Locating the send button of whatsapp web on the screen.
         if ag is None:
            print('Could not send Message')
            # If button not found then we wait for 1 second and try if we could find it this time.
            time.sleep(1)
            pass
         elif ag is not None:
            # If the button was found on the screen.
            ptui.moveTo(ag)
            # Move the cursor to the button.
            ptui.click()
            # Click the button finally and send the whatsapp message.
            break
         else:
            print('Could not send Message')
            ag = ptui.locateCenterOnScreen('web_whatsapp_send_msg_img_2.png')
            # If we fail to find the button then here we try relocating it on the screen.
      time.sleep(5.5)
      # Waiting for the message to get sent completely.
      kbad.press_and_release('ctrl+w')
      # Closing the browser.
except Exception as i_crashed:
   # Handling if any exception occurs while trying to send the message.
   i_crashed = str(i_crashed)
   # Converting the exception to a string.
   print(f'Sir I faced :',i_crashed,'issue so the program crashed')
   # Telling the user what error occured during the execution of the program. 

if __name__ == '__main__':
   # Creating a if name == main so that if we try to import the function anywhere else the below code does not get run automatically there also.
   try:
      # Trying to catch error if any.
      today = dtim.datetime.now().strftime('%d-%m')
      # Getting the current date and month.
      yearNow = dtim.datetime.now().strftime('%Y')
      # Getting the current year so that to know if we have sent message in this year or not.
      gk = pds.read_excel('Birthday list.xlsx', sheet_name=0, index_col=0, usecols=[0, 1, 2, 3, 4, 5, 6])
      wit = []
      # Creating a empty list to get the index of the name of the person whom to send birthday message.
      for index, item in gk.iterrows():
         # Getting the index and the item from the excel file where we have stored the birthday information of the friends.
         bday = item['Birthday_date'].strftime('%d-%m')
         # Getting the birthday date and month on the person.
         print(bday)
         # Showing the user the birtday dates and years of his friends or family members.
         if(today==bday) and yearNow not in str(item['Year Wished']):
            # Checking if today is the birthday of any of his friends and if we have already wished them or not.
            sendmsg(str(item['Name']), str(item['Phone Number']),str(item['Message']))
            # If today is the birthday and we have not yet wished the person then preparring to send the birthday message by calling our sendmsg function and passing the nasme phone number and mesage arguement.
            wit.append(index)
            # Here we are adding the index of the name of the person who's birthday is today to our wit list.
      for i in wit:
         # Here we are updating our excel file after sending the birthday message to the person and adding the year we have wished them in our excel file so that we do not resend the message to him.
         yr = gk.loc[i, 'Year Wished']
         gk.loc[i, 'Year Wished'] = str(yr) + ',' + str(yearNow)
      gk.to_excel('Birthday list.xlsx')
      # Here we have finally updated our excel file.
   except Exception as exception_occured:
      # Handling errors (if any) that occur during the execution of the program.
      exception_occured = str(exception_occured)
      # Converting the error to a string.
      print('Sorry Sir but I crashed because :', exception_occured, 'pls fix this issue or conatct the owner of this repository for help - \'github.com/Harshal-y\'')
      # Telling the user what error has occured and how to resolve it.


# Hope you did not face any errors with this program.



