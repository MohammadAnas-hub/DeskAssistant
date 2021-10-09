"""
    this is my desktop assistant ("EMMA")


    NOTE: It is perfectly running in my pc so if you get any error it's not my falult
          AND plese note that you must check the path specifed (directories) as it is not possible for me to have
          all the used files in the same location and with the same name as it is in your pc

"""



"""Importing required modules

    NOTE: Note that some of the following module required pre installation

"""
from win32com.client import Dispatch
speak = Dispatch("SAPI.SpVoice")
from pygame import mixer
import os
import random
import wikipedia
import webbrowser
import datetime
import re
import smtplib
import getpass



"""
    This funtion takes a string and simply speak the given string
"""
def speaker(str):
    speak.speak(str)


"""
    This function ask user to input the task to be performed by me
"""
def task():
    user_input = input("---")
    return user_input


"""
    This function sumply add, subtract, multiply or devide the two inputed numbers
"""
def calculator():
    while True:
        x = float(input("-"))
        operation = input("-")
        y = float(input("-"))
        if operation=='+':
            print(x+y)
        elif operation=='-':
            print(x-y)
        elif operation=='*':
            print(x*y)
        elif operation=='/':
            print(x/y)
        rep = input("Do you want to continue - ")
        if rep:
            continue
        else:
            break

def fibonacci():
    inp = int(input("Press 1 to find the value at any index and "
                    " press 2 to generate the fibonacci sequence upto that index - "))
    if inp==1:
        index = int(input("Enter the index - "))
        a = 0
        b = 1
        if index==1:
            print(0)
        elif index==2:
            print(1)
        else:
            for i in range(1, index):  
                a, b = b, a+b
            print(a)

    elif inp==2:
        index = int(input("Enter the index - "))
        output = []
        a = 0
        b = 1
        if index==1:
            print(0)
        elif index==2:
            print(1)
        else:
            for i in range(index):
                output.append(str(a))
                a, b = b, a+b
        print(' '.join(output))


def factorial():
    inp = int(input("Enter an integer - "))
    val = 1

    if inp==0:
        print(1)
    elif inp==1:
        print(1)
    else:
        for i in range(1, inp+1):
            val = val * i
        print(val)
    



if __name__ == '__main__':
    while True:

        query = task()

        if 'open code' in query.lower():
            os.startfile("C:\\Users\\Anas\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Visual Studio Code\\Visual Studio Code")

        elif 'open pycharm' in query.lower():
            os.startfile("C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\JetBrains\\PyCharm Community Edition 2020.3.lnk")

        elif 'open sublime' in query.lower():
            os.startfile("C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Sublime Text 3")

        elif 'open cb' in query.lower():
            os.startfile("C:\\Users\\Anas\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\CodeBlocks\\CodeBlocks")

        elif 'open anaconda' in query.lower():
            os.startfile("C:\\Users\\Anas\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Anaconda3 (64-bit)\\Anaconda Navigator (Anaconda)")


        elif 'play music' in query.lower():
            music_list = []
            for item in os.listdir("D:\\its me\\audio"):
                music_list.append(item)
            currentmusic = random.choice(music_list)
            mixer.init()
            mixer.music.load(f"D:\\its me\\audio\\{currentmusic}")
            mixer.music.play()
            while True:

                inp = int(input("press 1 to pause, 2 to resume, 3 to next song and 4 to stop music"))
                if inp==1:
                    mixer.music.pause()
                elif inp==2:
                    mixer.music.unpause()
                elif inp==3:
                    currentmusic = random.choice(music_list)
                    # print(random.choice(os.listdir))
                    mixer.init()
                    mixer.music.load(f"D:\\its me\\audio\\{currentmusic}")
                    mixer.music.play()
                    continue

                elif inp==4:
                    mixer.music.stop()
                    break


        elif 'calculator' in query.lower():
            calculator()


        elif 'wikipedia' in query.lower():
            query.replace('wikipedia', '')
            sentences = int(input("Input the number of sentences you want - "))
            print("Searching wikipedia...")
            result = wikipedia.summary(query, sentences=sentences)
            print(result)
            speaker(result)


        elif 'open youtube' in query.lower():
            webbrowser.open("youtube.com")

        elif 'open google' in query.lower():
            webbrowser.open("google.com")

        elif 'open mail' in query.lower():
            webbrowser.open("gmail.com")

        elif 'matrix calc' in query.lower():
            webbrowser.open("matrix-calculator.net")


        elif 'write notes' in query.lower():
            print("start writing")
            note = input("--")
            with open('notes.txt', 'a') as f:
                f.write(f"{datetime.datetime.now()} --- {note}\n")


        elif 'time' in query.lower():
            print(datetime.datetime.now())
            speaker(datetime.datetime.now())

        
        elif 'factorial' in query.lower():
            factorial()

        
        elif 'fibonacci' in query.lower():
            fibonacci()


        elif 'pattern in' in query.lower():
            ptrn = input("Enter the pattern\n--")
            locn = input("Enter the location\n--")
            with open(f"{locn}", 'r') as f:
                text = f.read()
                output = re.findall(ptrn, text)
            print(output)


        elif 'send mail' in query.lower():
            smtp_object = smtplib.SMTP('smtp.gmail.com', 587)
            smtp_object.ehlo()
            smtp_object.starttls()

            email = getpass.getpass("Enter your email - ")
            password = getpass.getpass("Enter your password - ")

            smtp_object.login(email, password)
            from_address = email
            to_address = getpass.getpass("To - ")
            subject = input("Subject - ")
            content = input("Content - ")
            msg = 'Subject : ' + subject + '\n' + content
            smtp_object.sendmail(from_address, to_address, msg)
            smtp_object.quit()

        elif 'QUIT' in query.upper():
            break