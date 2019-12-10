import win32com.client as wincl #For TTS engine
import clipboard #For copy to clipboard
import pyautogui as pya #For sending mouse clicks
import time #For delay function
from pynput import keyboard #For catching keyboard strokes
from comtypes.client import CreateObject
from os import system, name #For screen clear


speak = wincl.Dispatch("SAPI.SpVoice")

# Reads text received as an argument
def readThis(TextToRead):
    speak.Speak(TextToRead, 1) #SVSFlagsAsync = 1

# Stops Reading
def stopReading():
    speak.Speak("", 2) #SVSFPurgeBeforeSpeak = 2

# Increase Reading Speed. Returns new speed.
def incSpeed():
    if speak.Rate < 10:
        speak.Rate = speak.Rate + 1
    return speak.Rate

# Decrease Reading Speed. Returns new speed.
def decSpeed():
    if speak.Rate > -10:
        speak.Rate = speak.Rate - 1
    return speak.Rate

# Invokes mouse triple click
def tripleclick():
    pya.click()
    pya.click()
    pya.click()
    time.sleep(.01)

# Clears the console screen
def clear():
    system('cls')

# Writes the "GUI" in the console
def textGUI(Pressedkey):
    clear()
    print("F2TextToSpeech")
    print("--------------")
    print("Read Speed: [" + str(speak.Rate) + "]")
    print("Clicked Key: " + Pressedkey)
    print("")
    print("Keys")
    print("----")
    print("F2    :Read")
    print("Esc   :Stop Reading")
    print("F2    :Read")
    print("F7    :Speed Down")
    print("F8    :Speed Up")
    print("Enter :Exit")

# Exectue function based on clicked key
def on_press(key):
    # print('{0} released'.format(key))
    textGUI('{0}'.format(key))

    if key == keyboard.Key.f2:
        tripleclick()
        pya.hotkey('ctrl', 'c')
        stopReading() #optional, depends on the desired working mode
        readThis(clipboard.paste())

    if key == keyboard.Key.esc:
            stopReading()

    if key == keyboard.Key.f7:
        print(decSpeed())

    if key == keyboard.Key.f8:
        print(incSpeed())

# Welcome statement
readThis("Hi, welcome to the Simple TTS app.")

# Assiging event to function
listener = keyboard.Listener(
    on_press=on_press)

# initiating listener
listener.start()

textGUI('')
input('')
