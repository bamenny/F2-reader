import win32com.client as wincl #For the TTS engine
import clipboard #For copying to clipboard
import pyautogui as pya #For sending mouse clicks to GUI
import time #For the delay function
from pynput import keyboard #For catching keyboard strokes
from comtypes.client import CreateObject
from os import system, name #For screen clear
from googletrans import Translator #For Translating text
from win10toast import ToastNotifier #For Desktop Notifications

#Creating Objects
speak = wincl.Dispatch("SAPI.SpVoice") #TTS Engine
translator = Translator() #For Google Translator
toaster = ToastNotifier() #For Desktop Notifications

#Reseting Variables
ReadingPaused = False #False means App Not Paused, True means App paused
Esc_counter = 0 #Being used to count how many times ESC was pressed consecutively
FirstClick = time.time()
RequireDoubleClickForReading = False #True means user must click F2 twice to activate reading

#Settings
speak.Rate = 6 #Rate between -10 to 10
DoubleClickMaxGap = 0.5 #Max duration between double press on F2
TranslationDestinationLanguage = 'he' #Traget Language for Translation
OpeningStatement = "Hi, welcome to the F2 Reader. You Rock." #opening Statement
NumberOfEscsNeededForQuitingTheApp = 3 #Number of consecutive ESCs needed to quit the app

# Reads text received as an argument
def readThis(TextToRead):
    speak.Speak(TextToRead, 1) #SVSFlagsAsync = 1
    # Other flags can be found here:
    # https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms720892(v%3Dvs.85)

# Stops Reading
def stopReading():
    speak.Speak("", 2) #SVSFPurgeBeforeSpeak = 2

# Increase Reading Speed. Returns new speed.
def incSpeed():
    if speak.Rate < 10: #10 is the max in windows
        speak.Rate = speak.Rate + 1
    return speak.Rate

# Decrease Reading Speed. Returns new speed.
def decSpeed():
    if speak.Rate > -10: #-10 is the min in windows
        speak.Rate = speak.Rate - 1
    return speak.Rate

# Invokes mouse triple click
def tripleclick():
    pya.click()
    pya.click()
    pya.click()
    time.sleep(.01)

# Invokes mouse double click
def doubleclick():
    pya.click()
    pya.click()
    time.sleep(.01)

# Clears the console screen
def clear():
    system('cls')

# Writes the "GUI" in the console
def textGUI(Pressedkey):
    global ReadingPaused
    global Esc_counter
    global DoubleClickMaxGap

    clear()

    print("F2 Reader")
    print("--------------")
    print("Read Speed: [" + str(speak.Rate) + "]")
    print("Clicked Key: " + Pressedkey)
    print("")
    print("Keys")
    print("----")
    print("F2    : Read")
    print("Esc   : Stop Reading")
    print("F6    : Pause / Double Click / Single Click")
    print("F7    : Speed Down")
    print("F8    : Speed Up")
    print(str(NumberOfEscsNeededForQuitingTheApp) + " Esc : Exit (" + str(Esc_counter) + ")" )
    print("--------------")

    if ReadingPaused == False:
        print("Status: ACTIVE")
    elif ReadingPaused == True:
        print("Status: PAUSED")

# Exectue functions based on the clicked key
def on_press(key):

    global ReadingPaused
    global Esc_counter
    global FirstClick
    global SecondClick
    global RequireDoubleClickForReading
    global TranslationDestinationLanguage

    if key == keyboard.Key.esc: #If user click the ESC key
        stopReading() #stop current outloud reading activity
        Esc_counter = Esc_counter + 1
        textGUI('{0}'.format(key))
    else: #If any other key than ESC was clicked, reset the counter
        Esc_counter = 0

    if key == keyboard.Key.f2:
        if ReadingPaused == False: #If User haven't click the pause button (F6)
            if RequireDoubleClickForReading == True:
                SecondClick = time.time()
                if SecondClick - FirstClick < DoubleClickMaxGap:
                    tripleclick()
                    pya.hotkey('ctrl', 'c')  #copy selected text to the clipboard
                    stopReading() #optional, depends on the desired working mode
                    readThis(clipboard.paste())
                FirstClick = SecondClick
            else:
                tripleclick()
                pya.hotkey('ctrl', 'c') #copy selected text to the clipboard
                stopReading() #optional, depends on the desired working mode
                readThis(clipboard.paste())

            textGUI('{0}'.format(key))

    if key == keyboard.Key.f3: #Google Translation
        if ReadingPaused == False: #If app was not paused
            doubleclick()
            pya.hotkey('ctrl', 'c') #copy selected text to the clipboard
            ResultStr = str(translator.translate(clipboard.paste(), dest=TranslationDestinationLanguage))

            #Creates Subtext of the translated word only and removes all other unnecesery text
            ResultStr = ResultStr[ResultStr.find('text=') + 5:ResultStr.find(', pron')]

            #print(''.join(reversed(ResultStr))) #If text needs to be reverserd

            toaster.show_toast("'" + clipboard.paste() + "' = '" + ResultStr + "'", "F2 Translation", threaded=True, icon_path=None, duration=5)
            textGUI('{0}'.format(key))




    if key == keyboard.Key.f6: #Pause App

        if ReadingPaused == False and RequireDoubleClickForReading == False:
            RequireDoubleClickForReading = True
            toaster.show_toast("F2 DoubleClick Activated", "DoubleClick", threaded=True, icon_path=None, duration=2)

        elif ReadingPaused == False and RequireDoubleClickForReading == True:
            RequireDoubleClickForReading = False
            ReadingPaused=True
            toaster.show_toast("F2 Paused", "Paused", threaded=True, icon_path=None, duration=2)

        elif ReadingPaused==True:
            RequireDoubleClickForReading = False
            ReadingPaused=False
            toaster.show_toast("F2 Activated", "Activated", threaded=True, icon_path=None, duration=2)

        textGUI('{0}'.format(key))

    if key == keyboard.Key.f7: #Decrease Reading Speed
        decSpeed()
        textGUI('{0}'.format(key))


    if key == keyboard.Key.f8: #Increase Reading Speed
        incSpeed()
        textGUI('{0}'.format(key))

# Welcome statement
readThis(OpeningStatement)

# Assiging event to function
listener = keyboard.Listener(on_press=on_press)

# initiating listener
listener.start()

textGUI('')

while Esc_counter < NumberOfEscsNeededForQuitingTheApp:
    time.sleep(.01)
