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
F2FirstClick = time.time()
F4FirstClick = time.time()
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
    print("F2    : Hover Read")
    print("F4 x2 : Read Selected / Marked Text")
    print("Esc   : Stop Reading")
    print("F6    : Pause / Activate")
    print("F7    : Speed Down")
    print("F8    : Speed Up")
    print("F9    : Double / Single Click more")
    print(str(NumberOfEscsNeededForQuitingTheApp) + " Esc : Exit (" + str(Esc_counter) + ")" )
    print("--------------")

    if ReadingPaused == False:
        StatusString = "Status: ACTIVE"

        if RequireDoubleClickForReading == False:
            StatusString = StatusString + ", " + "SINGLE Click Mode"
        else:
            StatusString = StatusString + ", " + "DOUBLE Click Mode"
    elif ReadingPaused == True:
        StatusString = "Status: PAUSED"



    print(StatusString)

# Exectue functions based on the clicked key
def on_press(key):

    global ReadingPaused
    global Esc_counter
    global F2FirstClick
    global F4FirstClick
    global F2SecondClick
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
                F2SecondClick = time.time()
                if F2SecondClick - F2FirstClick < DoubleClickMaxGap:
                    tripleclick()
                    pya.hotkey('ctrl', 'c')  #copy selected text to the clipboard
                    stopReading() #optional, depends on the desired working mode
                    readThis(clipboard.paste())
                F2FirstClick = F2SecondClick
            else:
                tripleclick()
                pya.hotkey('ctrl', 'c') #copy selected text to the clipboard
                stopReading() #optional, depends on the desired working mode
                readThis(clipboard.paste())

            textGUI('{0}'.format(key))

    if key == keyboard.Key.f4:
        if ReadingPaused == False: #If User haven't click the pause button (F6)
            F4SecondClick = time.time()
            if F4SecondClick - F4FirstClick < DoubleClickMaxGap:
                pya.hotkey('ctrl', 'c')  #copy selected text to the clipboard
                stopReading() #optional, depends on the desired working mode
                readThis(clipboard.paste())
            F4FirstClick = F4SecondClick


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

        if ReadingPaused == False:
            ReadingPaused=True
            toaster.show_toast("F2 Paused", "F2 Paused", threaded=True, icon_path=None, duration=1)

        else:
            ReadingPaused = False
            toaster.show_toast("F2 Activated", "Activated", threaded=True, icon_path=None, duration=1)

        textGUI('{0}'.format(key))

    if key == keyboard.Key.f7: #Decrease Reading Speed
        decSpeed()
        textGUI('{0}'.format(key))


    if key == keyboard.Key.f8: #Increase Reading Speed
        incSpeed()
        textGUI('{0}'.format(key))

    if key == keyboard.Key.f9: #Double / Single click mode

        if RequireDoubleClickForReading == False:
            RequireDoubleClickForReading = True
            toaster.show_toast("F2 DoubleClick Activated", "DoubleClick", threaded=True, icon_path=None, duration=1)

        else:
           RequireDoubleClickForReading = False
           toaster.show_toast("F2 Singleclick Activated", "Singleclick", threaded=True, icon_path=None, duration=1)

# Welcome statement
readThis(OpeningStatement)

# Assiging event to function
listener = keyboard.Listener(on_press=on_press)

# initiating listener
listener.start()

textGUI('')

while Esc_counter < NumberOfEscsNeededForQuitingTheApp:
    time.sleep(.01)
