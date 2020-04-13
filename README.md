# F2-Reader

A simple text to voice app. Simply hover any text and click F2.

This app completely revolutionized the way I read text.

## Getting Started

[F2] - Reads paragraph text positioned under the cursor

[Esc] - Stops Reading

[Esc] x 3 - Exits the App

[F6] - 1st click - Double Click mode (Click F2 twice to read text)

     - 2nd click - Pause the App
     
     - 3rd click - Single click mode (Click F2 once to read text)
     
[F7] - Decrease reading Speed

[F8] - Increase reading Speed

## Things to notice

The way the app works is by sending three mouse click to the GUI, which leads to paragraph selection. Then it copies the paragraph to the clipboard. Then it sends it to Microsoft TTS engine.

That means that the app will overwrite your clipboard. Just keep that in mind.

### Prerequisites

Uses Python 3.8
Works on Windows 10 with the English TTS engine installed (Usually it is installed by default)
Will probably work on other versions of Windows

Install all relevant Dependencies:
import win32com.client as wincl #For the TTS engine
```
import clipboard #For copying to clipboard
import pyautogui as pya #For sending mouse clicks to GUI
import time #For the delay function
from pynput import keyboard #For catching keyboard strokes
from comtypes.client import CreateObject
from os import system, name #For screen clear
from googletrans import Translator #For Translating text
from win10toast import ToastNotifier #For Desktop Notifications
```

## Authors

**Menny Barzilay** - http://Mennyb.com

## License

This project is licensed under the GNU License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

Thanks you for everyone in the amazing open source community.

And thank you google. I couldn't have done this without you.
