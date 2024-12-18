# pyLibOCR
A GUI for Umi-OCR and translation with LibreTranslate complete with overlay, all strictly offline
### Requirements:
- Python (I used 3.6.8)
- pip install PySimpleGUI Pillow requests keyboard (Pillow is currently not used but is planned)
  If you want the unsubscribed version of PySimpleGUI, use this command: python -m pip install PySimpleGUI==4.60.5.0
- [Umi-OCR]( https://github.com/hiroi-sora/Umi-OCR)
- - ##### LibreTranslate (Argos) [localhost:5000]
[LibreTranslate on Docker](https://hub.docker.com/r/libretranslate/libretranslate)
Or there are other ways:
[LibreTranslate Github](https://github.com/LibreTranslate/LibreTranslate)
### Installation:
Just copy the two files, I put them in the Umi-OCR installation folder for convenience.
### Usage:
Start Umi-OCR and LibreTranslate, then make two clicks on the *.bat I gave you and the interface will start.
You'll be able to set up shortcuts to use it with the keyboard, useful along with the overlay option, the important thing is that you don't minimize pyLibOCR but it can stay under the other programs.
You can also use it just for translation with the re-translate button.
It is the first version there is still some little problem now and then with the overlay but otherwise it is fine.
