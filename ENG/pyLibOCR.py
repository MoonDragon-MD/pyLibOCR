# pyLibOCR by MoonDragon-MD ( https://github.com/MoonDragon-MD )
# Version: 1.0
# 12/2024
import os # To manage files
import PySimpleGUI as sg # Interface
import tkinter as tk # Overlay
import configparser # Manage ini files
from PIL import ImageGrab # To take screenshots
import subprocess # To launch external commands
import json # For the translation
import requests # For translation via LibreTranslate API
import keyboard  # To manage keyboard shortcuts

# Function to load settings from ini file
def load_preferences():
    config = configparser.ConfigParser()
    
    if os.path.isfile('pyLibOCR.ini'):
        # Load preferences if file exists
        config.read('pyLibOCR.ini')
        umi_ocr_path = config['Settings'].get('umi_ocr_path', '')
        source_lang = config['Settings'].get('source_lang', 'en')
        target_lang = config['Settings'].get('target_lang', 'it')

        shortcuts = {
            'ocr_temp': config['Settings'].get('ocr_shortcut_temp', 'alt+c'),
            'ocr_fixed': config['Settings'].get('ocr_shortcut_fixed', 'alt+f'),
            'ocr_shortcut_set_fixed': config['Settings'].get('ocr_shortcut_set_fixed', 'alt+s')
        }

        fixed_area = config['Settings'].get('fixed_area', '0,0,0,0').split(',')
        fixed_area = list(map(int, fixed_area)) if fixed_area else [0, 0, 0, 0]

        return umi_ocr_path, source_lang, target_lang, shortcuts, fixed_area
    else:
        # Create a new ini file with default settings if the file does not exist
        umi_ocr_path = ''
        source_lang = 'en'
        target_lang = 'it'
        shortcuts = {
            'ocr_temp': 'alt+c',
            'ocr_fixed': 'alt+f',
            'ocr_shortcut_set_fixed': 'alt+s'
        }
        fixed_area = [0, 0, 0, 0]

        # Window to select the folder containing Umi-OCR
        layout = [[sg.Text('Select the folder containing Umi-OCR')],
                  [sg.FolderBrowse('Select Folder', key='umi_ocr_folder')],
                  [sg.Button('Confirm')]]
        
        window = sg.Window('Select Umi-OCR Folder', layout)
        event, values = window.read()
        window.close()

        if event == 'Confirm' and values['umi_ocr_folder']:
            umi_ocr_path = values['umi_ocr_folder']
            umi_ocr_path = umi_ocr_path if umi_ocr_path.endswith(os.sep) else umi_ocr_path + os.sep
            save_preferences(umi_ocr_path, source_lang, target_lang, shortcuts, fixed_area)
            return umi_ocr_path, source_lang, target_lang, shortcuts, fixed_area
        else:
            sg.popup_error('Error: You must select a folder for Umi-OCR.')
            return '', 'en', 'it', shortcuts, fixed_area

def save_preferences(umi_ocr_path, source_lang, target_lang, shortcuts, fixed_area=None):
    config = configparser.ConfigParser()
    config['Settings'] = {
        'umi_ocr_path': umi_ocr_path,
        'source_lang': source_lang,
        'target_lang': target_lang,
        'ocr_shortcut_temp': shortcuts['ocr_temp'],
        'ocr_shortcut_fixed': shortcuts['ocr_fixed'],
        'ocr_shortcut_set_fixed': shortcuts['ocr_shortcut_set_fixed'],
        'fixed_area': ','.join(map(str, fixed_area))
    }
    with open('pyLibOCR.ini', 'w') as configfile:
        config.write(configfile)

def ocr_text(umi_ocr_path, area):
    if not os.path.isdir(umi_ocr_path):
        return f"Error: The path {umi_ocr_path} this is not a valid folder."

    umi_ocr_path = umi_ocr_path if umi_ocr_path.endswith(os.sep) else umi_ocr_path + os.sep
    umi_ocr_exe = os.path.join(umi_ocr_path, "Umi-OCR.exe")

    if not os.path.isfile(umi_ocr_exe):
        return f"Error: The file 'Umi-OCR.exe' was not found in {umi_ocr_path}."

    # Delete output file if it exists
    output_file = os.path.join(os.path.dirname(umi_ocr_path), "pyLibOCR.txt")
    if os.path.isfile(output_file):
        os.remove(output_file)

    x, y, width, height = area
    screenshot_command = [
        umi_ocr_exe, '--screenshot', f'screen=0', f'rect={x},{y},{width},{height}', '--output', output_file
    ]

    process = subprocess.Popen(screenshot_command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
    stdout, stderr = process.communicate()

    if process.returncode == 0:
        if os.path.isfile(output_file):
            with open(output_file, 'r', encoding='utf-8') as file:
                ocr_result = file.read()
            return ocr_result
        else:
            return "Error: Output file not created."
    else:
        return f"OCR error: {stderr}"

import requests

def translate_text(text, source_lang, target_lang):
    """Function to perform text translation via API, with handling of special characters and accents."""
    url = "http://localhost:5000/translate"
    headers = {
        "Content-Type": "application/json; charset=UTF-8"
    }

    # Encode text to UTF-8 before sending
    data = {
        "q": text,
        "source": source_lang,
        "target": target_lang
    }

    try:
        # Make a POST request to get the translation
        response = requests.post(url, headers=headers, json=data)

        # Force response encoding to UTF-8
        response.encoding = 'utf-8'

        # Make sure the answer is valid
        response.raise_for_status()

        # Get translated text from JSON response
        translated_text = response.json().get('translatedText', '').strip()

        return translated_text

    except requests.exceptions.RequestException as e:
        # Handling any connection errors
        raise Exception(f"Request error: {e}")
    except ValueError as e:
        # Handling errors in JSON decoding
        raise Exception(f"Error decoding JSON response: {e}")

def create_shortcut_window(shortcuts):
    layout = [
        [sg.Text('Momentary OCR (example: alt+c):')],
        [sg.Input(shortcuts['ocr_temp'], key='ocr_temp')],
        [sg.Text('Fixed OCR (example: alt+f):')],
        [sg.Input(shortcuts['ocr_fixed'], key='ocr_fixed')],
        [sg.Text('Set rectangle (example: alt+s):')],
        [sg.Input(shortcuts['ocr_shortcut_set_fixed'], key='ocr_shortcut_set_fixed')],
        [sg.Button('Save')]
    ]
    window = sg.Window('Configure Shortcuts', layout)
    return window

class MousePositionTracker:
    def __init__(self, root, callback):
        self.root = root
        self.callback = callback
        self.start = None
        self.end = None
        self.rect = None
        self.root.attributes("-fullscreen", True)
        self.canvas = tk.Canvas(self.root, cursor="cross", bg="white")
        self.canvas.pack(fill="both", expand=True)
        self.root.bind("<ButtonPress-1>", self.start_selection)
        self.root.bind("<B1-Motion>", self.track_mouse)
        self.root.bind("<ButtonRelease-1>", self.end_selection)

    def start_selection(self, event):
        self.start = (event.x, event.y)
        self.end = self.start
        self.rect = self.canvas.create_rectangle(self.start[0], self.start[1], self.end[0], self.end[1], outline="red", width=2)

    def track_mouse(self, event):
        self.end = (event.x, event.y)
        self.canvas.coords(self.rect, self.start[0], self.start[1], self.end[0], self.end[1])

    def end_selection(self, event):
        self.end = (event.x, event.y)
        self.canvas.delete(self.rect)
        self.callback(self.start, self.end)
        close_tkinter_window(self.root)  # Force close the Tkinter window

import PySimpleGUI as sg

def create_overlay_window(translated_text, area, overlay_window=None):
    translated_text = translated_text.encode('utf-8').decode('utf-8')
    x, y, width, height = area
    button_width = 50
    height += 30

    # Close the overlay window if it is already open
    if overlay_window is not None and overlay_window.winfo_exists():
        overlay_window.close()  # Closes the existing window
        overlay_window = None  # Reset window

    layout = [
        [sg.Multiline(translated_text, font=("DejaVu Sans", 14), size=(50, 5), autoscroll=True, no_scrollbar=True, reroute_stdout=True, expand_x=True, expand_y=True)],
        [sg.Button("X", size=(3, 1), key="close_overlay", button_color=("white", "red"))]  # Red X button
    ]

    # Create the overlay window
    overlay_window = sg.Window('OCR Overlay', layout,
                               keep_on_top=True, finalize=True,
                               no_titlebar=True, transparent_color='white',
                               alpha_channel=0.8,
                               location=(x, y),
                               size=(width, height))

    # File non trovato!
    while True:
        event, _ = overlay_window.read(timeout=80)  # Timeout to not block the program
        if event in (sg.WINDOW_CLOSED, 'close_overlay', 'Escape'):
            overlay_window.close()  # Closes the overlay window
            overlay_window = None  # Reset window
            break

    return overlay_window
	
# Function to force close the Tkinter window
def close_tkinter_window(root):
    root.quit()   # Esci dal mainloop
    root.destroy()  # Distruggi la finestra

def check_keyboard_shortcuts(shortcuts, window):
    if keyboard.is_pressed(shortcuts['ocr_temp']):
        window.write_event_value('ocr_temp', None)
    elif keyboard.is_pressed(shortcuts['ocr_fixed']):
        window.write_event_value('fixed_ocr', None)
    elif keyboard.is_pressed(shortcuts['ocr_shortcut_set_fixed']):
        window.write_event_value('set_fixed_area', None)

# Main function for GUI
umi_ocr_path, source_lang, target_lang, shortcuts, fixed_area = load_preferences()

layout = [
    [sg.Text('Input Language:', size=(15, 1)), 
     sg.InputCombo(['en', 'de', 'fr', 'es', 'ru', 'zh', 'ja', 'hi', 'it', 'ko', 'nl', 'pl', 'ar', 'tr', 'uk'], key='source_lang', default_value='en'), 
     sg.Text('Output Language:', size=(15, 1)), 
     sg.InputCombo(['it', 'en', 'de', 'fr', 'es', 'ru', 'zh', 'ja', 'hi', 'ko', 'nl', 'pl', 'ar', 'tr', 'uk'], key='target_lang', default_value='it'),
	 sg.Push(), sg.Button('Info')], 
    [sg.Text('OCR Text:', size=(10, 1)), sg.Multiline(size=(41, 10), key='ocr_text')],
    [sg.Button('Re-translate')],
    [sg.Text('Translation:', size=(10, 1)), sg.Multiline(size=(41, 10), key='translated_text')],
    [sg.Button('Momentary OCR', key='ocr_temp'), 
     sg.Button('Set Fixed OCR Rectangle', key='set_fixed_area'), 
     sg.Button('Fixed OCR', key='fixed_ocr')],
    [sg.Checkbox('Overlay', key='overlay'), sg.Button('Shortcuts')]
]

window = sg.Window('pyLibOCR', layout, return_keyboard_events=True)
overlay_window = None

# Main function to manage OCR scanning
while True:
    event, values = window.read(timeout=100)  # Timeout to continuously check the keyboard
    check_keyboard_shortcuts(shortcuts, window)

    if event == sg.WINDOW_CLOSED:
        close_tkinter_window(root)  # Pass 'root' to the function
        # Delete output file on exit
        output_file = os.path.join(os.path.dirname(umi_ocr_path), "pyLibOCR.txt")
        if os.path.isfile(output_file):
            os.remove(output_file)
        break

    if event == 'ocr_temp':
        # Close the overlay window if it is already open
        if overlay_window is not None:
            overlay_window.close()
            overlay_window = None
            close_tkinter_window(root)  # Passa 'root' alla funzione

        # Use tkinter to select the area only when needed
        root = tk.Tk()
        root.state('normal')  # Window not full screen
        root.attributes('-alpha', 0.5)
        tracker = MousePositionTracker(root, lambda start, end: print("Selected area:", start, end))
        root.mainloop()

        if tracker.start and tracker.end:
            area = (tracker.start[0], tracker.start[1], tracker.end[0] - tracker.start[0], tracker.end[1] - tracker.start[1])
            ocr_result = ocr_text(umi_ocr_path, area)
            window['ocr_text'].update(ocr_result)
            translated_text = translate_text(ocr_result, values['source_lang'], values['target_lang'])
            window['translated_text'].update(translated_text)

            # Check if the checkbox is selected
            if values['overlay']:  # The overlay option is selected
                overlay_window = create_overlay_window(translated_text, area)

    elif event == 'fixed_ocr':
        # Close the overlay window if it is already open
        if overlay_window is not None:
            overlay_window.close()
            close_tkinter_window(self.root)  # Force close the Tkinter window
            overlay_window = None  # Resetting the variable after closing it

        area = fixed_area
        if all(coord == 0 for coord in area):
            sg.popup_error("Error: No fixed area set. Please set fixed area first.")
            continue
        ocr_result = ocr_text(umi_ocr_path, area)
        window['ocr_text'].update(ocr_result)
        translated_text = translate_text(ocr_result, values['source_lang'], values['target_lang'])
        window['translated_text'].update(translated_text)

        # Show overlay if option is selected
        if values['overlay']:
            overlay_window = create_overlay_window(translated_text, area)

    elif event == 'set_fixed_area':
        # Close the overlay window if it is already open
        if overlay_window:
            overlay_window.close()
            close_tkinter_window(self.root)  # Force close the Tkinter window
            overlay_window = None  # We reset the variable to prevent the creation of a new window

        root = tk.Tk()
        root.state('zoomed')
        root.attributes('-alpha', 0.5)
        root.title('Select the area with the mouse')
        tracker = MousePositionTracker(root, lambda start, end: print("Selected area:", start, end))
        root.mainloop()

        if tracker.start and tracker.end:
            fixed_area = (tracker.start[0], tracker.start[1], tracker.end[0] - tracker.start[0], tracker.end[1] - tracker.start[1])
            save_preferences(umi_ocr_path, values['source_lang'], values['target_lang'], shortcuts, fixed_area)

    if event == 'overlay':
        if values['overlay']:  # If the overlay checkbox is selected
            if overlay_window is None or not overlay_window.winfo_exists():  # Check if the window does not exist
                print("Creating overlay window...")
                overlay_window = create_overlay_window(values['translated_text'], fixed_area)
        else:  # If the checkbox is unchecked
            if overlay_window:  # If the window is open
                print("Closing overlay window...")
                overlay_window.close()
                overlay_window = None

    elif event == 'overlay' and not values['overlay']:
        # If the checkbox is unchecked, we close the overlay window
        print("Closing overlay window...")
        if overlay_window is not None:
            overlay_window.close()
            overlay_window = None  # We reset the variable to prevent the creation of a new window

    if event == 'Shortcuts':
        config_window = create_shortcut_window(shortcuts)
        while True:
            event, values = config_window.read()
            if event == sg.WINDOW_CLOSED:
                break
            if event == 'Save':
                save_preferences(umi_ocr_path, values['source_lang'], values['target_lang'], shortcuts, fixed_area)
        config_window.close()

    if event == 'Ri-traduci':
        ocr_text_input = values['ocr_text']
        translated_text = translate_text(ocr_text_input, values['source_lang'], values['target_lang'])
        window['translated_text'].update(translated_text)
		
    # Check if the user clicked the Info button
    if event == 'Info':
        # information window
        info_layout = [
            [sg.Text('pyLibOCR')],
            [sg.Text('Version 1.0')],
			[sg.Text('Created by MoonDragon-MD')],
            [sg.Text('Site: http://moondragon.ilbello.com/')],
            [sg.Text('Works with Umi-OCR and LibreTranslate')],
            [sg.Button('Close')]
        ]
        
        # Create the information window
        info_window = sg.Window('Information', info_layout)
        
        # Loop to handle information window events
        while True:
            info_event, info_values = info_window.read()
            
            # Check if the user has closed the information window
            if info_event == sg.WINDOW_CLOSED or info_event == 'Close':
                break
        
        # Close the information window
        info_window.close()

window.close()
