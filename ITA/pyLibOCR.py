# pyLibOCR by MoonDragon-MD ( https://github.com/MoonDragon-MD )
# Version: 1.0
# 12/2024
import os # Per gestire i file
import PySimpleGUI as sg # Interfaccia
import tkinter as tk # Sovrimpressione
import configparser # Gestire file ini
from PIL import ImageGrab # Per acquisire screenshot
import subprocess # Per lanciare comandi esterni
import json # Per la traduzione
import requests # Per la traduzione via API LibreTranslate
import keyboard  # Per gestire le scorciatoie da tastiera

# Funzione per caricare le impostazioni dal file ini
def load_preferences():
    config = configparser.ConfigParser()
    
    if os.path.isfile('pyLibOCR.ini'):
        # Carica le preferenze se il file esiste
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
        # Crea un nuovo file ini con impostazioni predefinite se il file non esiste
        umi_ocr_path = ''
        source_lang = 'en'
        target_lang = 'it'
        shortcuts = {
            'ocr_temp': 'alt+c',
            'ocr_fixed': 'alt+f',
            'ocr_shortcut_set_fixed': 'alt+s'
        }
        fixed_area = [0, 0, 0, 0]

        # Finestra per selezionare la cartella contenente Umi-OCR
        layout = [[sg.Text('Seleziona la cartella contenente Umi-OCR')],
                  [sg.FolderBrowse('Seleziona Cartella', key='umi_ocr_folder')],
                  [sg.Button('Conferma')]]
        
        window = sg.Window('Seleziona Cartella Umi-OCR', layout)
        event, values = window.read()
        window.close()

        if event == 'Conferma' and values['umi_ocr_folder']:
            umi_ocr_path = values['umi_ocr_folder']
            umi_ocr_path = umi_ocr_path if umi_ocr_path.endswith(os.sep) else umi_ocr_path + os.sep
            save_preferences(umi_ocr_path, source_lang, target_lang, shortcuts, fixed_area)
            return umi_ocr_path, source_lang, target_lang, shortcuts, fixed_area
        else:
            sg.popup_error('Errore: devi selezionare una cartella per Umi-OCR.')
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
        return f"Errore: il percorso {umi_ocr_path} non è una cartella valida."

    umi_ocr_path = umi_ocr_path if umi_ocr_path.endswith(os.sep) else umi_ocr_path + os.sep
    umi_ocr_exe = os.path.join(umi_ocr_path, "Umi-OCR.exe")

    if not os.path.isfile(umi_ocr_exe):
        return f"Errore: il file 'Umi-OCR.exe' non è stato trovato in {umi_ocr_path}."

    # Elimina il file di output se esiste
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
            return "Errore: il file di output non è stato creato."
    else:
        return f"Errore nell'OCR: {stderr}"

import requests

def translate_text(text, source_lang, target_lang):
    """Funzione per eseguire la traduzione del testo tramite l'API, con gestione dei caratteri speciali e accenti."""
    url = "http://localhost:5000/translate"
    headers = {
        "Content-Type": "application/json; charset=UTF-8"
    }

    # Codifica il testo in UTF-8 prima di inviarlo
    data = {
        "q": text,
        "source": source_lang,
        "target": target_lang
    }

    try:
        # Fai la richiesta POST per ottenere la traduzione
        response = requests.post(url, headers=headers, json=data)

        # Forza la codifica della risposta in UTF-8
        response.encoding = 'utf-8'

        # Assicurati che la risposta sia valida
        response.raise_for_status()

        # Prendi il testo tradotto dalla risposta JSON
        translated_text = response.json().get('translatedText', '').strip()

        return translated_text

    except requests.exceptions.RequestException as e:
        # Gestione di eventuali errori di connessione
        raise Exception(f"Errore nella richiesta: {e}")
    except ValueError as e:
        # Gestione di errori nella decodifica del JSON
        raise Exception(f"Errore nella decodifica della risposta JSON: {e}")

def create_shortcut_window(shortcuts):
    layout = [
        [sg.Text('OCR Momentaneo (esempio: alt+c):')],
        [sg.Input(shortcuts['ocr_temp'], key='ocr_temp')],
        [sg.Text('OCR Fisso (esempio: alt+f):')],
        [sg.Input(shortcuts['ocr_fixed'], key='ocr_fixed')],
        [sg.Text('Imposta rettangolo (esempio: alt+s):')],
        [sg.Input(shortcuts['ocr_shortcut_set_fixed'], key='ocr_shortcut_set_fixed')],
        [sg.Button('Salva')]
    ]
    window = sg.Window('Configura Scorciatoie', layout)
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
        close_tkinter_window(self.root)  # Forza la chiusura della finestra Tkinter

import PySimpleGUI as sg

def create_overlay_window(translated_text, area, overlay_window=None):
    translated_text = translated_text.encode('utf-8').decode('utf-8')
    x, y, width, height = area
    button_width = 50
    height += 30

    # Chiudi la finestra di sovrimpressione se è già aperta
    if overlay_window is not None and overlay_window.winfo_exists():
        overlay_window.close()  # Chiude la finestra esistente
        overlay_window = None  # Reset della finestra

    layout = [
        [sg.Multiline(translated_text, font=("DejaVu Sans", 14), size=(50, 5), autoscroll=True, no_scrollbar=True, reroute_stdout=True, expand_x=True, expand_y=True)],
        [sg.Button("X", size=(3, 1), key="close_overlay", button_color=("white", "red"))]  # Pulsante X rosso
    ]

    # Crea la finestra di sovrimpressione
    overlay_window = sg.Window('Sovrimpressione OCR', layout,
                               keep_on_top=True, finalize=True,
                               no_titlebar=True, transparent_color='white',
                               alpha_channel=0.8,
                               location=(x, y),
                               size=(width, height))

    # Ciclo per gestire gli eventi della finestra
    while True:
        event, _ = overlay_window.read(timeout=80)  # Timeout per non bloccare il programma
        if event in (sg.WINDOW_CLOSED, 'close_overlay', 'Escape'):
            overlay_window.close()  # Chiude la finestra di sovrimpressione
            overlay_window = None  # Reset della finestra
            break

    return overlay_window
	
# Funzione per chiudere forzatamente la finestra Tkinter
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

# Funzione principale per la GUI
umi_ocr_path, source_lang, target_lang, shortcuts, fixed_area = load_preferences()

layout = [
    [sg.Text('Lingua in Ingresso:', size=(15, 1)), 
     sg.InputCombo(['en', 'de', 'fr', 'es', 'ru', 'zh', 'ja', 'hi', 'it', 'ko', 'nl', 'pl', 'ar', 'tr', 'uk'], key='source_lang', default_value='en'), 
     sg.Text('Lingua in Uscita:', size=(13, 1)), 
     sg.InputCombo(['it', 'en', 'de', 'fr', 'es', 'ru', 'zh', 'ja', 'hi', 'ko', 'nl', 'pl', 'ar', 'tr', 'uk'], key='target_lang', default_value='it'),
	 sg.Push(), sg.Button('Info')], 
    [sg.Text('Testo OCR:', size=(10, 1)), sg.Multiline(size=(41, 10), key='ocr_text')],
    [sg.Button('Ri-traduci')],
    [sg.Text('Traduzione:', size=(10, 1)), sg.Multiline(size=(41, 10), key='translated_text')],
    [sg.Button('OCR Momentaneo', key='ocr_temp'), 
     sg.Button('Imposta Rettangolo OCR Fisso', key='set_fixed_area'), 
     sg.Button('OCR Fisso', key='fixed_ocr')],
    [sg.Checkbox('Sovrimpressione', key='overlay'), sg.Button('Scorciatoie')]
]

window = sg.Window('pyLibOCR', layout, return_keyboard_events=True)
overlay_window = None

# Funzione principale per gestire la scansione OCR
while True:
    event, values = window.read(timeout=100)  # Timeout per controllare continuamente la tastiera
    check_keyboard_shortcuts(shortcuts, window)

    if event == sg.WINDOW_CLOSED:
        close_tkinter_window(root)  # Passa 'root' alla funzione
        # Elimina il file di output all'uscita
        output_file = os.path.join(os.path.dirname(umi_ocr_path), "pyLibOCR.txt")
        if os.path.isfile(output_file):
            os.remove(output_file)
        break

    if event == 'ocr_temp':
        # Chiudi la finestra di sovrimpressione se già aperta
        if overlay_window is not None:
            overlay_window.close()
            overlay_window = None
            close_tkinter_window(root)  # Passa 'root' alla funzione

        # Usa tkinter per selezionare l'area solo quando necessario
        root = tk.Tk()
        root.state('normal')  # Finestra non ingrandita (non a tutto schermo)
        root.attributes('-alpha', 0.5)
        tracker = MousePositionTracker(root, lambda start, end: print("Area selezionata:", start, end))
        root.mainloop()

        if tracker.start and tracker.end:
            area = (tracker.start[0], tracker.start[1], tracker.end[0] - tracker.start[0], tracker.end[1] - tracker.start[1])
            ocr_result = ocr_text(umi_ocr_path, area)
            window['ocr_text'].update(ocr_result)
            translated_text = translate_text(ocr_result, values['source_lang'], values['target_lang'])
            window['translated_text'].update(translated_text)

            # Verifica se la checkbox è selezionata
            if values['overlay']:  # L'opzione overlay è selezionata
                overlay_window = create_overlay_window(translated_text, area)

    elif event == 'fixed_ocr':
        # Chiudi la finestra di sovrimpressione se già aperta
        if overlay_window is not None:
            overlay_window.close()
            close_tkinter_window(self.root)  # Forza la chiusura della finestra Tkinter
            overlay_window = None  # Reset della variabile dopo averla chiusa

        area = fixed_area
        if all(coord == 0 for coord in area):
            sg.popup_error("Errore: Nessuna area fissa impostata. Imposta prima l'area fissa.")
            continue
        ocr_result = ocr_text(umi_ocr_path, area)
        window['ocr_text'].update(ocr_result)
        translated_text = translate_text(ocr_result, values['source_lang'], values['target_lang'])
        window['translated_text'].update(translated_text)

        # Mostra la sovrimpressione se l'opzione è selezionata
        if values['overlay']:
            overlay_window = create_overlay_window(translated_text, area)

    elif event == 'set_fixed_area':
        # Chiudi la finestra di sovrimpressione se già aperta
        if overlay_window:
            overlay_window.close()
            close_tkinter_window(self.root)  # Forza la chiusura della finestra Tkinter
            overlay_window = None  # Azzeriamo la variabile per prevenire la creazione di una nuova finestra

        root = tk.Tk()
        root.state('zoomed')
        root.attributes('-alpha', 0.5)
        root.title('Seleziona l\'area con il mouse')
        tracker = MousePositionTracker(root, lambda start, end: print("Selected area:", start, end))
        root.mainloop()

        if tracker.start and tracker.end:
            fixed_area = (tracker.start[0], tracker.start[1], tracker.end[0] - tracker.start[0], tracker.end[1] - tracker.start[1])
            save_preferences(umi_ocr_path, values['source_lang'], values['target_lang'], shortcuts, fixed_area)

    if event == 'overlay':
        if values['overlay']:  # Se la checkbox overlay è selezionata
            if overlay_window is None or not overlay_window.winfo_exists():  # Verifica se la finestra non esiste
                print("Creando finestra di sovrimpressione...")
                overlay_window = create_overlay_window(values['translated_text'], fixed_area)
        else:  # Se la checkbox è deselezionata
            if overlay_window:  # Se la finestra è aperta
                print("Chiudendo finestra di sovrimpressione...")
                overlay_window.close()
                overlay_window = None

    elif event == 'overlay' and not values['overlay']:
        # Se il checkbox è deselezionato, chiudiamo la finestra di sovrimpressione
        print("Chiudendo finestra di sovrimpressione...")
        if overlay_window is not None:
            overlay_window.close()
            overlay_window = None  # Azzeriamo la variabile per prevenire la creazione di una nuova finestra

    if event == 'Scorciatoie':
        config_window = create_shortcut_window(shortcuts)
        while True:
            event, values = config_window.read()
            if event == sg.WINDOW_CLOSED:
                break
            if event == 'Salva':
                save_preferences(umi_ocr_path, values['source_lang'], values['target_lang'], shortcuts, fixed_area)
        config_window.close()

    if event == 'Ri-traduci':
        ocr_text_input = values['ocr_text']
        translated_text = translate_text(ocr_text_input, values['source_lang'], values['target_lang'])
        window['translated_text'].update(translated_text)
		
    # Controlla se l'utente ha cliccato sul pulsante Info
    if event == 'Info':
        # finestra delle informazioni
        info_layout = [
            [sg.Text('pyLibOCR')],
            [sg.Text('Versione 1.0')],
			[sg.Text('Creata da MoonDragon-MD')],
            [sg.Text('Sito: http://moondragon.ilbello.com/')],
            [sg.Text('Funziona con Umi-OCR e LibreTranslate')],
            [sg.Button('Chiudi')]
        ]
        
        # Crea la finestra delle informazioni
        info_window = sg.Window('Informazioni', info_layout)
        
        # Ciclo per gestire gli eventi della finestra delle informazioni
        while True:
            info_event, info_values = info_window.read()
            
            # Controlla se l'utente ha chiuso la finestra delle informazioni
            if info_event == sg.WINDOW_CLOSED or info_event == 'Chiudi':
                break
        
        # Chiudi la finestra delle informazioni
        info_window.close()

window.close()