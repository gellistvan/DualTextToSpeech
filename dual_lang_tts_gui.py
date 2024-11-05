import threading
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import edge_tts
import asyncio
from pydub import AudioSegment
import pandas as pd
import os

AudioSegment.ffmpeg = ".\\ffmpeg.exe"
AudioSegment.ffprobe = ".\\ffprobe.exe"

# Available languages and corresponding Edge TTS voice options
LANGUAGE_VOICE_MAP = {
    "Hungarian": "hu-HU",
    "English": "en-US",
    "Spanish": "es-ES",
    "German": "de-DE",
    "Swedish": "sv-SE",
    "Russian": "ru-RU",
}

swedish_audio_file = f"swedish_temp.mp3"
hungarian_audio_file = f"hungarian_temp.mp3"
            
# Initialize main application window
root = tk.Tk()
root.title("Text to Speech Converter")

# Variables for file path and selected languages and voices
file_path = tk.StringVar()
first_language = tk.StringVar()
second_language = tk.StringVar()
first_voice = tk.StringVar()
second_voice = tk.StringVar()

# File chooser
def choose_file():
    file = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("Excel files", "*.xls *.xlsx")])
    file_path.set(file)


            
# Populate voice dropdown based on selected language
async def update_voice_options(language_var, voice_var, dropdown):
    language_code = LANGUAGE_VOICE_MAP.get(language_var.get(), "")
    voices = await edge_tts.list_voices()  # Retrieve all voices
    # Filter voices by language code and extract only the ShortName
    voice_options = [voice["ShortName"] for voice in voices if voice["ShortName"].startswith(language_code)]
    dropdown["values"] = voice_options
    if voice_options:
        voice_var.set(voice_options[0])  # Set default voice to the first available option

# Function to create MP3 from phrases in a text file
async def create_mp3_from_txt(filename):
    phrases = read_phrases_from_txt(filename)
    combined_audio = AudioSegment.empty()
    global cancel_requested

    
    for swedish, hungarian in phrases:
        swedish_audio_file = "swedish_temp.mp3"
        hungarian_audio_file = "hungarian_temp.mp3"
        
        await synthesize_text(swedish, swedish_audio_file, first_voice.get())
        await synthesize_text(hungarian, hungarian_audio_file, second_voice.get())
        
        swedish_audio = AudioSegment.from_file(swedish_audio_file)
        hungarian_audio = AudioSegment.from_file(hungarian_audio_file)
        
        # Calculate pause length based on Swedish phrase length
        pause_len = max(3000, int(3000 + 1.25 * (len(swedish) // 10) * 1000))
        combined_audio += swedish_audio + add_silence(pause_len) + hungarian_audio + add_silence(2000)
        
        os.remove(swedish_audio_file)
        os.remove(hungarian_audio_file)
        if cancel_requested:
            return
    
    # Export combined audio as a single MP3
    output_filename = filename.rsplit(".", 1)[0] + ".mp3"
    combined_audio.export(output_filename, format="mp3")
    print(f"MP3 file created: {output_filename}")

# Function to create MP3 for each sheet in the Excel file
async def create_mp3_from_excel(filename):
    phrases_by_sheet = read_phrases_from_excel(filename)
    global cancel_requested
    for sheet_index, phrases in phrases_by_sheet.items():
        combined_audio = AudioSegment.empty()
        
        for swedish, hungarian in phrases:            
            await synthesize_text(swedish, swedish_audio_file, first_voice.get())
            await synthesize_text(hungarian, hungarian_audio_file, second_voice.get())
            
            swedish_audio = AudioSegment.from_file(swedish_audio_file)
            hungarian_audio = AudioSegment.from_file(hungarian_audio_file)
            
            # Calculate pause length based on Swedish phrase length
            pause_len = max(3000, int(3000 + 1.25 * (len(swedish) // 10) * 1000))
            combined_audio += swedish_audio + add_silence(pause_len) + hungarian_audio + add_silence(2000)
            
            os.remove(swedish_audio_file)
            os.remove(hungarian_audio_file)
            
            if cancel_requested:
                return
        
        output_filename = f"{filename.rsplit('.', 1)[0]}_{sheet_index}.mp3"
        combined_audio.export(output_filename, format="mp3")
        print(f"MP3 file created for sheet {sheet_index}: {output_filename}")

# Async function to fetch voices for both language selections
async def fetch_voices():
    await update_voice_options(first_language, first_voice, first_voice_dropdown)
    await update_voice_options(second_language, second_voice, second_voice_dropdown)

# Start TTS processing in a separate thread
def start_processing():
    global cancel_requested
    cancel_requested = False
    threading.Thread(target=lambda: asyncio.run(process_file())).start()

# Cancel processing
def cancel_processing():
    global cancel_requested
    cancel_requested = True
    messagebox.showinfo("Canceled", "Processing has been canceled.")

# Process file, checking for cancellation
async def process_file():
    file = file_path.get()
    if not file:
        messagebox.showerror("Error", "Please select a file.")
        return

    if cancel_requested:
        return

    # Use the appropriate processing function based on file type
    if file.endswith(".txt"):
        await create_mp3_from_txt(file)
    elif file.endswith(".xls") or file.endswith(".xlsx"):
        await create_mp3_from_excel(file)
    else:
        messagebox.showerror("Error", "Unsupported file type. Please choose a .txt, .xls, or .xlsx file.")
        return
    messagebox.showinfo("Success", "MP3 files created successfully.")

# Implement create_mp3_from_txt and create_mp3_from_excel functions (same as before)


# Functions for synthesis and processing (similar to the original script)
async def synthesize_text(text, output_file, voice):
    tts = edge_tts.Communicate(text, voice)
    await tts.save(output_file)

def add_silence(duration_ms):
    return AudioSegment.silent(duration=duration_ms)


# Function to read phrases from a text file
def read_phrases_from_txt(filename):
    phrases = []
    with open(filename, "r", encoding="utf-8") as file:
        for line in file:
            if "\t" in line:
                hungarian, swedish = line.strip().split("\t")
                phrases.append((swedish, hungarian))
    return phrases

# Function to read Excel file and return phrases as lists
def read_phrases_from_excel(filename):
    sheets_data = {}
    excel_data = pd.ExcelFile(filename)
    
    for index, sheet_name in enumerate(excel_data.sheet_names):
        # Read each sheet, and get the first two columns only, dropping empty rows
        df = excel_data.parse(sheet_name, usecols=[0, 1]).dropna()
        sheets_data[index] = df.values.tolist()  # Convert DataFrame rows to list of tuples
    
    return sheets_data


# Layout: GUI Elements
file_label = tk.Label(root, text="Select File:")
file_label.grid(row=0, column=0, padx=5, pady=5)
file_button = tk.Button(root, text="Browse", command=choose_file)
file_button.grid(row=0, column=1, padx=5, pady=5)
file_entry = tk.Entry(root, textvariable=file_path, width=40)
file_entry.grid(row=0, column=2, padx=5, pady=5)

# First language and voice selection
first_lang_label = tk.Label(root, text="First Language:")
first_lang_label.grid(row=1, column=0, padx=5, pady=5)
first_lang_dropdown = ttk.Combobox(root, textvariable=first_language, values=list(LANGUAGE_VOICE_MAP.keys()))
first_lang_dropdown.grid(row=1, column=1, padx=5, pady=5)
first_voice_dropdown = ttk.Combobox(root, textvariable=first_voice)
first_voice_dropdown.grid(row=1, column=2, padx=5, pady=5)

# Second language and voice selection
second_lang_label = tk.Label(root, text="Second Language:")
second_lang_label.grid(row=2, column=0, padx=5, pady=5)
second_lang_dropdown = ttk.Combobox(root, textvariable=second_language, values=list(LANGUAGE_VOICE_MAP.keys()))
second_lang_dropdown.grid(row=2, column=1, padx=5, pady=5)
second_voice_dropdown = ttk.Combobox(root, textvariable=second_voice)
second_voice_dropdown.grid(row=2, column=2, padx=5, pady=5)

# Button actions for Process and Cancel
process_button = tk.Button(root, text="Process", command=start_processing)
process_button.grid(row=3, column=0, columnspan=2, padx=5, pady=10)

cancel_button = tk.Button(root, text="Cancel", command=cancel_processing)
cancel_button.grid(row=3, column=2, padx=5, pady=10)

# Event bindings to update voices when language is selected
first_lang_dropdown.bind("<<ComboboxSelected>>", lambda e: asyncio.run(update_voice_options(first_language, first_voice, first_voice_dropdown)))
second_lang_dropdown.bind("<<ComboboxSelected>>", lambda e: asyncio.run(update_voice_options(second_language, second_voice, second_voice_dropdown)))

root.mainloop()