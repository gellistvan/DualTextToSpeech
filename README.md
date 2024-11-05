# DualTextToSpeech

DualTextToSpeech is a GUI-based application that converts phrases in a text (`.txt`) or Excel (`.xls` or `.xlsx`) file into MP3 audio files, with support for two languages. Users can select the languages and voices for each language using Edge TTS.
![image](https://github.com/user-attachments/assets/ae46c0f6-4534-4c02-9474-3ec9f04f4b7b)

## Prerequisites

1. **Python**: Make sure Python 3.8+ is installed on your system.
2. **FFmpeg**: Required by `pydub` for audio processing.

## Setup Instructions

1. **Download FFmpeg**:
   - Visit the official [FFmpeg download page](https://ffmpeg.org/download.html).
   - Download the appropriate version for your operating system.
   - Extract `ffmpeg.exe` and `ffprobe.exe` and place them next to `dual_lang_tts_gui.py`.

2. **Set Up Virtual Environment**:
   - Open a terminal or command prompt in the project folder.
   - Create a virtual environment:
     ```bash
     python -m venv venv
     ```
   - Activate the virtual environment:
     - **Windows**:
       ```bash
       venv\Scripts\activate
       ```
     - **Mac/Linux**:
       ```bash
       source venv/bin/activate
       ```

3. **Install Dependencies**:
   - Install the required packages from `requirements.txt`:
     ```bash
     pip install -r requirements.txt
     ```

## Usage

1. Run the script:
   ```bash
   python dual_lang_tts_gui.py
