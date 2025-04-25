# PDF-to-Audio-Converter
PDF to Audio Converter
A Python application with a graphical user interface (GUI) built using customtkinter. It converts PDF files to audio by extracting text, supporting translation to multiple languages, generating audio in MP3/WAV/OGG formats, and optionally sending audio files via email.
Team
This project was developed by a team of five members:

A Hemanth Kumar [https://github.com/HEMANTH-A-7](url)
Ch Seetharama Kartheek [https://github.com/srk-ch](url) 
G Sai Karthikeya [](url)
R Pujith [https://github.com/pujithravi](url)

Features

Extract text from PDFs using PyPDF2 or pytesseract (OCR for scanned PDFs).
Translate text to various languages using googletrans.
Convert text to audio using gTTS or pyttsx3.
Save audio in MP3, WAV, or OGG formats.
Send audio files via email using Gmail SMTP.
Editable text display with support for multiple PDFs and navigation.

Prerequisites

We did it on Python 3.11(make sure having it not working on 3.13 at present(25/04/25))
Git installed
Poppler (for pdf2image):
Windows: Download from poppler-utils and add to PATH.
Mac: brew install poppler
Linux: sudo apt-get install poppler-utils


Tesseract OCR (for pytesseract):
Windows: Download from tesseract-ocr and add to PATH.
Mac: brew install tesseract
Linux: sudo apt-get install tesseract-ocr


FFmpeg (for WAV/OGG conversion):
Install via brew install ffmpeg (Mac), sudo apt-get install ffmpeg (Linux), or download for Windows and add to PATH.



Installation

Clone the repository:git clone https://github.com/yourusername/pdf-to-audio-converter.git
cd pdf-to-audio-converter


Create and activate a virtual environment:python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate


Install dependencies:pip install -r requirements.txt


Ensure mountain_gradient.jpg is in the project root (or update the path in pdf_to_audio.py).

Usage

Run the application:python pdf_to_audio.py


Select PDF files using the "Browse" button.
Edit extracted text, choose a language, audio format, and voice.
Click "Convert to Audio" to generate audio files.
Optionally, enter a recipient’s email to send audio files.


Notes

Ensure poppler, tesseract-ocr, and ffmpeg are installed and accessible in your system PATH.
The audios/ folder is excluded via .gitignore as it contains generated audio files.
Text extraction may fail for encrypted or corrupted PDFs.

License
This project is licensed under the MIT License. See the LICENSE file for details.
Copyright
© 2025 A Hemanth Kumar,G Sai Karthikeya, Ch Seetharama Kartheek,R Pujith. All rights reserved.
