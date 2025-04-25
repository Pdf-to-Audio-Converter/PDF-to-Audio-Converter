import os



import threading
import tkinter.filedialog
import tkinter.messagebox
import customtkinter as ctk
from PIL import Image, ImageTk
import traceback
from email.message import EmailMessage
import smtplib
from googletrans import Translator, LANGUAGES
import pyttsx3
from PyPDF2 import PdfReader    
import time
import pytesseract
from pdf2image import convert_from_path
import subprocess

# Configure appearance
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

class PDFToAudioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Audio Converter")
        self.root.geometry("1000x700")
        self.root.minsize(900, 600)
        
        # Email credentials
        self.sender_email = "neocolab07@gmail.com"
        self.sender_password = "xoftzmvtcoltmuxw"
        
        # Initialize variables
        self.file_paths = []
        self.current_file_index = -1
        self.extracted_texts = []
        self.current_width = 1000
        self.current_height = 700
        
        # Setup colors
        self.colors = {
            "bg": "#ffffff",
            "accent": "#7b1fa2",
            "accent_light": "#ae52d4",
            "accent_gradient": "#ff4081",
            "button": "#7b1fa2",
            "button_hover": "#6a1b9a",
            "text": "#333333",
            "secondary_text": "#757575",
            "widget_bg": "#f5f5f5",
            "border": "#e0e0e0",
            "input_bg": "#ffffff"
        }

        # Language-to-voice mapping
        self.voice_language_map = {
            'en': ['english', 'us', 'uk', 'british', 'american'],
            'es': ['spanish', 'español'],
            'fr': ['french', 'français'],
            'de': ['german', 'deutsch'],
            'zh-cn': ['chinese', 'mandarin', 'zh'],
            'ja': ['japanese', 'jp'],
            'ru': ['russian', 'ru'],
        }

        # Setup UI
        self.setup_split_layout()
        self.create_widgets()
        
        # Bind events
        self.root.bind('<Configure>', self.on_window_resize)

    def setup_split_layout(self):
        self.main_frame = ctk.CTkFrame(self.root, fg_color=self.colors["bg"], corner_radius=15)
        self.main_frame.pack(fill="both", expand=True)
        
        self.left_panel = ctk.CTkFrame(self.main_frame, fg_color=self.colors["accent"], corner_radius=15, width=500)
        self.left_panel.pack(side="left", fill="y")
        self.left_panel.pack_propagate(False)
        
        try:
            self.gradient_image = Image.open("mountain_gradient.jpg")
            self.gradient_img = ctk.CTkImage(self.gradient_image, size=(500, 700))
            self.gradient_label = ctk.CTkLabel(self.left_panel, text="", image=self.gradient_img)
            self.gradient_label.place(relx=0.5, rely=0.5, anchor="center")
        except Exception as e:
            print(f"Gradient image not found, using solid color: {e}")
            self.left_panel.configure(fg_color=self.colors["accent"])
        
        self.right_panel = ctk.CTkScrollableFrame(self.main_frame, fg_color=self.colors["bg"], corner_radius=15)
        self.right_panel.pack(side="right", fill="both", expand=True)

    def on_window_resize(self, event):
        if event.widget == self.root:
            try:
                self.current_width = event.width
                self.current_height = event.height
                if hasattr(self, 'gradient_image') and self.gradient_image:
                    self.gradient_img.configure(size=(500, self.current_height))
                self.reposition_widgets()
            except Exception as e:
                print(f"Error during resize: {e}")

    def create_widgets(self):
        try:
            self.title_label = ctk.CTkLabel(
                self.right_panel,
                text="PDF to Audio Converter",
                font=("Segoe UI", 28, "bold"),
                text_color=self.colors["accent"]
            )
            self.title_label.pack(pady=(20, 20), anchor="center")

            self.content_frame = ctk.CTkFrame(
                self.right_panel,
                fg_color=self.colors["bg"],
                corner_radius=15,
                width=450
            )
            self.content_frame.pack(pady=10, padx=20, fill="both", expand=True)

            # File selection
            self.file_frame = self.create_form_group(self.content_frame, "Select PDF Files")
            self.browse_button = ctk.CTkButton(
                self.file_frame,
                text="Browse",
                command=self.select_pdfs,
                fg_color=self.colors["button"],
                hover_color=self.colors["button_hover"],
                corner_radius=6,
                width=80,
                height=30
            )
            self.browse_button.pack(pady=(5, 0), anchor="center")
            
            self.file_display = ctk.CTkEntry(
                self.file_frame,
                width=300,
                height=30,
                fg_color=self.colors["input_bg"],
                text_color=self.colors["text"],
                border_color=self.colors["border"],
                placeholder_text="No files selected"
            )
            self.file_display.pack(pady=(5, 5), fill="x", anchor="center")
            self.file_display.configure(state="disabled")

            # Navigation buttons
            self.nav_frame = ctk.CTkFrame(self.file_frame, fg_color="transparent")
            self.nav_frame.pack(pady=5)
            self.prev_button = ctk.CTkButton(
                self.nav_frame,
                text="Previous",
                command=self.show_prev_file,
                fg_color=self.colors["button"],
                hover_color=self.colors["button_hover"],
                corner_radius=6,
                width=80,
                height=30
            )
            self.prev_button.pack(side="left", padx=5)
            self.next_button = ctk.CTkButton(
                self.nav_frame,
                text="Next",
                command=self.show_next_file,
                fg_color=self.colors["button"],
                hover_color=self.colors["button_hover"],
                corner_radius=6,
                width=80,
                height=30
            )
            self.next_button.pack(side="right", padx=5)

            self.text_frame = ctk.CTkFrame(self.file_frame, fg_color=self.colors["bg"], corner_radius=6)
            self.text_frame.pack(pady=(5, 10), fill="both", expand=True)
            
            self.text_label_frame = ctk.CTkFrame(self.text_frame, fg_color="transparent")
            self.text_label_frame.pack(fill="x", padx=5, pady=(0, 5))
            self.text_label = ctk.CTkLabel(
                self.text_label_frame,
                text="Extracted Text (Editable):",
                font=("Segoe UI", 14, "bold"),
                text_color=self.colors["text"],
                anchor="w"
            )
            self.text_label.pack(side="left")
            self.text_counter = ctk.CTkLabel(
                self.text_label_frame,
                text="",
                font=("Segoe UI", 14),
                text_color=self.colors["secondary_text"],
                anchor="e"
            )
            self.text_counter.pack(side="right")
            
            self.text_display = ctk.CTkTextbox(
                self.text_frame,
                width=300,
                height=150,
                fg_color=self.colors["input_bg"],
                text_color=self.colors["text"],
                border_color=self.colors["button"],
                border_width=2,
                font=("Segoe UI", 12)
            )
            self.text_display.pack(pady=(0, 5), fill="both", expand=True)

            # Language selection
            self.lang_frame = self.create_form_group(self.content_frame, "Choose Language")
            
            self.lang_search_entry = ctk.CTkEntry(
                self.lang_frame,
                width=300,
                height=30,
                fg_color=self.colors["input_bg"],
                text_color=self.colors["text"],
                border_color=self.colors["border"],
                placeholder_text="Search languages..."
            )
            self.lang_search_entry.pack(pady=(5, 5), fill="x", anchor="center")
            
            self.lang_search_button = ctk.CTkButton(
                self.lang_frame,
                text="Search",
                command=self.search_languages,
                fg_color=self.colors["button"],
                hover_color=self.colors["button_hover"],
                corner_radius=6,
                width=80,
                height=30
            )
            self.lang_search_button.pack(pady=(0, 5), anchor="center")
            
            self.lang_scroll_frame = ctk.CTkScrollableFrame(
                self.lang_frame,
                width=300,
                height=100,
                fg_color=self.colors["input_bg"],
                corner_radius=6
            )
            self.lang_scroll_frame.pack(pady=(0, 10), fill="x", anchor="center")
            
            self.language_options = [(lang, code) for code, lang in LANGUAGES.items()]
            self.language_options.sort(key=lambda x: x[0])
            self.full_language_list = [f"{lang} ({code})" for lang, code in self.language_options]
            self.current_language_list = self.full_language_list.copy()
            self.language_var = ctk.StringVar(value="")  # Default to no selection
            self.update_language_options()

            # Format selection
            self.format_frame = self.create_form_group(self.content_frame, "Choose Audio Format")
            self.format_combobox = ctk.CTkComboBox(
                self.format_frame,
                values=["MP3", "WAV", "OGG"],
                width=300,
                height=30,
                fg_color=self.colors["input_bg"],
                text_color=self.colors["text"],
                border_color=self.colors["border"],
                button_color=self.colors["button"],
                button_hover_color=self.colors["button_hover"],
                dropdown_fg_color=self.colors["input_bg"],
                dropdown_text_color=self.colors["text"]
            )
            self.format_combobox.set("MP3")
            self.format_combobox.pack(pady=(5, 10), fill="x", anchor="center")

            # Voice selection
            self.voice_frame = self.create_form_group(self.content_frame, "Choose Voice")
            engine = pyttsx3.init()
            self.all_voices = engine.getProperty('voices')
            self.voice_options = ["None (Default)"] + [f"{voice.name} ({voice.id.split('.')[-1]})" for voice in self.all_voices]
            self.voice_combobox = ctk.CTkComboBox(
                self.voice_frame,
                values=self.voice_options,
                width=300,
                height=30,
                fg_color=self.colors["input_bg"],
                text_color=self.colors["text"],
                border_color=self.colors["border"],
                button_color=self.colors["button"],
                button_hover_color=self.colors["button_hover"],
                dropdown_fg_color=self.colors["input_bg"],
                dropdown_text_color=self.colors["text"],
                command=self.on_voice_select
            )
            self.voice_combobox.set("None (Default)")
            self.voice_combobox.pack(pady=(5, 10), fill="x", anchor="center")

            # Email option
            self.email_frame = self.create_form_group(self.content_frame, "Email Options")
            self.email_checkbox = ctk.CTkCheckBox(
                self.email_frame,
                text="Send to Email",
                fg_color=self.colors["accent"],
                hover_color=self.colors["accent_light"],
                text_color=self.colors["text"],
                checkbox_height=20,
                checkbox_width=20
            )
            self.email_checkbox.pack(pady=(5, 10), anchor="center")
            
            self.email_entry = ctk.CTkEntry(
                self.email_frame,
                width=300,
                height=30,
                fg_color=self.colors["input_bg"],
                text_color=self.colors["text"],
                border_color=self.colors["border"],
                placeholder_text="recipient@example.com"
            )
            self.email_entry.pack(pady=(0, 10), fill="x", anchor="center")

            # Bottom frame
            self.bottom_frame = ctk.CTkFrame(
                self.right_panel,
                fg_color=self.colors["bg"],
                corner_radius=15,
                width=450
            )
            self.bottom_frame.pack(pady=10, padx=20, fill="x")

            self.progress_frame = ctk.CTkFrame(self.bottom_frame, fg_color="transparent", corner_radius=10)
            self.progress_frame.pack(pady=(0, 5), fill="x")
            
            self.progress = ctk.CTkProgressBar(
                self.progress_frame,
                width=300,
                height=15,
                fg_color=self.colors["widget_bg"],
                progress_color=self.colors["accent_gradient"],
                orientation="horizontal"
            )
            self.progress.pack(pady=(0, 5), anchor="center")
            self.progress.set(0)
            
            self.status_label = ctk.CTkLabel(
                self.bottom_frame,
                text="Ready to convert PDFs to audio",
                font=("Segoe UI", 12),
                text_color=self.colors["secondary_text"]
            )
            self.status_label.pack(pady=(0, 10), anchor="center")

            self.convert_button = ctk.CTkButton(
                self.bottom_frame,
                text="Convert to Audio",
                command=self.start_conversion_thread,
                fg_color=self.colors["button"],
                hover_color=self.colors["button_hover"],
                corner_radius=6,
                font=("Segoe UI", 14),
                height=35,
                width=120
            )
            self.convert_button.pack(pady=(10, 20), anchor="center")

        except Exception as e:
            print(f"Error creating widgets: {e}")
            traceback.print_exc()
            tkinter.messagebox.showerror("Initialization Error", f"Failed to create UI: {e}")

    def create_form_group(self, parent, label_text):
        frame = ctk.CTkFrame(parent, fg_color="transparent", corner_radius=10)
        frame.pack(fill="x", pady=10, padx=20)
        
        label = ctk.CTkLabel(
            frame,
            text=label_text,
            font=("Segoe UI", 14, "bold"),
            text_color=self.colors["text"],
            anchor="w"
        )
        label.pack(anchor="w")
        
        return frame

    def update_language_options(self):
        # Clear existing widgets
        for widget in self.lang_scroll_frame.winfo_children():
            widget.destroy()
        
        # Populate with current language list
        for lang in self.current_language_list:
            rb = ctk.CTkRadioButton(
                self.lang_scroll_frame,
                text=lang,
                variable=self.language_var,
                value=lang,
                command=lambda l=lang: self.on_language_select(l),
                fg_color=self.colors["accent"],
                hover_color=self.colors["accent_light"],
                text_color=self.colors["text"]
            )
            rb.pack(anchor="w", padx=5, pady=2)

    def search_languages(self):
        try:
            search_term = self.lang_search_entry.get().lower()
            if not search_term:
                self.current_language_list = self.full_language_list.copy()
            else:
                self.current_language_list = [lang for lang in self.full_language_list if search_term in lang.lower()]
            self.update_language_options()
        except Exception as e:
            print(f"Error searching languages: {e}")
            self.show_error("Search Error", f"Failed to search languages: {str(e)}")

    def on_language_select(self, selected_language):
        try:
            lang_code = selected_language.split("(")[-1].strip(")")
            self.update_voices(lang_code)
        except Exception as e:
            print(f"Error in on_language_select: {e}")

    def on_voice_select(self, value):
        pass

    def update_voices(self, lang_code):
        try:
            if not lang_code:
                filtered_voices = ["None (Default)"] + [f"{voice.name} ({voice.id.split('.')[-1]})" for voice in self.all_voices]
            else:
                filtered_voices = ["None (Default)"]
                lang_patterns = self.voice_language_map.get(lang_code, [])
                for voice in self.all_voices:
                    voice_name = voice.name.lower()
                    if any(pattern in voice_name for pattern in lang_patterns):
                        filtered_voices.append(f"{voice.name} ({voice.id.split('.')[-1]})")
            
            self.voice_combobox.configure(values=filtered_voices)
            if self.voice_combobox.get() not in filtered_voices:
                self.voice_combobox.set("None (Default)")
        except Exception as e:
            print(f"Error updating voices: {e}")

    def select_pdfs(self):
        try:
            file_paths = tkinter.filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
            if file_paths:
                self.file_paths = list(file_paths)
                self.extracted_texts = ["" for _ in self.file_paths]
                self.current_file_index = 0
                self.display_current_file()
                self.extract_and_display_text()
        except Exception as e:
            print(f"Error selecting PDFs: {e}")
            self.status_label.configure(text=f"Error: {str(e)}")
            self.show_error("File Error", f"Could not open PDF files: {str(e)}")

    def display_current_file(self):
        try:
            if not self.file_paths or self.current_file_index < 0 or self.current_file_index >= len(self.file_paths):
                self.file_display.configure(state="normal")
                self.file_display.delete(0, "end")
                self.file_display.insert(0, "No files selected")
                self.file_display.configure(state="disabled")
                self.prev_button.configure(state="disabled")
                self.next_button.configure(state="disabled")
                self.text_counter.configure(text="")
                return
            
            self.file_display.configure(state="normal")
            self.file_display.delete(0, "end")
            self.file_display.insert(0, os.path.basename(self.file_paths[self.current_file_index]))
            self.file_display.configure(state="disabled")
            self.prev_button.configure(state="normal" if self.current_file_index > 0 else "disabled")
            self.next_button.configure(state="normal" if self.current_file_index < len(self.file_paths) - 1 else "disabled")
            self.text_counter.configure(text=f"{self.current_file_index + 1}/{len(self.file_paths)}")
        except Exception as e:
            print(f"Error displaying current file: {e}")

    def show_prev_file(self):
        try:
            if self.current_file_index > 0:
                self.save_current_text()
                self.current_file_index -= 1
                self.display_current_file()
                self.display_extracted_text()
        except Exception as e:
            print(f"Error showing previous file: {e}")

    def show_next_file(self):
        try:
            if self.current_file_index < len(self.file_paths) - 1:
                self.save_current_text()
                self.current_file_index += 1
                self.display_current_file()
                self.display_extracted_text()
        except Exception as e:
            print(f"Error showing next file: {e}")

    def save_current_text(self):
        try:
            if 0 <= self.current_file_index < len(self.extracted_texts):
                self.extracted_texts[self.current_file_index] = self.text_display.get("0.0", "end").strip()
        except Exception as e:
            print(f"Error saving current text: {e}")

    def extract_and_display_text(self):
        try:
            if not self.file_paths or self.current_file_index < 0:
                return
            
            file_path = self.file_paths[self.current_file_index]
            if not self.extracted_texts[self.current_file_index]:
                text = ""
                # Try PyPDF2 first
                try:
                    with open(file_path, 'rb') as file:
                        reader = PdfReader(file)
                        sections = []
                        for page_num, page in enumerate(reader.pages, 1):
                            extracted = page.extract_text()
                            if extracted and extracted.strip():
                                sections.append(f"Section {page_num}:\n{extracted.strip()}\n{'-'*50}")
                        text = "\n".join(sections)
                except Exception as e:
                    print(f"PyPDF2 failed: {e}")
                
                # If PyPDF2 fails or extracts no usable text, use Tesseract OCR
                if not text.strip():
                    try:
                        self.status_label.configure(text=f"Extracting text with OCR from {os.path.basename(file_path)}...")
                        images = convert_from_path(file_path)  # Convert PDF to images using Poppler
                        sections = []
                        for i, img in enumerate(images, 1):
                            extracted = pytesseract.image_to_string(img)
                            if extracted and extracted.strip():
                                sections.append(f"Section {i}:\n{extracted.strip()}\n{'-'*50}")
                        text = "\n".join(sections)
                        if not text.strip():
                            raise Exception("No text extracted by OCR")
                    except Exception as e:
                        print(f"Tesseract OCR failed: {e}")
                        self.status_label.configure(
                            text=f"Failed to extract text from {os.path.basename(file_path)}",
                            text_color="#ff6b6b"
                        )
                        self.show_error("OCR Error", f"Could not extract text with OCR: {str(e)}")
                        return
                
                self.extracted_texts[self.current_file_index] = text[:5000] + ("...\n[Truncated]" if len(text) > 5000 else "")
            
            self.text_display.delete("0.0", "end")
            self.text_display.insert("0.0", self.extracted_texts[self.current_file_index])
            self.status_label.configure(
                text=f"PDF loaded: {os.path.basename(file_path)}",
                text_color=self.colors["accent"]
            )
                
        except Exception as e:
            print(f"Error extracting text: {e}")
            self.status_label.configure(
                text=f"Error reading PDF: {str(e)}",
                text_color="#ff6b6b"
            )
            self.show_error("PDF Error", f"Error reading PDF: {str(e)}")

    def display_extracted_text(self):
        self.extract_and_display_text()

    def start_conversion_thread(self):
        try:
            if not self.file_paths:
                self.show_error("Error", "Please select at least one PDF file.")
                return
            if not self.language_var.get():
                self.show_error("Error", "Please select a language for conversion.")
                return
                
            self.save_current_text()
            self.convert_button.configure(state="disabled")
            threading.Thread(target=self.convert_pdfs, daemon=True).start()
        except Exception as e:
            print(f"Error starting conversion thread: {e}")
            self.show_error("Thread Error", f"Could not start conversion: {str(e)}")
            self.convert_button.configure(state="normal")

    def show_error(self, title, message):
        try:
            ctk.CTkMessagebox(title=title, message=message, icon="cancel")
        except:
            tkinter.messagebox.showerror(title=title, message=message)

    def show_info(self, title, message):
        try:
            ctk.CTkMessagebox(title=title, message=message, icon="info")
        except:
            tkinter.messagebox.showinfo(title=title, message=message)

    def show_success(self, title, message):
        try:
            ctk.CTkMessagebox(title=title, message=message, icon="check")
        except:
            tkinter.messagebox.showinfo(title=title, message=message)

    def update_status(self, text, progress=None):
        try:
            self.root.after(0, lambda: self.status_label.configure(text=text))
            if progress is not None:
                self.root.after(0, lambda: self.progress.set(progress))
        except Exception as e:
            print(f"Error updating status: {e}")

    def convert_pdfs(self):
        try:
            from gtts import gTTS
            
            # Create a timestamp-based folder for this conversion
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            base_output_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'audios')
            output_folder = os.path.join(base_output_folder, f"conversion_{timestamp}")
            os.makedirs(output_folder, exist_ok=True)
            
            total_files = len(self.file_paths)
            output_paths = []
            existing_files = []

            # Check for existing files in any subfolder of 'audios'
            for file_path in self.file_paths:
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                audio_format = self.format_combobox.get().lower()
                expected_name = f"{base_name}.{audio_format}"
                for root, _, files in os.walk(base_output_folder):
                    if expected_name in files:
                        existing_files.append(os.path.join(root, expected_name))

            # Show dialog if files already exist
            if existing_files:
                self.root.after(0, lambda: self.show_info(
                    "Files Already Exist",
                    f"The following audio files already exist:\n" + "\n".join(existing_files) +
                    "\nNew files will be created in a new folder."
                ))

            # Convert files
            for idx, (file_path, text) in enumerate(zip(self.file_paths, self.extracted_texts)):
                self.update_status(f"Processing file {idx + 1}/{total_files}...", idx / total_files * 0.5)
                
                if not text:
                    self.update_status(f"No text found in file {idx + 1}", 0)
                    self.show_error("Error", f"No text could be extracted from {os.path.basename(file_path)}.")
                    continue

                if len(text) > 100000:
                    text = text[:100000]
                    self.root.after(0, lambda: self.show_error("Warning", 
                        f"Text in {os.path.basename(file_path)} truncated to 100,000 characters."))

                selected_language = self.language_var.get()
                lang_code = selected_language.split("(")[-1].strip(")")
                
                if lang_code != 'en':
                    try:
                        self.update_status(f"Translating file {idx + 1}...", 0.5 + (idx / total_files * 0.2))
                        translator = Translator()
                        translated_text = ''
                        chunk_size = 1000
                        chunks = [text[i:i+chunk_size] for i in range(0, len(text), chunk_size)]
                        
                        for i, chunk in enumerate(chunks):
                            translated_chunk = translator.translate(chunk, dest=lang_code).text
                            translated_text += translated_chunk
                        text = translated_text
                    except Exception as e:
                        print(f"Translation failed for file {idx + 1}: {e}")
                        self.update_status(f"Translation failed for file {idx + 1}, using original text", 0.6)

                base_name = os.path.splitext(os.path.basename(file_path))[0]
                audio_format = self.format_combobox.get().lower()
                output_path = os.path.join(output_folder, f"{base_name}.{audio_format}")
                temp_mp3 = os.path.join(output_folder, f"{base_name}_temp.mp3")

                selected_voice = self.voice_combobox.get()
                if selected_voice == "None (Default)":
                    tts = gTTS(text=text[:5000], lang=lang_code)
                    tts.save(temp_mp3)
                else:
                    engine = pyttsx3.init()
                    voices = engine.getProperty('voices')
                    voice_id = next((v.id for v in voices if v.name in selected_voice), voices[0].id)
                    engine.setProperty('voice', voice_id)
                    engine.save_to_file(text[:5000], temp_mp3)
                    engine.runAndWait()
                    time.sleep(1)  # Ensure file is fully written
                    if not os.path.exists(temp_mp3) or os.path.getsize(temp_mp3) == 0:
                        raise Exception(f"Failed to create temporary MP3 file for {base_name}")

                # Convert to desired format using FFmpeg
                if audio_format in ['wav', 'ogg']:
                    ffmpeg_cmd = [
                        'ffmpeg', '-i', temp_mp3, '-y'  # -y to overwrite without prompting
                    ]
                    if audio_format == 'wav':
                        ffmpeg_cmd.append(output_path)
                    elif audio_format == 'ogg':
                        ffmpeg_cmd.extend(['-c:a', 'libvorbis', output_path])
                    
                    result = subprocess.run(ffmpeg_cmd, capture_output=True, text=True)
                    if result.returncode != 0:
                        raise Exception(f"FFmpeg conversion failed: {result.stderr}")
                    os.remove(temp_mp3)
                else:  # MP3
                    os.rename(temp_mp3, output_path)

                output_paths.append(output_path)

            self.update_status("Finalizing...", 0.9)
            
            # Show "Audio Files Created" dialog if no email or after conversion
            self.root.after(0, lambda: self.show_success(
                "Audio Files Created",
                f"{len(output_paths)} audio files created:\n" + "\n".join(output_paths)
            ))

            # Handle email if checkbox is selected
            if self.email_checkbox.get() == 1 and output_paths:
                recipient = self.email_entry.get()
                if recipient:
                    self.send_email(recipient, output_paths)
                else:
                    self.update_status("No recipient email provided", 0.95)
            else:
                self.update_status(f"Success! Saved {len(output_paths)} files", 1)
                
        except Exception as e:
            print(f"Error in conversion process: {e}")
            traceback.print_exc()
            self.update_status(f"Error: {str(e)}", 0)
            self.show_error("Error", f"An unexpected error occurred: {str(e)}")
        finally:
            self.root.after(0, lambda: self.convert_button.configure(state="normal"))

    def send_email(self, to_email, file_paths):
        try:
            self.update_status("Preparing email...", 0.95)
            
            msg = EmailMessage()
            msg["Subject"] = "Your Audio Files from PDF to Audio Converter"
            msg["From"] = self.sender_email
            msg["To"] = to_email
            msg.set_content("Please find your audio files attached.")

            for file_path in file_paths:
                if not os.path.exists(file_path):
                    self.update_status(f"File not found: {file_path}", 0.95)
                    continue

                with open(file_path, "rb") as f:
                    file_data = f.read()
                    file_name = os.path.basename(file_path)
                    
                file_ext = file_path.split('.')[-1].lower()
                subtype = {'mp3': 'mpeg', 'wav': 'wav', 'ogg': 'ogg'}.get(file_ext, file_ext)
                msg.add_attachment(
                    file_data, 
                    maintype="audio", 
                    subtype=subtype, 
                    filename=file_name
                )

            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(self.sender_email, self.sender_password)
                smtp.send_message(msg)
                
            self.update_status("Email sent successfully!", 1)
            self.root.after(0, lambda: self.show_success("Email Sent", 
                "The audio files have been sent successfully!"))
            
        except Exception as e:
            print(f"Email error: {e}")
            self.update_status(f"Email failed: {str(e)}", 0.95)
            self.show_error("Email Error", f"Failed to send email: {str(e)}")

    def reposition_widgets(self):
        try:
            self.content_frame.pack(pady=10, padx=20, fill="both", expand=True)
            self.bottom_frame.pack(pady=10, padx=20, fill="x")
        except Exception as e:
            print(f"Error repositioning widgets: {e}")

if __name__ == "__main__":
    try:
        root = ctk.CTk()
        app = PDFToAudioApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Fatal error: {e}")
        traceback.print_exc()
        tkinter.messagebox.showerror("Fatal Error", f"The application crashed: {str(e)}")