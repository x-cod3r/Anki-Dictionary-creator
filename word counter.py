import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from collections import Counter
import re
import asyncio
from googletrans import Translator
from gtts import gTTS
import os
import time
import logging
import pygame
import genanki
import random
import shutil
import pyperclip

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

class WordCounterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Anki dictionary creator by insta: @Mahmoud.aboulnasr")
        self.root.geometry("600x400")  # Increased height

        # Initialize pygame mixer
        pygame.mixer.init()

        # Style the application
        style = ttk.Style()
        style.configure("TLabel", font=("Arial", 12))
        style.configure("TButton", font=("Arial", 12), padding=10)
        style.configure("TEntry", font=("Arial", 12), padding=5)
        style.configure("TFrame", padding=10)
        style.configure("TScrollbar", padding=0)

        # Main container frame
        frame = ttk.Frame(root)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        # File selection label
        self.file_label = ttk.Label(frame, text="No file(s) selected")
        self.file_label.grid(row=0, column=0, columnspan=6, sticky="w", pady=5)

        self.file_paths = []
        self.translator = Translator()

        # Buttons
        self.browse_button = ttk.Button(frame, text="Select File(s)", command=self.browse_files)
        self.browse_button.grid(row=1, column=0, sticky="w", pady=5)

        self.process_button = ttk.Button(frame, text="Process Files", command=self.process_files)
        self.process_button.grid(row=1, column=1, sticky="e", pady=5)

        self.translation_var = tk.StringVar(value="None")  # "None", "English", "Arabic"
        self.translation_menu = ttk.OptionMenu(frame, self.translation_var, "None", "None", "English", "Arabic",
                                                command=self.update_translation)
        self.translation_menu.grid(row=1, column=2, sticky="e", pady=5)

        # Input language selection
        self.language_var = tk.StringVar(value="en")  # Default is English
        self.language_menu = ttk.OptionMenu(frame, self.language_var, "en", "en", "ar", "de", "es", "fr", "it", "es",
                                            "pt-PT", "tr", "nl", "iw", "ja", "ko", "ru", "zh-cn")
        self.language_menu.grid(row=1, column=4, sticky="ew", pady=5)

        language_label = ttk.Label(frame, text="Input Language:")
        language_label.grid(row=1, column=3, sticky="ew", pady=5)

        # Speak All Words Button
        self.speak_all_button = ttk.Button(frame, text="Speak All Words", command=self.speak_all_words)
        self.speak_all_button.grid(row=1, column=5, sticky="e", pady=5)

        # Word Limit entry
        self.word_limit_label = ttk.Label(frame, text="Word Limit:")
        self.word_limit_label.grid(row=3, column=2, sticky="w", pady=5)
        self.word_limit_entry = ttk.Entry(frame)
        self.word_limit_entry.grid(row=3, column=3, sticky="ew", pady=5)
        self.word_limit_entry.insert(0, "0")  # Default to 0, which means no limit

        # Deck name label and entry
        self.deck_name_label = ttk.Label(frame, text="Deck Name:")
        self.deck_name_label.grid(row=3, column=4, sticky="w", pady=5)
        self.deck_name_entry = ttk.Entry(frame)
        self.deck_name_entry.grid(row=3, column=5, sticky="ew", pady=5)
        self.deck_name_entry.insert(0, "Word Deck")  # Default deck name

        # Export menu
        self.export_var = tk.StringVar(value="word_front_translation_back")
        self.export_menu = ttk.OptionMenu(frame, self.export_var, "word_front_translation_back",
                                          "word_front_translation_back",
                                          "translation_front_word_back", "word_front_speech_back",
                                          "translation_front_speech_word_back", "word_front_speech_translation_back")
        self.export_menu.grid(row=3, column=0, sticky="e", pady=5)

        export_label = ttk.Label(frame, text="Export As:")
        export_label.grid(row=3, column=0, sticky="w", pady=5)

        # Export Button
        self.export_button = ttk.Button(frame, text="Export Anki Deck", command=self.export_anki_deck)
        self.export_button.grid(row=3, column=1, sticky="e", pady=5)

        # Progress bar
        self.progress_bar = ttk.Progressbar(frame, orient="horizontal", length=200, mode="determinate")
        self.progress_bar.grid(row=4, column=0, columnspan=6, sticky="ew", pady=10)


        # Treeview
        self.result_tree = ttk.Treeview(frame, columns=("Word", "Count", "Translation", "Speak"), show="headings")
        self.result_tree.heading("Word", text="Word")
        self.result_tree.heading("Count", text="Count")
        self.result_tree.heading("Translation", text="Translation")
        self.result_tree.heading("Speak", text="Speak")
        self.result_tree.column("Translation", width=200)
        self.result_tree.column("Speak", width=100)
        self.result_tree.grid(row=5, column=0, columnspan=6, sticky="nsew", pady=10)
        self.result_tree.bind("<Button-1>", self.treeview_click)
        self.result_tree.bind("<Control-c>", self.copy_selected_words)

        # Vertical scrollbar for treeview
        self.vsb = ttk.Scrollbar(frame, orient="vertical", command=self.result_tree.yview)
        self.vsb.grid(row=4, column=6, sticky="ns")
        self.result_tree.configure(yscrollcommand=self.vsb.set)

        # Configure grid layout
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(2, weight=1)
        frame.columnconfigure(3, weight=1)
        frame.columnconfigure(4, weight=1)
        frame.columnconfigure(5, weight=1)
        frame.rowconfigure(5, weight=1)

    def browse_files(self):
        filetypes = (("Text files", "*.txt"), ("All files", "*.*"))
        paths = filedialog.askopenfilenames(title="Select Text Files", filetypes=filetypes)
        if paths:
            self.file_paths = list(paths)
            file_count = len(self.file_paths)
            self.file_label.config(text=f"{file_count} file(s) selected")
        else:
            self.file_paths = []
            self.file_label.config(text="No file(s) selected")

    def update_translation(self, event=None):
        asyncio.run(self.process_files())

    async def _translate_word(self, word, target_lang):
        try:
            translation_obj = await self.translator.translate(word, dest=target_lang)
            return translation_obj.text
        except Exception as e:
            return f"Error: {e}"

    def _extract_words(self, text, lang):
        """Extracts words based on the input language."""
        if lang == "zh-cn":
            # Regex to keep only Chinese characters and some common punctuation
            return re.findall(r'[\u4e00-\u9fff]', text.lower()) #changed to word by word
        else:
             return re.findall(r'\b[a-zA-Z]{2,}\b', text.lower())  # Default behavior for other languages


    async def process_files(self):
        if not self.file_paths:
            messagebox.showerror("Error", "No file(s) selected.")
            return

        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = 100
        self.progress_bar.update()
        all_words = []
        try:
            for index, path in enumerate(self.file_paths):
                with open(path, "r", encoding="utf-8") as file:
                    text = file.read()
                    words = self._extract_words(text,self.language_var.get())
                    all_words.extend(words)
                self.progress_bar["value"] = int(((index + 1) / len(self.file_paths)) * 100)
                self.progress_bar.update()
        except Exception as e:
            messagebox.showerror("Error", f"Could not process files because of the following error:\n{e}")
            return

        word_counts = Counter(all_words)
        sorted_words = word_counts.most_common()

        # Clear previous results
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        target_lang = self.translation_var.get()
        if target_lang == 'English':
            target_lang = 'en'
        elif target_lang == 'Arabic':
            target_lang = 'ar'

        word_limit = self.word_limit_entry.get()
        try:
            word_limit = int(word_limit)
        except ValueError:
            messagebox.showerror("Error", "Invalid word limit. Please enter a number.")
            return

        # Insert results into the treeview, applying word limit
        for index, (word, count) in enumerate(sorted_words):
            if word_limit > 0 and index >= word_limit:
                break

            translation = ""
            if target_lang != "None":
                translation = await self._translate_word(word, target_lang)

            self.result_tree.insert("", "end", values=(word, count, translation, "🔊"))
            self.progress_bar["value"] = int(
                ((index + 1) / min(len(sorted_words), word_limit if word_limit > 0 else len(sorted_words))) * 100)
            self.progress_bar.update()

        self.progress_bar["value"] = 100
        self.progress_bar.update()

        # **New Code Snippet Start**
        total_output_words = len(sorted_words)
        messagebox.showinfo("Processing Complete", f"Total output words: {total_output_words}")
         # **New Code Snippet End**

    def text_to_speech(self, text, lang='en', filename='output.mp3'):
        """Converts text to speech and saves it as an mp3 file."""
        try:
            logging.debug(f"Starting text-to-speech conversion for text: '{text}'")
            tts = gTTS(text=text, lang=lang)
            tts.save(filename)
            logging.info(f"Audio saved to {filename}")
            return filename
        except Exception as e:
            logging.error(f"Error during speech generation: {e}")
            return None

    def play_audio(self, filename):
        """Plays the audio file using pygame."""
        try:
            logging.info(f"Attempting to play audio file: {filename}")

            # Initialize pygame mixer
            pygame.mixer.init()

            # Load the audio file
            pygame.mixer.music.load(filename)

            # Play the audio
            pygame.mixer.music.play()

            # Wait for the audio to finish playing
            while pygame.mixer.music.get_busy():
                pygame.time.Clock().tick(10)  # Control the loop speed

            logging.info(f"Audio playback finished successfully for: {filename}")
        except Exception as e:
            logging.error(f"Error during audio playback: {e}")
        finally:
            # Clean up pygame mixer
            pygame.mixer.quit()

    def treeview_click(self, event):
        item_id = self.result_tree.identify_row(event.y)
        column = self.result_tree.identify_column(event.x)

        if column == '#4' and item_id:
            values = self.result_tree.item(item_id)['values']
            word = values[0]
            self.speak_word(word)

    def speak_word(self, word):
        if word:
            output_file = 'temp_audio.mp3'
            audio_file = self.text_to_speech(word, self.language_var.get(), output_file)

            if audio_file:
                # Play the audio
                logging.info("Playing audio")
                self.play_audio(audio_file)

                # Optional: Remove file after playing
                time.sleep(0.5)
                try:
                    os.remove(audio_file)
                    logging.info("Audio file removed.")
                except:
                    pass

    def speak_all_words(self):
        # Get all words from the treeview
        words = [self.result_tree.item(item)['values'][0] for item in self.result_tree.get_children()]
        if words:
            # Speak words one by one
            for word in words:
                self.speak_word(word)
        else:
            messagebox.showinfo("Info", "No words to speak.")

    def export_anki_deck(self):
        """Exports the results as an Anki deck."""
        deck_name = self.deck_name_entry.get()  # Gets from the input
        model_name = "Basic Card Model"

        export_type = self.export_var.get()
        # Generate random IDs
        model_id = random.randrange(1 << 30, 1 << 31)
        deck_id = random.randrange(1 << 30, 1 << 31)

        model = genanki.Model(
            model_id,
            model_name,
            fields=[
                {"name": "Front"},
                {"name": "Back"},
                {"name": "Audio"},
            ],
            templates=[
                {
                    "name": "Card 1",
                    "qfmt": '<div style="text-align: center; font-size: 1.3em;"><b>{{Front}}</b></div>',
                    "afmt": ''' 
                                <div style="text-align: center; font-size: 1.3em;">{{Front}}</div><br><div style="text-align: center; font-size: 1.3em;">{{Back}}</div>
                                <div id="player_container" style="display: none;">
                                  <audio controls>
                                  <source src="{{Audio}}">
                                  </audio>
                                </div>
                                <button onclick="var player = document.getElementById('player_container'); player.style.display = (player.style.display === 'none') ? 'block' : 'none';"> Toggle Audio </button>
                                <hr id="answer">

                           ''',
                }
            ],
            css=".card {font-family: arial;font-size: 20px;text-align: center;color: black;background-color: white;} hr#answer {border-color: black;}"
        )
        deck = genanki.Deck(deck_id, deck_name)

        # Get the directory to save the apkg
        filepath = filedialog.asksaveasfilename(
            defaultextension=".apkg",
            filetypes=[("Anki Package", "*.apkg")],
            title="Save Anki Deck As",
        )
        if not filepath:
            return  # User canceled the dialog

        # Get the directory for the audio files
        audio_dir = filedialog.askdirectory(title="Select Audio Export Folder")
        if not audio_dir:
            return  # User canceled the dialog

        media_files = []

        total_words = len(self.result_tree.get_children())
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = total_words
        self.progress_bar.update()

        package = genanki.Package(deck)
        # Prepare the data in a structured format based on selected export mode
        for index, item in enumerate(self.result_tree.get_children()):
            values = self.result_tree.item(item)['values']
            word, _, translation, _ = values  # _ means we don't need those

            front = ""
            back = ""
            audio_field = ""
            audio_filepath = ""

            if export_type == "word_front_translation_back":
                front = word
                back = translation
            elif export_type == "translation_front_word_back":
                front = translation
                back = word
            elif export_type == "word_front_speech_back":
                front = word
                back = ""
                audio_filepath = os.path.join(audio_dir, f"{word}.mp3")
                audio_file = self.text_to_speech(word, self.language_var.get(),
                                                 audio_filepath)  # Save to the folder the user specified
                if audio_file:
                    audio_field = f"{os.path.basename(audio_file)}"
                    package.media_files.append(audio_file)
            elif export_type == "translation_front_speech_word_back":
                front = translation
                audio_filepath = os.path.join(audio_dir, f"{word}.mp3")
                audio_file = self.text_to_speech(word, self.language_var.get(), audio_filepath)
                if audio_file:
                    audio_field = f"{os.path.basename(audio_file)}"
                    package.media_files.append(audio_file)
                back = word  # Word back here
            elif export_type == "word_front_speech_translation_back":
                front = word
                audio_filepath = os.path.join(audio_dir, f"{word}.mp3")
                audio_file = self.text_to_speech(word, self.language_var.get(), audio_filepath)
                if audio_file:
                    audio_field = f"{os.path.basename(audio_file)}"
                    package.media_files.append(audio_file)
                back = translation  # add the translation back here

            note = genanki.Note(
                model=model,
                fields=[front, back, audio_field]
            )
            deck.add_note(note)
            self.progress_bar["value"] = index + 1
            self.progress_bar.update()

        try:
            package.write_to_file(filepath)

            # Open the folder after saving on windows, mac or linux
            if os.name == 'nt':  # If Windows
                os.startfile(os.path.dirname(filepath))
            elif os.name == 'posix':  # If Linux or Mac
                os.system(f'open "{os.path.dirname(filepath)}"')
            else:
                messagebox.showerror("Error", "Could not open file explorer")
        except Exception as e:
            messagebox.showerror("Error", f"Could not generate Anki Deck because of the following error:\n{e}")
        self.progress_bar["value"] = 0
        self.progress_bar.update()

    def copy_selected_words(self, event):
        """Copy selected words from the Treeview to clipboard."""
        selected_items = self.result_tree.selection()
        if selected_items:
            words = [self.result_tree.item(item)['values'][0] for item in selected_items]
            pyperclip.copy("\n".join(words))
            messagebox.showinfo("Copied", "Selected words copied to clipboard.")


# Run the app
if __name__ == "__main__":
    root = tk.Tk()
    app = WordCounterApp(root)
    root.mainloop()