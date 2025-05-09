import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from collections import Counter
import re
import asyncio
from googletrans import Translator # Ensure this is version 4.0.0-rc1 or compatible
                                   # pip install googletrans==4.0.0-rc1
import gtts 
import os
import time
import logging
import pygame
import genanki
import random
import pyperclip
import sys 
import subprocess 
import inspect # For inspect.iscoroutinefunction

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class WordCounterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Anki dictionary creator by insta: @Mahmoud.aboulnasr")
        self.root.geometry("750x550")

        try:
            pygame.mixer.init()
        except pygame.error as e:
            logging.warning(f"Pygame mixer could not be initialized: {e}. Audio playback might not work.")
            # messagebox.showwarning("Audio Warning", f"Pygame mixer could not be initialized: {e}\nAudio playback might not work.")

        style = ttk.Style()
        try:
            style.theme_use('clam') 
        except tk.TclError:
            logging.warning("Clam theme not available, using default.")
        style.configure("TLabel", font=("Arial", 11))
        style.configure("TButton", font=("Arial", 11), padding=6)
        style.configure("TEntry", font=("Arial", 11), padding=3)
        style.configure("TFrame", padding=10)
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))

        frame = ttk.Frame(root)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.file_label = ttk.Label(frame, text="No file(s) selected")
        self.file_label.grid(row=0, column=0, columnspan=6, sticky="w", pady=(0,10))

        self.file_paths = []
        try:
            self.translator = Translator()
        except Exception as e:
            logging.error(f"Failed to initialize Translator: {e}")
            messagebox.showerror("Translator Error", f"Failed to initialize Google Translator: {e}\nPlease check your internet connection or googletrans library version.")
            self.translator = None

        self.browse_button = ttk.Button(frame, text="Select File(s)", command=self.browse_files)
        self.browse_button.grid(row=1, column=0, sticky="ew", padx=2)

        self.process_button = ttk.Button(frame, text="Process Files", command=lambda: asyncio.run(self.process_files()))
        self.process_button.grid(row=1, column=1, sticky="ew", padx=2)
        
        language_label = ttk.Label(frame, text="Input Lang:")
        language_label.grid(row=1, column=2, sticky="w", padx=(5,0))
        self.language_var = tk.StringVar(value="en")
        lang_options = ["en", "ar", "de", "es", "fr", "it", "pt", "tr", "nl", "he", "ja", "ko", "ru", "zh-cn", "sv", "pl", "fi", "el", "hi", "id"]
        self.language_menu = ttk.OptionMenu(frame, self.language_var, "en", *lang_options)
        self.language_menu.grid(row=1, column=3, sticky="ew", padx=2)

        translation_label = ttk.Label(frame, text="Translate To:")
        translation_label.grid(row=1, column=4, sticky="w", padx=(5,0))
        self.translation_var = tk.StringVar(value="None")
        trans_options = ["None", "English", "Arabic", "German", "Spanish", "French", "Italian", "Portuguese"]
        self.translation_menu = ttk.OptionMenu(frame, self.translation_var, "None", *trans_options)
        self.translation_menu.grid(row=1, column=5, sticky="ew", padx=2)
        self.translation_var.trace_add("write", lambda *args: asyncio.run(self.process_files()) if self.file_paths else None)

        self.word_limit_label = ttk.Label(frame, text="Word Limit:")
        self.word_limit_label.grid(row=2, column=0, sticky="w", pady=5, padx=2)
        self.word_limit_entry = ttk.Entry(frame, width=7)
        self.word_limit_entry.grid(row=2, column=1, sticky="ew", pady=5, padx=2)
        self.word_limit_entry.insert(0, "50") # Default to 50 for faster testing

        self.deck_name_label = ttk.Label(frame, text="Deck Name:")
        self.deck_name_label.grid(row=2, column=2, sticky="w", pady=5, padx=(5,0))
        self.deck_name_entry = ttk.Entry(frame, width=15)
        self.deck_name_entry.grid(row=2, column=3, sticky="ew", pady=5, padx=2)
        self.deck_name_entry.insert(0, "Word Deck")

        self.speak_all_button = ttk.Button(frame, text="Speak All Visible", command=self.speak_all_words)
        self.speak_all_button.grid(row=2, column=4, columnspan=2, sticky="ew", pady=5, padx=2)

        export_label = ttk.Label(frame, text="Export As:")
        export_label.grid(row=3, column=0, sticky="w", pady=5, padx=2)
        self.export_var = tk.StringVar(value="word_front_translation_back")
        export_options = [
            "word_front_translation_back", "translation_front_word_back",
            "word_front_speech_back", "translation_front_speech_word_back",
            "word_front_speech_translation_back"
        ]
        self.export_menu = ttk.OptionMenu(frame, self.export_var, export_options[0], *export_options)
        self.export_menu.grid(row=3, column=1, columnspan=3, sticky="ew", pady=5, padx=2)
        
        self.export_button = ttk.Button(frame, text="Export Anki Deck", command=self.export_anki_deck)
        self.export_button.grid(row=3, column=4, columnspan=2, sticky="ew", pady=5, padx=2)

        self.progress_bar = ttk.Progressbar(frame, orient="horizontal", length=200, mode="determinate")
        self.progress_bar.grid(row=4, column=0, columnspan=6, sticky="ew", pady=(10,5))

        self.result_tree = ttk.Treeview(frame, columns=("Word", "Count", "Translation", "Speak"), show="headings")
        self.result_tree.heading("Word", text="Word")
        self.result_tree.heading("Count", text="Count")
        self.result_tree.heading("Translation", text="Translation")
        self.result_tree.heading("Speak", text="Speak")
        self.result_tree.column("Word", width=150, stretch=tk.YES)
        self.result_tree.column("Count", width=60, stretch=tk.NO, anchor="center")
        self.result_tree.column("Translation", width=200, stretch=tk.YES)
        self.result_tree.column("Speak", width=60, stretch=tk.NO, anchor="center")
        self.result_tree.grid(row=5, column=0, columnspan=6, sticky="nsew", pady=(0,5))
        self.result_tree.bind("<ButtonRelease-1>", self.treeview_click)
        self.result_tree.bind("<Control-c>", self.copy_selected_words)

        self.vsb = ttk.Scrollbar(frame, orient="vertical", command=self.result_tree.yview)
        self.vsb.grid(row=5, column=6, sticky="ns")
        self.result_tree.configure(yscrollcommand=self.vsb.set)

        for i in range(6): frame.columnconfigure(i, weight=1)
        frame.rowconfigure(5, weight=1)

    def browse_files(self):
        filetypes = (("Text files", "*.txt"), ("All files", "*.*"))
        paths = filedialog.askopenfilenames(title="Select Text Files", filetypes=filetypes)
        if paths:
            self.file_paths = list(paths)
            basenames = [os.path.basename(p) for p in paths]
            display_text = f"{len(paths)} file(s): {', '.join(basenames)}"
            if len(display_text) > 100:
                display_text = f"{len(paths)} file(s): {basenames[0]}..."
            self.file_label.config(text=display_text)
        else:
            self.file_paths = []
            self.file_label.config(text="No file(s) selected")

    async def _translate_word(self, word, target_lang_code):
        if not self.translator:
            return "Translator N/A"
        if not word.strip():
            return ""
        
        # Initial check
        is_async_translate = False
        if hasattr(self.translator, 'translate'):
             # For googletrans 4.0.0rc1, translate is NOT a coroutine function itself.
             # The issue might be how it interacts when called repeatedly or if the underlying httpx client is used in a specific way.
             # So, we will primarily rely on asyncio.to_thread for this version of googletrans.
             # The original error was that asyncio.to_thread was *returning* a coroutine, which is not its standard behavior for sync functions.
             # This suggests the Translator object itself or its translate method was being replaced or wrapped.
             pass # Keep is_async_translate as False, default to using to_thread

        try:
            for attempt in range(3): 
                try:
                    # For googletrans 4.0.0rc1, 'translate' is a synchronous method.
                    # The most reliable way to call it from an async context is via to_thread.
                    logging.debug(f"Attempting asyncio.to_thread for translator.translate (attempt {attempt+1}) for '{word}'")
                    translation_obj = await asyncio.to_thread(self.translator.translate, word, dest=target_lang_code)
                    
                    # Check if translation_obj is None or if .text is missing
                    if translation_obj is None:
                        logging.warning(f"Translator.translate returned None for '{word}' (attempt {attempt+1}).")
                        raise ValueError("Translator returned None") # Force retry or failure
                    if not hasattr(translation_obj, 'text'):
                        logging.warning(f"translation_obj for '{word}' has no 'text' attribute. Type: {type(translation_obj)} (attempt {attempt+1})")
                        # This is where the 'coroutine' object has no attribute 'text' would hit if to_thread somehow returned a coroutine
                        if inspect.iscoroutine(translation_obj):
                            logging.error(f"asyncio.to_thread returned a coroutine! This is unexpected for googletrans 4.0.0rc1. Attempting to await it.")
                            actual_result = await translation_obj # Try awaiting the unexpected coroutine
                            if hasattr(actual_result, 'text'):
                                return actual_result.text
                            else:
                                raise AttributeError(f"Awaited coroutine for '{word}' still lacks 'text'.")
                        raise AttributeError("translation_obj lacks 'text' attribute")

                    return translation_obj.text
                
                except AttributeError as ae: 
                    logging.warning(f"AttributeError during translation for '{word}' (attempt {attempt+1}): {ae}. Re-initializing translator.")
                    # Re-initialize translator, as its internal state might be problematic
                    self.translator = Translator() 
                    await asyncio.sleep(0.5 * (attempt + 1))

                except Exception as e: # Catch other exceptions like network issues, ValueError from above
                    logging.error(f"Translation error for '{word}' to '{target_lang_code}' (attempt {attempt+1}): {e}")
                    if "TooManyRequests" in str(e) or "429" in str(e):
                        await asyncio.sleep(2 * (attempt + 1)) 
                    else:
                        await asyncio.sleep(0.5 * (attempt + 1))
            
            return "Translation Failed"

        except Exception as e:
            logging.error(f"Outer translation processing error for '{word}' to '{target_lang_code}': {e}")
            return "Error"

    def _extract_words(self, text, lang_code):
        logging.debug(f"Extracting words from text for language: {lang_code}")
        if lang_code in ["zh-cn", "ja", "ko"]:
            cjk_chars = r'\u2E80-\u2FFF\u3040-\u309F\u30A0-\u30FF\u31F0-\u31FF\u3200-\u32FF\u3400-\u4DBF\u4E00-\u9FFF\uAC00-\uD7AF\uF900-\uFAFF'
            mixed_pattern = rf'[{cjk_chars}a-zA-Z0-9]+'
            potential_words = re.findall(mixed_pattern, text)
        else:
            potential_words = re.split(r'\s+', text)
        min_len = 1 if lang_code in ["zh-cn", "ja", "ko"] else 2
        words = [word for word in potential_words if word and len(word) >= min_len]
        words = [
            word.lower()  # <<< MAKE EACH WORD LOWERCASE HERE
            for word in potential_words 
            if word and len(word) >= min_len
        ]
        logging.debug(f"Extracted {len(words)} words. First few: {words[:10]}")
        return words

    async def process_files(self):
        if not self.file_paths:
            logging.info("Process files called without file paths.")
            return

        logging.info(f"Processing files: {self.file_paths}")
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = 100
        self.progress_bar.update()
        all_words = []
        current_input_lang = self.language_var.get()

        try:
            for index, path in enumerate(self.file_paths):
                logging.debug(f"Reading file: {path}")
                with open(path, "r", encoding="utf-8") as file: text = file.read()
                words = self._extract_words(text, current_input_lang)
                all_words.extend(words)
                self.progress_bar["value"] = int(((index + 1) / len(self.file_paths)) * 50)
                self.progress_bar.update()
        except Exception as e:
            logging.error(f"Could not process files: {e}")
            messagebox.showerror("Error", f"Could not process files:\n{e}")
            return

        if not all_words:
            messagebox.showinfo("Info", "No words extracted from the selected files.")
            self.progress_bar["value"] = 100; self.progress_bar.update()
            return
            
        word_counts = Counter(all_words)
        sorted_words_with_counts = word_counts.most_common()

        for item in self.result_tree.get_children(): self.result_tree.delete(item)

        target_lang_map = {"english": "en", "arabic": "ar", "german": "de", "spanish": "es", "french": "fr", "italian": "it", "portuguese": "pt"}
        target_lang_name = self.translation_var.get().lower()
        target_lang_code = target_lang_map.get(target_lang_name, "None")
        
        try: word_limit = int(self.word_limit_entry.get()); word_limit = max(0, word_limit)
        except ValueError: messagebox.showerror("Error", "Invalid word limit. Using 0 (no limit)."); word_limit = 0
        
        words_to_display = sorted_words_with_counts[:word_limit] if word_limit > 0 else sorted_words_with_counts

        self.progress_bar["maximum"] = len(words_to_display) if words_to_display else 1
        self.progress_bar["value"] = 0
        
        logging.info(f"Displaying {len(words_to_display)} words. Translation target: {target_lang_code}")
        for index, (word, count) in enumerate(words_to_display):
            translation = await self._translate_word(word, target_lang_code) if target_lang_code != "None" and self.translator else ""
            self.result_tree.insert("", "end", values=(word, count, translation, "🔊"))
            if index % 10 == 0: self.progress_bar["value"] = index + 1; self.root.update_idletasks()

        self.progress_bar["value"] = len(words_to_display)
        self.progress_bar.update()
        messagebox.showinfo("Processing Complete", f"Displayed {len(words_to_display)} words (out of {len(sorted_words_with_counts)} unique words found).")
        logging.info("File processing and display complete.")

    def text_to_speech(self, text, lang='en', filename='output.mp3'):
        if not text.strip(): logging.warning("Attempted to synthesize empty text."); return None
        try:
            logging.debug(f"gTTS: text='{text}', lang='{lang}', file='{filename}'")
            tts = gtts.gTTS(text=text, lang=lang)
            tts.save(filename)
            logging.info(f"Audio saved to {filename}")
            return filename
        except Exception as e:
            logging.error(f"Error during gTTS speech generation for '{text}': {e}")
            return None

    def play_audio(self, filename):
        if not pygame.mixer.get_init():
            logging.warning("Pygame mixer not initialized. Attempting to re-initialize.")
            try: pygame.mixer.init()
            except pygame.error as e: logging.error(f"Failed to re-initialize pygame mixer: {e}"); messagebox.showerror("Audio Error", f"Pygame mixer is not available: {e}"); return
        try:
            logging.info(f"Playing audio: {filename}")
            pygame.mixer.music.load(filename)
            pygame.mixer.music.play()
        except pygame.error as e:
            logging.error(f"Error during pygame audio playback: {e}")
            messagebox.showerror("Playback Error", f"Could not play audio '{os.path.basename(filename)}':\n{e}")

    def treeview_click(self, event):
        region = self.result_tree.identify_region(event.x, event.y)
        if region == "cell":
            item_id = self.result_tree.identify_row(event.y)
            column_idx = self.result_tree.identify_column(event.x)
            if column_idx == '#4' and item_id: 
                values = self.result_tree.item(item_id)['values']
                self.speak_word(values[0])

    def speak_word(self, word):
        if word:
            temp_audio_dir = "temp_audio_files"; os.makedirs(temp_audio_dir, exist_ok=True)
            safe_fn = re.sub(r'[^\w\s-]', '', word).strip().replace(' ', '_') or "audio"
            # Add a timestamp or unique ID to avoid overwriting if words are similar after sanitizing
            output_file = os.path.join(temp_audio_dir, f"{safe_fn}_{int(time.time()*1000)}.mp3") 
            audio_file = self.text_to_speech(word, self.language_var.get(), output_file)
            if audio_file: self.play_audio(audio_file)
    
    def speak_all_words(self):
        words = [self.result_tree.item(item)['values'][0] for item in self.result_tree.get_children()]
        if not words: messagebox.showinfo("Info", "No words in the list to speak."); return
        def speak_sequentially(word_list, index=0):
            if index < len(word_list):
                word_to_speak = word_list[index]
                self.speak_word(word_to_speak)
                delay_ms = 1500 + len(word_to_speak) * 80 
                self.root.after(delay_ms, lambda: speak_sequentially(word_list, index + 1))
            else: logging.info("Finished speaking all words.")
        speak_sequentially(words)

    def export_anki_deck(self):
        if not self.result_tree.get_children():
            messagebox.showerror("Error", "No words to export. Please process files first.")
            return

        deck_name = self.deck_name_entry.get().strip()
        if not deck_name:
            messagebox.showerror("Error", "Please enter a deck name.")
            return
        
        model_name = "Vocabulary Card Model (Autoplay Audio)"
        export_type = self.export_var.get()

        model_id = random.randrange(1 << 30, 1 << 31)
        deck_id = random.randrange(1 << 30, 1 << 31)

        fields = [{"name": "Front"}, {"name": "Back"}]
        if "speech" in export_type:
            fields.append({"name": "Audio"})

        qfmt = '<div style="text-align: center; font-size: 24px;"><b>{{Front}}</b></div>'
        afmt_parts = [
            '<div style="text-align: center; font-size: 20px;">{{Front}}</div>',
            '<hr id="answer">',
            '<div style="text-align: center; font-size: 22px; margin-top:10px;">{{Back}}</div>'
        ]
        if "speech" in export_type:
            afmt_parts.append('{{#Audio}}')
            afmt_parts.append('<div id="anki-audio-player" style="text-align: center; margin-top:15px;">{{Audio}}</div>')
            afmt_parts.append('''
                <script>
                    var audioContainer = document.getElementById("anki-audio-player");
                    if (audioContainer) {
                        var audioEle = audioContainer.querySelector("audio");
                        if (audioEle && audioEle.paused) {
                            var playPromise = audioEle.play();
                            if (playPromise !== undefined) {
                                playPromise.catch(error => { console.log("Autoplay prevented: " + error); });
                            }
                        }
                    }
                </script>
            ''')
            afmt_parts.append('{{/Audio}}')
        
        model_css = (".card { font-family: Arial, sans-serif; background-color: #F0F0F0; color: #333; } "
                     "hr#answer { border-top: 1px solid #CCC; margin: 10px 0; } "
                     ".nightMode .card { background-color: #333; color: #F0F0F0; } "
                     ".nightMode hr#answer { border-top: 1px solid #555; }"
                     "#anki-audio-player audio { max-width: 100%; }")

        model = genanki.Model(model_id, model_name, fields=fields, 
                              templates=[{"name": "Card 1", "qfmt": qfmt, "afmt": "\n".join(afmt_parts)}],
                              css=model_css)
        deck = genanki.Deck(deck_id, deck_name)
        package = genanki.Package(deck)

        filepath = filedialog.asksaveasfilename(defaultextension=".apkg", 
                                               filetypes=[("Anki Package", "*.apkg")],
                                               title="Save Anki Deck As", initialfile=f"{deck_name}.apkg")
        if not filepath:
            return

        audio_dir_selected = None
        if "speech" in export_type:
            audio_dir_selected = filedialog.askdirectory(title="Select Folder to *Temporarily* Save Audio Files for Packaging")
            if not audio_dir_selected: 
                messagebox.showinfo("Info", "Audio export requires a temporary folder for audio files. Aborting export.")
                return
            os.makedirs(audio_dir_selected, exist_ok=True)

        total_items = len(self.result_tree.get_children())
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = total_items
        self.progress_bar.update()
        media_filenames_added = set()

        for index, item_id in enumerate(self.result_tree.get_children()):
            raw_values = self.result_tree.item(item_id)['values']
            
            # Ensure word and translation are strings, as Treeview might convert numeric-like strings
            word = str(raw_values[0])
            # raw_values[1] is count, which we don't directly use here for card content
            translation = str(raw_values[2]) 
            # raw_values[3] is the "Speak" icon text

            front_content, back_content, audio_anki_tag = "", "", ""
            
            safe_fn_base = re.sub(r'[^\w-]', '', word).strip().replace(' ', '_')
            if not safe_fn_base: 
                safe_fn_base = f"audio_{index}" 
            
            unique_suffix = str(random.randint(10000, 99999)) 
            audio_mp3_fn = f"{safe_fn_base}_{unique_suffix}.mp3"

            if "speech" in export_type and audio_dir_selected:
                full_audio_path = os.path.join(audio_dir_selected, audio_mp3_fn)
                gen_path = self.text_to_speech(word, self.language_var.get(), full_audio_path)
                if gen_path and os.path.exists(gen_path):
                    audio_anki_tag = f"[sound:{audio_mp3_fn}]"
                    if audio_mp3_fn not in media_filenames_added:
                        package.media_files.append(gen_path)
                        media_filenames_added.add(audio_mp3_fn)
                else:
                    logging.warning(f"Audio generation/finding failed for '{word}' (expected at {full_audio_path})")

            if export_type == "word_front_translation_back":
                front_content, back_content = word, translation
            elif export_type == "translation_front_word_back":
                front_content, back_content = translation, word
            elif export_type == "word_front_speech_back":
                front_content, back_content = word, "" 
            elif export_type == "translation_front_speech_word_back":
                front_content, back_content = translation, word
            elif export_type == "word_front_speech_translation_back":
                front_content, back_content = word, translation
            
            note_fields = [front_content, back_content]
            if "speech" in export_type:
                note_fields.append(audio_anki_tag)
            
            deck.add_note(genanki.Note(model=model, fields=note_fields))
            
            if index % 5 == 0: # Update progress bar periodically
                self.progress_bar["value"] = index + 1
                self.root.update_idletasks()

        self.progress_bar["value"] = total_items
        self.progress_bar.update()

        try:
            package.write_to_file(filepath)
            messagebox.showinfo("Success", f"Anki deck '{os.path.basename(filepath)}' exported successfully!")
            try:
                if os.name == 'nt':
                    os.startfile(os.path.dirname(filepath))
                elif sys.platform == 'darwin':
                    subprocess.run(['open', os.path.dirname(filepath)], check=False)
                else:
                    subprocess.run(['xdg-open', os.path.dirname(filepath)], check=False)
            except Exception as e_open:
                logging.warning(f"Could not open explorer/finder window: {e_open}")
        except Exception as e:
            logging.error(f"Error writing Anki package: {e}")
            messagebox.showerror("Export Error", f"Could not generate Anki Deck:\n{e}")
        
        self.progress_bar["value"] = 0
        self.progress_bar.update()

    def copy_selected_words(self, event=None):
        selected_items = self.result_tree.selection()
        if selected_items:
            words_to_copy = [self.result_tree.item(item)['values'][0] for item in selected_items]
            try:
                pyperclip.copy("\n".join(words_to_copy))
                messagebox.showinfo("Copied", f"{len(words_to_copy)} word(s) copied to clipboard.")
            except pyperclip.PyperclipException as e:
                logging.error(f"Error copying to clipboard: {e}")
                messagebox.showerror("Clipboard Error", f"Could not copy to clipboard:\n{e}\nEnsure a copy/paste utility (e.g., xclip/xsel on Linux) is installed.")
        else:
            messagebox.showinfo("Info", "No words selected to copy.")

if __name__ == "__main__":
    # Required for googletrans on Windows with asyncio in some contexts
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

    root = tk.Tk()
    app = WordCounterApp(root)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        print("Application exiting...")
    finally:
        if pygame.mixer.get_init(): pygame.mixer.quit()
        temp_dir_speak = "temp_audio_files" # For speak_word
        if os.path.exists(temp_dir_speak) and os.listdir(temp_dir_speak):
             logging.info(f"Temporary audio files from 'Speak Word' may exist in '{temp_dir_speak}'. You can manually delete this folder.")
        # The folder selected during export is managed by the user.
