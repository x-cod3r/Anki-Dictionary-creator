import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from collections import Counter
import re
import asyncio
from googletrans import Translator 
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
import inspect

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class WordCounterApp:
    def __init__(self, root, loop):
        self.root = root
        self.loop = loop 
        self.root.title("Anki dictionary creator by insta: @Mahmoud.aboulnasr")
        self.root.geometry("750x550")

        try:
            pygame.mixer.init()
        except pygame.error as e:
            logging.warning(f"Pygame mixer could not be initialized: {e}.")

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
        self.translator = None
        try:
            self.translator = Translator()
        except Exception as e:
            logging.error(f"Failed to initialize Translator: {e}")
            messagebox.showerror("Translator Error", f"Failed to initialize Google Translator: {e}")

        self.browse_button = ttk.Button(frame, text="Select File(s)", command=self.browse_files)
        self.browse_button.grid(row=1, column=0, sticky="ew", padx=2)

        self.process_button = ttk.Button(frame, text="Process Files", command=self.start_processing_files)
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
        self.translation_var.trace_add("write", lambda *args: self.start_processing_files() if self.file_paths and not self.is_processing else None)

        self.word_limit_label = ttk.Label(frame, text="Word Limit:")
        self.word_limit_label.grid(row=2, column=0, sticky="w", pady=5, padx=2)
        self.word_limit_entry = ttk.Entry(frame, width=7)
        self.word_limit_entry.grid(row=2, column=1, sticky="ew", pady=5, padx=2)
        self.word_limit_entry.insert(0, "50")

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

        self.is_processing = False
        self.words_to_process_list = []
        self.current_processing_index = 0
        self.total_words_for_progress = 0
        self.processed_count = 0
        self.target_lang_map = {"english": "en", "arabic": "ar", "german": "de", "spanish": "es", "french": "fr", "italian": "it", "portuguese": "pt"}
        self.current_target_lang_name = ""
        self.current_target_lang_code = ""
        self.words_to_speak_queue = []
        self._speak_job_id = None
        self.loop_manager = None


    def browse_files(self):
        if self.is_processing:
            messagebox.showinfo("Busy", "Cannot browse files while processing.")
            return
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

    async def _translate_words_batch(self, words_batch, target_lang_code):
        if not self.translator or not words_batch:
            return [""] * len(words_batch)
        
        original_indices_map = {} 
        non_empty_words = []
        for i, word in enumerate(words_batch):
            if word and str(word).strip():
                original_indices_map[len(non_empty_words)] = i
                non_empty_words.append(str(word))
        
        translated_results = [""] * len(words_batch)

        if not non_empty_words:
            return translated_results

        try:
            for attempt in range(3):
                try:
                    logging.debug(f"Attempting batch translation for {len(non_empty_words)} words (attempt {attempt+1}).")
                    
                    translation_call_result = await asyncio.to_thread(
                        self.translator.translate, non_empty_words, dest=target_lang_code
                    )

                    if inspect.iscoroutine(translation_call_result):
                        logging.warning("asyncio.to_thread returned a coroutine for batch translate. Awaiting it.")
                        translation_objs = await translation_call_result
                    else:
                        translation_objs = translation_call_result

                    if not isinstance(translation_objs, list):
                        logging.error(f"Batch translation (after potential await) did not return a list. Got: {type(translation_objs)}. Attempting re-init.")
                        raise ValueError("Translator returned non-list for batch.")

                    all_items_translated_successfully = True
                    for i, trans_obj in enumerate(translation_objs):
                        original_batch_idx = original_indices_map[i] 
                        
                        if inspect.iscoroutine(trans_obj): 
                            logging.warning(f"Individual item in batch result is a coroutine for '{non_empty_words[i]}'. Awaiting.")
                            try:
                                actual_item_result = await trans_obj
                                if hasattr(actual_item_result, 'text'):
                                    translated_results[original_batch_idx] = actual_item_result.text
                                else:
                                    logging.warning(f"Awaited item for '{non_empty_words[i]}' lacks .text. Type: {type(actual_item_result)}")
                                    translated_results[original_batch_idx] = "Item Error (No Text)"
                                    all_items_translated_successfully = False
                            except Exception as e_await_item:
                                logging.error(f"Error awaiting individual item coroutine '{non_empty_words[i]}': {e_await_item}")
                                translated_results[original_batch_idx] = "Item Await Error"
                                all_items_translated_successfully = False
                        elif trans_obj and hasattr(trans_obj, 'text'):
                            translated_results[original_batch_idx] = trans_obj.text
                        else: 
                            logging.warning(f"Batch translation item for '{non_empty_words[i]}' is problematic. Type: {type(trans_obj)}, Value: {trans_obj}")
                            translated_results[original_batch_idx] = "Item Invalid"
                            all_items_translated_successfully = False
                    
                    if all_items_translated_successfully:
                        return translated_results 
                    else:
                        logging.warning(f"Not all items translated successfully in attempt {attempt+1}. Retrying batch.")
                        raise ValueError("Partial success in batch, retrying.")

                except (AttributeError, ValueError) as e_val_attr: 
                    logging.warning(f"Error during batch processing (attempt {attempt+1}): {e_val_attr}. Re-initializing translator.")
                    try:
                        self.translator = Translator() 
                    except Exception as e_init_trans:
                        logging.error(f"Failed to re-initialize translator: {e_init_trans}")
                        for orig_idx in original_indices_map.values(): translated_results[orig_idx] = "Translator Re-init Err"
                        return translated_results 
                    await asyncio.sleep(1 * (attempt + 1)) 

                except Exception as e_general: 
                    logging.error(f"General batch translation error (attempt {attempt+1}) for '{target_lang_code}': {e_general}")
                    if "TooManyRequests" in str(e_general) or "429" in str(e_general):
                        await asyncio.sleep(3 * (attempt + 1)) 
                    else:
                        await asyncio.sleep(1.5 * (attempt + 1))
            
            logging.error(f"All {attempt+1} translation attempts failed for a batch.")
            for orig_idx in original_indices_map.values():
                if not translated_results[orig_idx]: 
                     translated_results[orig_idx] = "Translation Failed"
            return translated_results

        except Exception as e_outer: 
            logging.critical(f"CRITICAL error in _translate_words_batch structure: {e_outer}", exc_info=True)
            for orig_idx in original_indices_map.values():
                translated_results[orig_idx] = "Critical Error"
            return translated_results

    def _extract_words(self, text, lang_code):
        logging.debug(f"Extracting words from text for language: {lang_code}")
        if lang_code in ["zh-cn", "ja", "ko"]:
            cjk_chars = r'\u2E80-\u2FFF\u3040-\u309F\u30A0-\u30FF\u31F0-\u31FF\u3200-\u32FF\u3400-\u4DBF\u4E00-\u9FFF\uAC00-\uD7AF\uF900-\uFAFF'
            mixed_pattern = rf'[{cjk_chars}a-zA-Z0-9]+'
            potential_words = re.findall(mixed_pattern, text)
        else:
            potential_words = re.split(r'\s+', text)
        
        min_len = 1 if lang_code in ["zh-cn", "ja", "ko"] else 2
        words = [
            word.lower()
            for word in potential_words 
            if word and len(word.strip()) >= min_len and not word.strip().isdigit()
        ]
        logging.debug(f"Extracted {len(words)} words. First few: {words[:10]}")
        return words

    def start_processing_files(self):
        if self.is_processing:
            messagebox.showinfo("Info", "Processing is already in progress.")
            return
        if not self.file_paths:
            messagebox.showinfo("Info","No file(s) selected. Please select files first.")
            return

        self.is_processing = True
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = 100 
        self.root.update_idletasks()

        all_words = []
        current_input_lang = self.language_var.get()
        try:
            for index, path in enumerate(self.file_paths):
                with open(path, "r", encoding="utf-8") as file: text = file.read()
                words = self._extract_words(text, current_input_lang)
                all_words.extend(words)
                self.progress_bar["value"] = int(((index + 1) / len(self.file_paths)) * 20)
                self.root.update_idletasks()
        except Exception as e:
            logging.error(f"Could not process files: {e}")
            messagebox.showerror("Error", f"Could not process files:\n{e}")
            self.is_processing = False
            return

        if not all_words:
            messagebox.showinfo("Info", "No words extracted from the selected files.")
            self.progress_bar["value"] = 100; self.is_processing = False; self.root.update_idletasks(); return
            
        word_counts = Counter(all_words)
        sorted_words_with_counts = word_counts.most_common()

        try: 
            word_limit = int(self.word_limit_entry.get())
            word_limit = max(0, word_limit)
        except ValueError: 
            messagebox.showerror("Error", "Invalid word limit. Using 0 (no limit).")
            word_limit = 0
        
        self.words_to_process_list = sorted_words_with_counts[:word_limit] if word_limit > 0 else sorted_words_with_counts
        
        if not self.words_to_process_list:
            messagebox.showinfo("Info", "No words to display/translate based on limit.")
            self.is_processing = False
            self.progress_bar["value"] = 100; self.root.update_idletasks(); return

        self.total_words_for_progress = len(self.words_to_process_list)
        self.progress_bar["maximum"] = self.total_words_for_progress 
        self.progress_bar["value"] = 0
        self.processed_count = 0

        self.current_target_lang_name = self.translation_var.get().lower()
        self.current_target_lang_code = self.target_lang_map.get(self.current_target_lang_name, "None")

        logging.info(f"Starting translation/display for {self.total_words_for_progress} words. Translation target: {self.current_target_lang_code}")
        
        self.current_processing_index = 0
        if self.loop_manager:
            self.loop_manager.schedule_async_processing() 
        self.process_next_word_batch()


    def process_next_word_batch(self, batch_size=30): 
        if not self.is_processing or self.current_processing_index >= self.total_words_for_progress:
            if self.is_processing: self.finish_processing()
            return

        start_idx = self.current_processing_index
        end_idx = min(start_idx + batch_size, self.total_words_for_progress)
        current_batch_data = self.words_to_process_list[start_idx:end_idx]

        if not current_batch_data:
            self.finish_processing(); return

        words_only_batch = [str(item[0]) for item in current_batch_data]
        counts_batch = [item[1] for item in current_batch_data]
        
        try:
            asyncio.ensure_future(self.translate_batch_and_update_ui(words_only_batch, counts_batch), loop=self.loop)
        except RuntimeError as e:
             logging.error(f"RuntimeError creating task: {e}", exc_info=True)
             try: 
                 current_loop_for_task = asyncio.get_event_loop()
                 asyncio.ensure_future(self.translate_batch_and_update_ui(words_only_batch, counts_batch), loop=current_loop_for_task)
             except Exception as e2:
                 logging.critical(f"CRITICAL: Failed to schedule async task: {e2}", exc_info=True)
                 messagebox.showerror("Critical Async Error", "Could not schedule background tasks. Please restart.")
                 self.finish_processing() 
                 return
        self.current_processing_index = end_idx

    async def translate_batch_and_update_ui(self, words_batch, counts_batch):
        translations_batch = [""] * len(words_batch)
        if self.current_target_lang_code != "None" and self.translator:
            translations_batch = await self._translate_words_batch(words_batch, self.current_target_lang_code)
        
        for i in range(len(words_batch)):
            word = words_batch[i]
            count = counts_batch[i]
            translation = translations_batch[i]
            self.result_tree.insert("", "end", values=(word, count, translation, "🔊"))

        self.processed_count += len(words_batch)
        self.processed_count = min(self.processed_count, self.total_words_for_progress)
        
        self.progress_bar["value"] = self.processed_count
        
        if self.current_processing_index < self.total_words_for_progress and self.is_processing:
            self.root.after(100, self.process_next_word_batch) 
        elif self.is_processing:
            self.finish_processing() 
        self.root.update_idletasks() 


    def finish_processing(self):
        if not self.is_processing: return 
        self.progress_bar["value"] = self.processed_count 
        self.root.update_idletasks()
        messagebox.showinfo("Processing Complete", f"Displayed {self.processed_count} of {self.total_words_for_progress} targeted words.")
        logging.info("File processing and display complete.")
        self.is_processing = False


    def text_to_speech(self, text, lang='en', filename='output.mp3'):
        if not text or not str(text).strip(): logging.warning("TTS: empty text."); return None
        try:
            logging.debug(f"gTTS: text='{text}', lang='{lang}', file='{filename}'")
            tts = gtts.gTTS(text=str(text), lang=lang)
            tts.save(filename)
            logging.info(f"Audio saved to {filename}")
            return filename
        except Exception as e:
            logging.error(f"Error during gTTS speech generation for '{text}': {e}")
            if "No text to send" in str(e) and (not text or not str(text).strip()):
                return None
            return None

    def play_audio(self, filename):
        if not pygame.mixer.get_init():
            logging.warning("Pygame mixer not initialized. Attempting to re-initialize.")
            try: pygame.mixer.init()
            except pygame.error as e: 
                logging.error(f"Failed to re-initialize pygame mixer: {e}")
                messagebox.showerror("Audio Error", f"Pygame mixer is not available: {e}")
                return
        if not filename or not os.path.exists(filename):
            logging.error(f"Audio file not found: {filename}")
            return
        try:
            logging.info(f"Playing audio: {filename}")
            pygame.mixer.music.load(filename)
            pygame.mixer.music.play()
        except pygame.error as e:
            logging.error(f"Error during pygame audio playback for {filename}: {e}")
            messagebox.showerror("Playback Error", f"Could not play audio '{os.path.basename(filename)}':\n{e}")

    def treeview_click(self, event):
        if self.is_processing: return 
        region = self.result_tree.identify_region(event.x, event.y)
        if region == "cell":
            item_id = self.result_tree.identify_row(event.y)
            column_idx = self.result_tree.identify_column(event.x)
            if column_idx == '#4' and item_id: 
                values = self.result_tree.item(item_id)['values']
                if values and len(values) > 0:
                    self.speak_word(str(values[0]))

    def speak_word(self, word_to_speak):
        word_to_speak_str = str(word_to_speak).strip()
        if not word_to_speak_str: return
        
        temp_audio_dir = "temp_audio_files"; os.makedirs(temp_audio_dir, exist_ok=True)
        safe_fn = re.sub(r'[^\w\s-]', '', word_to_speak_str).strip().replace(' ', '_') or "audio"
        output_file = os.path.join(temp_audio_dir, f"{safe_fn}_{int(time.time()*1000)}.mp3") 
        
        audio_file = self.text_to_speech(word_to_speak_str, self.language_var.get(), output_file)
        if audio_file: self.play_audio(audio_file)
    
    def speak_all_words(self):
        if self.is_processing:
            messagebox.showinfo("Busy", "Cannot speak all words while processing files.")
            return
        
        if self._speak_job_id:
            self.root.after_cancel(self._speak_job_id)
            self._speak_job_id = None
        self.words_to_speak_queue.clear()

        visible_items = self.result_tree.get_children()
        for item_id in visible_items:
            item_values = self.result_tree.item(item_id)['values']
            if item_values and len(item_values) > 0:
                self.words_to_speak_queue.append(str(item_values[0]))

        if not self.words_to_speak_queue: messagebox.showinfo("Info", "No words in the list to speak."); return
        
        self._speak_next_from_queue()

    def _speak_next_from_queue(self):
        if not self.words_to_speak_queue:
            logging.info("Finished speaking all words from queue.")
            self._speak_job_id = None
            return
        
        if pygame.mixer.get_init() and pygame.mixer.music.get_busy():
            self._speak_job_id = self.root.after(200, self._speak_next_from_queue)
            return

        word_to_speak = self.words_to_speak_queue.pop(0)
        self.speak_word(word_to_speak)
        
        delay_ms = 1000 + len(word_to_speak) * 80 
        self._speak_job_id = self.root.after(delay_ms, self._speak_next_from_queue)


    def export_anki_deck(self):
        if self.is_processing:
            messagebox.showinfo("Busy", "Cannot export while processing files.")
            return
        if not self.result_tree.get_children():
            messagebox.showerror("Error", "No words to export. Please process files first.")
            return

        deck_name = self.deck_name_entry.get().strip()
        if not deck_name: messagebox.showerror("Error", "Please enter a deck name."); return
        
        model_name = "Vocabulary Card Model (Autoplay Audio)"
        export_type = self.export_var.get()

        model_id = random.randrange(1 << 30, 1 << 31)
        deck_id = random.randrange(1 << 30, 1 << 31)

        fields = [{"name": "Front"}, {"name": "Back"}]
        if "speech" in export_type: fields.append({"name": "Audio"})

        qfmt = '<div style="text-align: center; font-size: 24px;"><b>{{Front}}</b></div>'
        afmt_parts = [
            '<div style="text-align: center; font-size: 20px;">{{Front}}</div>',
            '<hr id="answer">',
            '<div style="text-align: center; font-size: 22px; margin-top:10px;">{{Back}}</div>'
        ]
        if "speech" in export_type:
            afmt_parts.extend(['{{#Audio}}',
                               '<div id="anki-audio-player" style="text-align: center; margin-top:15px;">{{Audio}}</div>',
                               '''<script>
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
                               </script>''',
                               '{{/Audio}}'])
        
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
        if not filepath: return

        audio_dir_selected = None
        if "speech" in export_type:
            audio_dir_selected = filedialog.askdirectory(title="Select Folder to *Temporarily* Save Audio Files for Packaging")
            if not audio_dir_selected: 
                messagebox.showinfo("Info", "Audio export requires a temporary folder. Aborting export.")
                return
            os.makedirs(audio_dir_selected, exist_ok=True)

        items_to_export = self.result_tree.get_children()
        total_items = len(items_to_export)
        self.progress_bar["value"] = 0; self.progress_bar["maximum"] = total_items; self.progress_bar.update()
        media_filenames_added = set()

        for index, item_id in enumerate(items_to_export):
            raw_values = self.result_tree.item(item_id)['values']
            word = str(raw_values[0])
            translation = str(raw_values[2]) if len(raw_values) > 2 else ""
            
            front_content, back_content, audio_anki_tag = "", "", ""
            
            safe_fn_base = re.sub(r'[^\w-]', '', word).strip().replace(' ', '_')
            if not safe_fn_base: safe_fn_base = f"audio_{index}"
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
                else: logging.warning(f"Audio generation/finding failed for '{word}'")

            if export_type == "word_front_translation_back": front_content, back_content = word, translation
            elif export_type == "translation_front_word_back": front_content, back_content = translation, word
            elif export_type == "word_front_speech_back": front_content, back_content = word, "" 
            elif export_type == "translation_front_speech_word_back": front_content, back_content = translation, word
            elif export_type == "word_front_speech_translation_back": front_content, back_content = word, translation
            
            note_fields = [front_content, back_content]
            if "speech" in export_type: note_fields.append(audio_anki_tag)
            deck.add_note(genanki.Note(model=model, fields=note_fields))
            
            if index % 5 == 0 or index == total_items -1 : 
                self.progress_bar["value"] = index + 1
                self.root.update_idletasks()

        self.progress_bar["value"] = total_items; self.progress_bar.update()

        try:
            package.write_to_file(filepath)
            messagebox.showinfo("Success", f"Anki deck '{os.path.basename(filepath)}' exported successfully!")
            try:
                if os.name == 'nt': os.startfile(os.path.dirname(filepath))
                elif sys.platform == 'darwin': subprocess.run(['open', os.path.dirname(filepath)], check=False)
                else: subprocess.run(['xdg-open', os.path.dirname(filepath)], check=False)
            except Exception as e_open: logging.warning(f"Could not open explorer: {e_open}")
        except Exception as e:
            logging.error(f"Error writing Anki package: {e}")
            messagebox.showerror("Export Error", f"Could not generate Anki Deck:\n{e}")
        
        self.progress_bar["value"] = 0; self.progress_bar.update()

    def copy_selected_words(self, event=None):
        if self.is_processing: return
        selected_items = self.result_tree.selection()
        if selected_items:
            words_to_copy = [str(self.result_tree.item(item)['values'][0]) for item in selected_items]
            try:
                pyperclip.copy("\n".join(words_to_copy))
                messagebox.showinfo("Copied", f"{len(words_to_copy)} word(s) copied to clipboard.")
            except pyperclip.PyperclipException as e:
                logging.error(f"Error copying to clipboard: {e}")
                messagebox.showerror("Clipboard Error", f"Could not copy: {e}")
        else:
            messagebox.showinfo("Info", "No words selected to copy.")

class AsyncTkinterLoopManager:
    def __init__(self, tk_root, async_loop):
        self.tk_root = tk_root
        self.async_loop = async_loop
        self._after_id = None
        self._is_tkinter_driver_running = False 

    def _drive_async_loop(self):
        if not self._is_tkinter_driver_running:
            return
        try:
            # This runs tasks that are ready and yields control quickly.
            # It's important that this doesn't block indefinitely if the loop is stopping.
            self.async_loop.call_soon(self.async_loop.stop) # Stop run_forever after this round
            self.async_loop.run_forever() # Run all tasks currently in the queue
        except asyncio.CancelledError:
            logging.debug("AsyncTkinterLoopManager._drive_async_loop's internal run was cancelled.")
            # If cancelled, we should not reschedule via 'finally' if _is_tkinter_driver_running was set to False.
        except RuntimeError as e: # Catch "Event loop is closed" or "cannot call run_forever"
            if "Event loop is closed" in str(e) or "cannot call run_forever" in str(e).lower():
                logging.debug(f"Loop was closed or cannot run in _drive_async_loop: {e}")
                self._is_tkinter_driver_running = False # Stop driving if loop is closed/unrunnable
            else: # Re-raise other RuntimeErrors
                logging.error(f"Unexpected RuntimeError in _drive_async_loop: {e}", exc_info=True)
                # self._is_tkinter_driver_running = False # Consider stopping on unexpected errors too
        except Exception as e:
            logging.error(f"Unexpected Exception in _drive_async_loop: {e}", exc_info=True)
            # self._is_tkinter_driver_running = False # Consider stopping

        if self._is_tkinter_driver_running: 
            self._after_id = self.tk_root.after(10, self._drive_async_loop)


    def start(self):
        if not self._is_tkinter_driver_running:
            self._is_tkinter_driver_running = True
            try:
                asyncio.get_running_loop()
            except RuntimeError: 
                asyncio.set_event_loop(self.async_loop)
            
            self._after_id = self.tk_root.after(10, self._drive_async_loop)
            logging.info("AsyncTkinterLoopManager started.")

    def schedule_async_processing(self):
        pass 

    def stop_event_loop_integration(self):
        logging.info("Stopping AsyncTkinterLoopManager integration with Tkinter.")
        self._is_tkinter_driver_running = False 
        if self._after_id:
            if self.tk_root.winfo_exists(): 
                try:
                    self.tk_root.after_cancel(self._after_id)
                except tk.TclError: 
                    logging.debug("TclError cancelling _after_id, root likely destroyed.")
            self._after_id = None
        
        # If the asyncio loop's run_forever was started by _drive_async_loop and is still running,
        # it should be stopped. The call_soon(self.async_loop.stop) in _drive_async_loop aims for this.
        if self.async_loop.is_running():
            self.async_loop.call_soon_threadsafe(self.async_loop.stop)


    async def graceful_shutdown_async_tasks(self):
        logging.info("Attempting graceful shutdown of asyncio tasks.")
        try:
            asyncio.get_running_loop()
        except RuntimeError:
            asyncio.set_event_loop(self.async_loop)

        tasks = [t for t in asyncio.all_tasks(loop=self.async_loop) if t is not asyncio.current_task(loop=self.async_loop)]
        if tasks:
            logging.info(f"Cancelling {len(tasks)} outstanding asyncio tasks.")
            for task in tasks:
                task.cancel()
            
            results = await asyncio.gather(*tasks, return_exceptions=True)
            for i, result in enumerate(results):
                task_name = tasks[i].get_name() if hasattr(tasks[i], 'get_name') else f"Task-{i}"
                if isinstance(result, asyncio.CancelledError):
                    logging.debug(f"{task_name} was cancelled.")
                elif isinstance(result, Exception):
                    logging.error(f"{task_name} raised an exception during shutdown: {result}")
            logging.info("Asyncio tasks cancellation processed.")
        else:
            logging.info("No outstanding asyncio tasks to cancel.")


_is_shutting_down_flag = False 

async def main_shutdown_sequence(event_loop, tk_root, loop_mgr):
    global _is_shutting_down_flag
    
    logging.info("Main shutdown sequence initiated (async part).")
    
    if loop_mgr:
        loop_mgr.stop_event_loop_integration() 
        await loop_mgr.graceful_shutdown_async_tasks() 

    if pygame.mixer.get_init(): pygame.mixer.quit()
    temp_dir_speak = "temp_audio_files"
    if os.path.exists(temp_dir_speak) and os.listdir(temp_dir_speak):
        logging.info(f"Temporary audio files from 'Speak Word' may exist in '{temp_dir_speak}'.")
    
    if tk_root and tk_root.winfo_exists():
        tk_root.destroy() 

    if event_loop:
        try:
            # This ensures that any tasks spawned by `shutdown_asyncgens` can run.
            async def _shutdown_gens_and_stop():
                if hasattr(event_loop, 'shutdown_asyncgens') and callable(event_loop.shutdown_asyncgens):
                    await event_loop.shutdown_asyncgens()
                if event_loop.is_running():
                    event_loop.stop() # Stop after gens are shut down

            if not event_loop.is_closed():
                if not event_loop.is_running(): # If stopped, need to run it briefly
                    try:
                        event_loop.run_until_complete(_shutdown_gens_and_stop())
                    except RuntimeError as e: # May happen if loop can't be restarted
                        logging.debug(f"Could not run final shutdown tasks: {e}")
                else: # If running, schedule it (shouldn't be running if loop_mgr stopped)
                    await _shutdown_gens_and_stop()

            if not event_loop.is_closed():
                logging.info("Closing asyncio event loop.")
                event_loop.close()
            logging.info("Asyncio loop closed status: %s", event_loop.is_closed())
        except Exception as e_loop_close:
            logging.error(f"Error during asyncio loop cleanup: {e_loop_close}", exc_info=True)

    logging.info("Application shutdown sequence complete.")
    _is_shutting_down_flag = True # Mark as fully shut down


if __name__ == "__main__":
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    
    main_event_loop = None
    try:
        main_event_loop = asyncio.get_event_loop()
        if main_event_loop.is_closed():
            main_event_loop = asyncio.new_event_loop()
            asyncio.set_event_loop(main_event_loop)
    except RuntimeError: 
        main_event_loop = asyncio.new_event_loop()
        asyncio.set_event_loop(main_event_loop)

    root_tk = tk.Tk()
    app_instance = WordCounterApp(root_tk, main_event_loop) 
    loop_manager_instance = AsyncTkinterLoopManager(root_tk, main_event_loop)
    app_instance.loop_manager = loop_manager_instance

    def on_wm_delete_window_wrapper():
        global _is_shutting_down_flag
        if _is_shutting_down_flag:
            logging.debug("WM_DELETE_WINDOW: Shutdown already in progress.")
            return 

        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            logging.info("WM_DELETE_WINDOW: Quit confirmed. Initiating shutdown.")
            _is_shutting_down_flag = True 
            try:
                asyncio.set_event_loop(main_event_loop) # Ensure correct loop context
                # Schedule the main_shutdown_sequence. This will handle destroying root.
                asyncio.ensure_future(main_shutdown_sequence(main_event_loop, root_tk, loop_manager_instance), loop=main_event_loop)
            except Exception as e: 
                logging.error(f"Failed to schedule shutdown from WM_DELETE_WINDOW: {e}. Forcing Tk destroy.")
                if root_tk.winfo_exists(): root_tk.destroy() 
        else: # User cancelled quit
            _is_shutting_down_flag = False # Reset flag

    root_tk.protocol("WM_DELETE_WINDOW", on_wm_delete_window_wrapper)
    
    try:
        loop_manager_instance.start() 
        root_tk.mainloop() 
    except KeyboardInterrupt:
        logging.info("KeyboardInterrupt caught. Initiating shutdown.")
        if not _is_shutting_down_flag:
            _is_shutting_down_flag = True 
            loop_manager_instance.stop_event_loop_integration() 
            if main_event_loop and not main_event_loop.is_closed():
                 main_event_loop.run_until_complete(main_shutdown_sequence(main_event_loop, root_tk, loop_manager_instance))
            elif root_tk.winfo_exists(): 
                root_tk.destroy()
    except SystemExit:
        logging.info("SystemExit caught, application likely shutting down through Tkinter.")
    except Exception as e_mainloop: 
        logging.critical(f"Unhandled exception in Tkinter mainloop: {e_mainloop}", exc_info=True)
        if not _is_shutting_down_flag: 
             _is_shutting_down_flag = True
             loop_manager_instance.stop_event_loop_integration()
             if main_event_loop and not main_event_loop.is_closed():
                main_event_loop.run_until_complete(main_shutdown_sequence(main_event_loop, root_tk, loop_manager_instance))
             elif root_tk.winfo_exists():
                root_tk.destroy()
    finally:
        logging.info("Tkinter mainloop has exited.")
        # If shutdown flag is not set, it means an abrupt exit not via WM_DELETE or KeyboardInterrupt
        # that was handled. We should try to run the shutdown sequence.
        if not _is_shutting_down_flag and main_event_loop and not main_event_loop.is_closed():
            logging.warning("Mainloop exited without standard shutdown. Attempting final cleanup.")
            _is_shutting_down_flag = True # Prevent re-entry from here
            # This assumes the loop can still run tasks.
            try:
                main_event_loop.run_until_complete(main_shutdown_sequence(main_event_loop, None, loop_manager_instance))
            except RuntimeError as e: # Loop might be closed or unrunnable
                logging.error(f"Error running final shutdown in finally: {e}")
                if not main_event_loop.is_closed(): main_event_loop.close()

        elif main_event_loop and not main_event_loop.is_closed():
             logging.info("Loop still open in final finally, but shutdown was initiated. Ensuring it's closed.")
             # This path is if _is_shutting_down_flag was true, but loop wasn't closed by main_shutdown_sequence
             # (e.g. if main_shutdown_sequence itself had an error before closing the loop)
             if main_event_loop.is_running(): main_event_loop.stop()
             main_event_loop.close()

        logging.info("Application __main__ finished.")
