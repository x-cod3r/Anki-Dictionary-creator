import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import zipfile
import rarfile # Requires 'unrar' command-line tool
import re
import threading # To prevent GUI freeze
import io # For BytesIO

# --- PDF and DOCX Libraries ---
try:
    from pdfminer.high_level import extract_text as pdf_extract_text
    from pdfminer.pdfparser import PDFSyntaxError
    PDFMINER_AVAILABLE = True
except ImportError:
    PDFMINER_AVAILABLE = False
    print("WARNING: pdfminer.six not installed. PDF extraction will be disabled. "
          "Install with: pip install pdfminer.six")

try:
    from docx import Document as DocxDocument
    from docx.opc.exceptions import PackageNotFoundError as DocxPackageNotFoundError
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False
    print("WARNING: python-docx not installed. DOCX extraction will be disabled. "
          "Install with: pip install python-docx")


# --- Configuration for rarfile ---
# rarfile.UNRAR_TOOL = "path/to/unrar" # Uncomment and set if unrar is not in PATH

class TextExtractorMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Text Extractor & Merger")
        self.root.geometry("700x550")

        self.selected_files = []

        # --- UI Elements ---
        self.frame_controls = tk.Frame(root, pady=10)
        self.frame_controls.pack(fill=tk.X)

        self.btn_select_files = tk.Button(self.frame_controls, text="Select Files/Archives (.zip, .rar, .pdf, .docx, .srt)", command=self.select_files)
        self.btn_select_files.pack(pady=5)
        self.info_label = tk.Label(self.frame_controls, text="(Extracts from archives OR directly from selected .pdf, .docx, .srt)")
        self.info_label.pack(pady=2)


        self.lbl_selected_count = tk.Label(self.frame_controls, text="Selected Items: 0")
        self.lbl_selected_count.pack(pady=2)

        self.btn_process = tk.Button(self.frame_controls, text="Extract, Clean & Merge Text", command=self.start_processing_thread, state=tk.DISABLED)
        self.btn_process.pack(pady=10)

        self.status_label = tk.Label(root, text="Status: Idle", relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

        self.log_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=20, state=tk.DISABLED)
        self.log_area.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        if not PDFMINER_AVAILABLE:
            self.log_message("WARNING: pdfminer.six library not found. PDF extraction is disabled.")
        if not DOCX_AVAILABLE:
            self.log_message("WARNING: python-docx library not found. DOCX extraction is disabled.")


    def log_message(self, message):
        self.log_area.config(state=tk.NORMAL)
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def update_status(self, message):
        self.status_label.config(text=f"Status: {message}")
        self.root.update_idletasks()

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Select Files or Archives",
            filetypes=(
                ("Supported Files", "*.zip *.rar *.pdf *.docx *.srt"),
                ("Archive files", "*.zip *.rar"),
                ("PDF files", "*.pdf"),
                ("Word documents", "*.docx"),
                ("Subtitle files", "*.srt"),
                ("All files", "*.*")
            )
        )
        if files:
            self.selected_files = list(files)
            self.lbl_selected_count.config(text=f"Selected Items: {len(self.selected_files)}")
            self.btn_process.config(state=tk.NORMAL)
            self.log_message(f"Selected {len(self.selected_files)} item(s):")
            for f in self.selected_files:
                self.log_message(f"  - {os.path.basename(f)}")
        else:
            self.log_message("No items selected.")

    def preprocess_srt_content(self, srt_content_raw):
        processed_lines = []
        lines = srt_content_raw.splitlines()
        for line in lines:
            line = line.strip()
            if not line or re.match(r'^\d+$', line) or '-->' in line:
                continue
            line = re.sub(r'<[^>]+>', '', line)
            if line:
                processed_lines.append(line)
        return "\n".join(processed_lines)

    def aggressive_word_cleaner(self, text_block):
         # Keep:
        # - Basic English letters (a-zA-Z)
        # - Numbers (0-9)
        # - Whitespace (\s)
        # - Apostrophe (') and Hyphen (-)
        # - Common European accented characters and specific punctuation
        #   (German: äöüÄÖÜß)
        #   (French: àâçéèêëîïôùûüÿÀÂÇÉÈÊËÎÏÔÙÛÜŸ)
        #   (Spanish: áéíóúüñÁÉÍÓÚÜÑ)
        #   (Italian: àèéìòóùÀÈÉÌÒÓÙ)
        #   (Portuguese: áàâãçéêíóôõúüÁÀÂÃÇÉÊÍÓÔÕÚÜ)
        #   (Common punctuation: . , ! ? ; : ¿ ¡ « » – —)
        #   This regex is getting long, but aims for broadness.
        #   It's often better to list what you *want to keep* than what you *want to remove*.

        # Consolidated set (may have some redundancy, but that's fine for a character class)
        if not text_block:
            return ""
        cleaned_lines = []
        for line in text_block.splitlines():
            allowed_chars = r"a-zA-Z0-9\s'" \
                        r"äöüÄÖÜß" \
                    #    r"àâæçéèêëîïôœùûüÿÀÂÆÇÉÈÊËÎÏÔŒÙÛÜŸ" \
                    #    r"áéíóúüñÁÉÍÓÚÜÑ" \
                    #    r"àèéìòóùÀÈÉÌÒÓÙ" \
                    #    r"áàâãçéêíóôõúüÁÀÂÃÇÉÊÍÓÔÕÚÜ" \
                    #    r"\-" \
                    #    r".,!?;:¿¡«»“”‘’" # Added more common punctuation

            # The regex becomes: "match anything NOT in the allowed_chars set"
            cleaned_line = re.sub(rf"[^{allowed_chars}]", '', line)

            # Normalize multiple spaces to one, and strip leading/trailing
            cleaned_line = re.sub(r'\s+', ' ', cleaned_line).strip()
            if cleaned_line:
                cleaned_lines.append(cleaned_line)
        return "\n".join(cleaned_lines)
    


    def extract_text_from_bytes(self, file_bytes, file_name_for_log):
        raw_text = None
        try:
            if file_name_for_log.lower().endswith('.srt'):
                try:
                    content_raw = file_bytes.decode('utf-8')
                except UnicodeDecodeError:
                    self.log_message(f"    UTF-8 decoding failed for {file_name_for_log}, trying latin-1...")
                    content_raw = file_bytes.decode('latin-1', errors='ignore')
                raw_text = self.preprocess_srt_content(content_raw)
                if raw_text:
                    self.log_message(f"    Preprocessed SRT: {file_name_for_log}")
                else:
                    self.log_message(f"    SRT {file_name_for_log} was empty after preprocessing.")

            elif PDFMINER_AVAILABLE and file_name_for_log.lower().endswith('.pdf'):
                pdf_file_like = io.BytesIO(file_bytes)
                try:
                    raw_text = pdf_extract_text(pdf_file_like)
                    self.log_message(f"    Extracted text from PDF: {file_name_for_log}")
                except PDFSyntaxError:
                    self.log_message(f"    Error: Invalid or password protected PDF: {file_name_for_log}")
                except Exception as e_pdf:
                    self.log_message(f"    Error extracting text from PDF {file_name_for_log}: {e_pdf}")

            elif DOCX_AVAILABLE and file_name_for_log.lower().endswith('.docx'):
                docx_file_like = io.BytesIO(file_bytes)
                try:
                    doc = DocxDocument(docx_file_like)
                    raw_text_parts = [para.text for para in doc.paragraphs]
                    raw_text = "\n".join(raw_text_parts)
                    self.log_message(f"    Extracted text from DOCX: {file_name_for_log}")
                except DocxPackageNotFoundError:
                     self.log_message(f"    Error: Not a valid DOCX (or corrupt): {file_name_for_log}")
                except Exception as e_docx:
                    self.log_message(f"    Error extracting text from DOCX {file_name_for_log}: {e_docx}")
            else:
                # This case should ideally not be hit if called correctly for supported types
                self.log_message(f"    Cannot extract text from {file_name_for_log} (unsupported for direct text extraction here).")
                return None
        except Exception as e_decode_extract:
            self.log_message(f"    Error decoding/extracting {file_name_for_log}: {e_decode_extract}")
            return None
        return raw_text

    def process_selected_items(self): # Renamed from process_archives
        if not self.selected_files:
            messagebox.showwarning("No Items", "Please select some files or archives first.")
            return

        all_final_cleaned_texts = []
        self.update_status("Starting extraction...")
        self.log_message("\n--- Processing Started ---")

        for i, item_path in enumerate(self.selected_files): # item_path can be archive or standalone file
            self.update_status(f"Processing {os.path.basename(item_path)} ({i+1}/{len(self.selected_files)})...")
            item_basename = os.path.basename(item_path)
            
            try:
                if item_path.lower().endswith('.zip'):
                    self.log_message(f"\nProcessing archive: {item_path}")
                    found_target_in_archive = False
                    with zipfile.ZipFile(item_path, 'r') as zf:
                        for member_info in zf.infolist():
                            if member_info.is_dir():
                                continue
                            member_name = member_info.filename
                            if any(member_name.lower().endswith(ext) for ext in ['.srt', '.pdf', '.docx']):
                                self.log_message(f"  Found target file in ZIP: {member_name}")
                                found_target_in_archive = True
                                file_bytes = zf.read(member_name)
                                raw_or_preprocessed_text = self.extract_text_from_bytes(file_bytes, member_name)
                                if raw_or_preprocessed_text:
                                    cleaned_text = self.aggressive_word_cleaner(raw_or_preprocessed_text)
                                    if cleaned_text:
                                        all_final_cleaned_texts.append(cleaned_text)
                                        self.log_message(f"    Successfully cleaned and added content from {member_name}.")
                                    else:
                                        self.log_message(f"    Content from {member_name} was empty after final cleaning.")
                    if not found_target_in_archive:
                        self.log_message(f"  No .srt, .pdf, or .docx files found in archive {item_basename}.")

                elif item_path.lower().endswith('.rar'):
                    self.log_message(f"\nProcessing archive: {item_path}")
                    found_target_in_archive = False
                    try:
                        with rarfile.RarFile(item_path, 'r') as rf:
                            for member_info in rf.infolist():
                                if member_info.isdir():
                                    continue
                                member_name = member_info.filename
                                if any(member_name.lower().endswith(ext) for ext in ['.srt', '.pdf', '.docx']):
                                    self.log_message(f"  Found target file in RAR: {member_name}")
                                    found_target_in_archive = True
                                    file_bytes = rf.read(member_info)
                                    raw_or_preprocessed_text = self.extract_text_from_bytes(file_bytes, member_name)
                                    if raw_or_preprocessed_text:
                                        cleaned_text = self.aggressive_word_cleaner(raw_or_preprocessed_text)
                                        if cleaned_text:
                                            all_final_cleaned_texts.append(cleaned_text)
                                            self.log_message(f"    Successfully cleaned and added content from {member_name}.")
                                        else:
                                            self.log_message(f"    Content from {member_name} was empty after final cleaning.")
                        if not found_target_in_archive:
                             self.log_message(f"  No .srt, .pdf, or .docx files found in archive {item_basename}.")
                    except rarfile.NeedFirstVolume:
                         self.log_message(f"Error: {item_path} is part of a multi-volume RAR. Please select the first volume.")
                    except rarfile.RarCannotExec as e_rar_exec:
                        self.log_message(f"Error with unrar: {e_rar_exec}. Ensure 'unrar' command-line tool is installed and in PATH.")
                        messagebox.showerror("Unrar Error", f"Could not execute unrar: {e_rar_exec}\n\nPlease ensure 'unrar' (the command-line utility) is installed and accessible.")
                        return
                    except Exception as e_rar:
                        self.log_message(f"Error opening/processing RAR {item_path}: {e_rar}")
                
                # Handle standalone PDF, DOCX, SRT files
                elif any(item_path.lower().endswith(ext) for ext in ['.pdf', '.docx', '.srt']):
                    self.log_message(f"\nProcessing standalone file: {item_path}")
                    try:
                        with open(item_path, 'rb') as f_standalone:
                            file_bytes = f_standalone.read()
                        
                        raw_or_preprocessed_text = self.extract_text_from_bytes(file_bytes, item_basename)
                        if raw_or_preprocessed_text:
                            cleaned_text = self.aggressive_word_cleaner(raw_or_preprocessed_text)
                            if cleaned_text:
                                all_final_cleaned_texts.append(cleaned_text)
                                self.log_message(f"  Successfully cleaned and added content from {item_basename}.")
                            else:
                                self.log_message(f"  Content from {item_basename} was empty after final cleaning.")
                        else:
                            self.log_message(f"  No text extracted or file empty for {item_basename}.")

                    except FileNotFoundError:
                         self.log_message(f"Error: Standalone file not found - {item_path}")
                    except Exception as e_standalone:
                         self.log_message(f"Error processing standalone file {item_path}: {e_standalone}")

                else:
                    self.log_message(f"\nSkipping unsupported file type: {item_path}")

            except FileNotFoundError: # For archives primarily
                self.log_message(f"Error: File or Archive not found - {item_path}")
            except Exception as e_outer:
                self.log_message(f"An unexpected error occurred with {item_path}: {e_outer}")

        if not all_final_cleaned_texts:
            self.log_message("\nNo text content was extracted or all extracted content was empty after cleaning.")
            messagebox.showinfo("No Content", "No text content was extracted or all was empty after cleaning.")
            self.update_status("Finished. No text content found.")
            self.btn_process.config(state=tk.NORMAL)
            return

        final_merged_text = "\n\n".join(all_final_cleaned_texts)

        output_file = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=(("Text files", "*.txt"), ("All files", "*.*")),
            title="Save Merged Text As..."
        )

        if output_file:
            try:
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(final_merged_text)
                self.log_message(f"\nSuccessfully merged and saved to: {output_file}")
                messagebox.showinfo("Success", f"Text content merged and saved to:\n{output_file}")
                self.update_status(f"Completed! Saved to {os.path.basename(output_file)}")
            except Exception as e:
                self.log_message(f"\nError saving file: {e}")
                messagebox.showerror("Save Error", f"Could not save the file: {e}")
                self.update_status("Error saving file.")
        else:
            self.log_message("\nSave operation cancelled by user.")
            self.update_status("Save cancelled.")

        self.log_message("\n--- Processing Finished ---")
        self.btn_process.config(state=tk.NORMAL)

    def start_processing_thread(self):
        self.btn_process.config(state=tk.DISABLED)
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete('1.0', tk.END)
        if not PDFMINER_AVAILABLE:
            self.log_message("WARNING: pdfminer.six library not found. PDF extraction is disabled.")
        if not DOCX_AVAILABLE:
            self.log_message("WARNING: python-docx library not found. DOCX extraction is disabled.")
        self.log_area.config(state=tk.DISABLED)

        processing_thread = threading.Thread(target=self.process_selected_items, daemon=True) # Call renamed method
        processing_thread.start()


if __name__ == "__main__":
    if os.name == 'nt':
        if not rarfile.UNRAR_TOOL:
            common_paths = [
                os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "WinRAR", "UnRAR.exe"),
                os.path.join(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"), "WinRAR", "UnRAR.exe"),
            ]
            for path in common_paths:
                if os.path.exists(path):
                    rarfile.UNRAR_TOOL = path
                    print(f"INFO: Automatically set rarfile.UNRAR_TOOL to: {path}")
                    break
            if not rarfile.UNRAR_TOOL:
                 print("INFO: 'UnRAR.exe' not found in typical WinRAR paths. "
                      "If RAR support is needed, ensure 'unrar.exe' is in your system PATH or set rarfile.UNRAR_TOOL manually.")

    root = tk.Tk()
    app = TextExtractorMergerApp(root)
    root.mainloop()
