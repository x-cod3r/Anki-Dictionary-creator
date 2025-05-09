# Anki Dictionary Creator

Create Anki flashcard decks from text files, with automatic word counting, translation, and text-to-speech audio generation.

![screeshot [(https://imgur.com/RwBlM8S)]]

## Features

*   **File Processing:** Select one or more `.txt` files to extract words.
*   **Word Counting:** Counts the frequency of each word.
*   **Customizable Word Limit:** Specify the maximum number of most frequent words to process.
*   **Language Support:**
    *   Specify input language for accurate word extraction (especially for CJK languages).
    *   Supported input languages include: English, Arabic, German, Spanish, French, Italian, Portuguese, Turkish, Dutch, Hebrew, Japanese, Korean, Russian, Chinese (Simplified), Swedish, Polish, Finnish, Greek, Hindi, Indonesian.
*   **Translation (Optional):**
    *   Translate extracted words to a target language using Google Translate.
    *   Supported target languages: English, Arabic, German, Spanish, French, Italian, Portuguese.
    *   Option to process words without translation.
*   **Text-to-Speech (TTS):**
    *   Generate audio pronunciation for words using Google Text-to-Speech.
    *   Speak individual words from the list.
    *   Speak all visible words in sequence.
*   **Anki Deck Export (.apkg):**
    *   Flexible export options for card fronts and backs:
        *   Word (Front) / Translation (Back)
        *   Translation (Front) / Word (Back)
        *   Word (Front) / Speech (Back)
        *   Translation (Front) / Speech + Word (Back)
        *   Word (Front) / Speech + Translation (Back)
    *   Audio is embedded in the Anki package.
    *   Customizable deck name.
    *   Option to select a temporary folder for audio file generation during export.
*   **User Interface:**
    *   Easy-to-use GUI built with Tkinter.
    *   Progress bar for file processing and Anki export.
    *   Results displayed in a sortable table (Word, Count, Translation).
    *   Copy selected words to clipboard.
*   **Responsive UI:** Uses `asyncio` to prevent UI freezes during long operations like translation and TTS generation.

## Requirements

*   Python 3.8+
*   `tkinter` (usually included with Python standard library)
*   `pygame` (for audio playback)
*   `googletrans==4.0.0-rc1` (for translation) - **Important: Use this specific version or a compatible one.**
*   `gTTS` (for text-to-speech)
*   `genanki` (for creating Anki .apkg files)
*   `pyperclip` (for clipboard operations)

## Installation

1.  **Clone the repository (or download the script):**
    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```

2.  **Install dependencies:**
    It's highly recommended to use a virtual environment.
    ```bash
    python -m venv venv
    # On Windows
    venv\Scripts\activate
    # On macOS/Linux
    source venv/bin/activate
    ```
    Then install the required packages:
    ```bash
    pip install pygame googletrans==4.0.0-rc1 gTTS genanki pyperclip
    ```
    *Note: If you encounter issues with `tkinter`, ensure it's installed with your Python distribution (it usually is, but on some Linux systems, it might be separate, e.g., `sudo apt-get install python3-tk`).*

## Usage

1.  Run the script:
    ```bash
    python anki_dictionary_creator.py 
    # (Replace anki_dictionary_creator.py with the actual script name if different)
    ```
2.  **Select File(s):** Click "Select File(s)" to choose one or more `.txt` files containing the text you want to process.
3.  **Set Input Language:** Choose the language of the text in your selected files from the "Input Lang" dropdown.
4.  **Set Translate To:** Choose the target language for translation. Select "None" if you don't want translation.
5.  **Word Limit:** Enter the maximum number of most frequent words you want to display and process.
6.  **Deck Name (for export):** Enter the desired name for your Anki deck.
7.  **Process Files:** Click "Process Files". The application will extract words, count them, translate (if a target language is selected), and display them in the table.
8.  **Interact with Results:**
    *   Click the "🔊" icon next to a word to hear its pronunciation (uses the "Input Lang" setting for TTS).
    *   Click "Speak All Visible" to hear all words in the current list.
    *   Select rows and press `Ctrl+C` (or `Cmd+C` on macOS) to copy words to the clipboard.
9.  **Export Anki Deck:**
    *   Choose an "Export As" format for your Anki cards.
    *   Click "Export Anki Deck".
    *   If your export format includes speech, you will be prompted to select a folder to temporarily store the generated audio files. These files will be packaged into the `.apkg` file.
    *   Save the `.apkg` file.
10. **Import into Anki:** Import the generated `.apkg` file into your Anki application.

## Temporary Files

*   **Audio for "Speak Word" / "Speak All":** When you use the speak functions, temporary audio files are created in a `temp_audio_files` sub-directory where the script is run. The application currently **does not** automatically delete this folder on exit, but it will log a message reminding you about it. You can manually delete this folder.
*   **Audio for Anki Export:** You select a directory for these temporary files during the export process. These files are then packaged by `genanki`. It is generally safe to clean this user-selected directory after the `.apkg` file has been successfully created.

## Known Issues / Considerations

*   **Google Translate API Limits:** The `googletrans` library uses an unofficial Google Translate API endpoint. Heavy usage (many words, frequent requests) can lead to temporary IP blocks (HTTP 429 errors). The application has some retry logic, but if you encounter persistent translation failures, try again later or process smaller batches of words.
*   **`googletrans` Version:** The `4.0.0-rc1` version is specified due to its past stability with the API. Other versions might behave differently or require code adjustments.
*   **Audio Playback (Pygame):** `pygame.mixer` initialization can sometimes fail on certain systems. If audio playback doesn't work, check the console for warnings.
*   **Large File Processing:** While `asyncio` is used to keep the UI responsive, processing extremely large files or a very high word limit might still consume significant resources and time.
*   **Shutdown:** Graceful shutdown of asyncio tasks and Tkinter is complex. If the application hangs on exit or doesn't close the console window immediately, there might be pending async operations or loop state issues.

## Contributing

Contributions, bug reports, and feature requests are welcome! Please open an issue or submit a pull request.


