# Docx_Translator

A Windows-first desktop tool for translating Microsoft Word (`.docx`) documents while preserving structure, formatting, tables, headers/footers, and images.

---

## Overview

**Docx_Translator** reads a `.docx` file, translates content (paragraphs, tables, headers/footers), preserves formatting, and saves a new `.docx`.

This project is especially useful if you:

- Want to translate Word documents programmatically
- Are learning Python through a real-world project
- Want to understand how `.docx` files are structured internally
- Need a starting point for more advanced document-processing tools

---

## Features

- GUI (Tkinter) with batch processing
- Per-file and batch progress with ETA
- Preserves runs/formatting, tables, headers/footers, and images
- Language selection UI (English/Spanish)
- Settings dialog for output language and options
- Validation + quality checks
- Optional DeepL provider (via API key)
- Translation memory (DeepL only)

### Current limitations

- Windows packaging only (for now)
- DeepL requires a valid API key

---

## How It Works (Conceptually)

1. Load `.docx`
2. Translate content in-place (runs, tables, headers/footers)
3. Preserve formatting and images
4. Save translated `.docx`

This project intentionally exposes these steps to make the learning process explicit.

---

## Requirements

- Python **3.10+** (recommended)
- A virtual environment (recommended)
- Dependencies listed in `requirements.txt`

---

## Installation (Dev)

### 1) Clone the repository

```bash
git clone https://github.com/JuanDarquea/Docx_Translator.git
cd Docx_Translator
```

### 2) Create and activate a virtual environment

```bash
python -m venv .venv
source .venv/bin/activate  # Linux / macOS
# .venv\\Scripts\\activate  # Windows
```

### 3) Install dependencies

```bash
pip install -r requirements.txt
```

---

## Usage (Dev)

Basic example:

```bash
python Docx_Translator.py
```

---

## Project Structure

```text
Docx_Translator/
│
├── Docx_Translator.py          # Main script
├── Files/                      # Default quality rules + protected words
├── build_windows.bat           # One-click Windows build
├── installer.iss               # Inno Setup script
├── Output/RELEASE_SIGNING.md   # Manual signing + release steps
├── requirements.txt
└── Readme.md
```

---

## Windows Build (One Click)

1. Ensure `.venv` exists and dependencies are installed:
   ```bash
   .venv\Scripts\activate
   pip install -r requirements.txt
   pip install pyinstaller
   ```
2. Run:
   ```bash
   build_windows.bat
   ```
3. Output:
   ```text
   dist\Docx_Translator.exe
   ```

## Download for Use

Once a release is published, download the Windows installer (`Docx_Translator_Setup.exe`) from the project’s Releases page.
Until then, use the Windows build steps above to produce your own `Docx_Translator.exe`.

## DeepL Setup

The app creates a user-local `Project_env.env` on first run.
Add your DeepL API key there:
```
DEEPL_API_KEY=your_key_here
```

Switch provider in `Docx_Translator.py`:
```
TRANSLATION_PROVIDER = "deepl"
```

## User Data Location

On Windows, user-local files are created here:
```
%APPDATA%\\Docx Translator\\
```

This includes:
- `settings.json`
- `Project_env.env`
- `translation_memory.json`
- `Files\\quality_rules.json`
- `Files\\protected_words.json`

## Release Specs (Windows)

- Installer: `Docx_Translator_Setup.exe` (built with Inno Setup)
- Desktop shortcut + Start Menu entry
- User data stored in `%APPDATA%\\Docx Translator\\`

## Notes

This project is evolving from a learning lab into a production-ready desktop tool.

---

## Contributing

Contributions are welcome — especially improvements that:

- Preserve formatting more accurately
- Add support for tables or headers
- Improve translation batching and performance

### Contribution steps

1. Fork the repository
2. Create a new branch
   ```bash
   git checkout -b feature/my-feature
   ```
3. Commit your changes
4. Push to your fork
5. Open a Pull Request

---

## Roadmap

- [ ] Signed executable
- [ ] Auto-updates

---

## 📄 License

This project is open source and available under the **MIT License**.

---

## 👤 Author

**Juan Darquea**
Python learner · Linux user · Document automation enthusiast

Feel free to open an issue, suggest improvements, or use this project as a learning reference.

---

## ⭐ Acknowledgments

- `python-docx` for Word document handling
- Translation APIs for making automation possible
- The open-source community for constant inspiration
