# 📄 Docx_Translator

A Python tool for translating Microsoft Word (`.docx`) documents from English language to Spanish language while preserving the document structure.

---

## 🚀 Overview

**Docx_Translator** is a learning-oriented Python project that reads a `.docx` file, processes its internal structure (paragraphs and runs), translates the text into a target language, and writes the translated content into a new `.docx` file.

This project is especially useful if you:

- Want to translate Word documents programmatically
- Are learning Python through a real-world project
- Want to understand how `.docx` files are structured internally
- Need a starting point for more advanced document-processing tools

---

## ✨ Features

- ✅ Reads `.docx` files using `python-docx`
- ✅ Translates paragraph text to a target language
- ✅ Generates a new translated `.docx` file
- ✅ Modular code structure for learning and extension
- 🧠 Educational focus: explores **paragraphs**, **runs**, and formatting

### Current limitations

- ❌ Does not yet preserve all text formatting (runs are being studied)
- ❌ No support for tables, images, headers, or footers (yet)
- ❌ API rate limits depend on the translation provider

---

## 🧠 How It Works (Conceptually)

1. Load a `.docx` file
2. Iterate through paragraphs
3. Inspect and analyze runs inside each paragraph
4. Send text to a translation service
5. Write translated text into a new document

This project intentionally exposes these steps to make the learning process explicit.

---

## 🛠️ Requirements

- Python **3.7+**
- A virtual environment (recommended)
- Dependencies listed in `requirements.txt`

---

## 📦 Installation

### 1️⃣ Clone the repository

```bash
git clone https://github.com/JuanDarquea/Docx_Translator.git
cd Docx_Translator
```

### 2️⃣ Create and activate a virtual environment

```bash
python -m venv .venv
source .venv/bin/activate  # Linux / macOS
# .venv\\Scripts\\activate  # Windows
```

### 3️⃣ Install dependencies

```bash
pip install -r requirements.txt
```

---

## ▶️ Usage

Basic example:

```bash
python Docx_Translator.py
# a pop-up window will appear for you to choose the file to translate, languages are set for now and the output path will be the same as this file.
```

---

## 📂 Project Structure

```text
Docx_Translator/
│
├── Docx_Translator.py      # Main script
├── requirements.txt       # Python dependencies
├── README.md              # Project documentation
├── .gitignore
└── examples/              # Sample documents (optional)
```

---

## 📚 Learning Goals of This Project

This repository is also a **learning lab**. The goals include:

- Understanding how Word documents store text
- Learning why text is split into multiple **runs**
- Practicing clean Python structure
- Working with virtual environments
- Using Git and GitHub effectively

If you are learning Python or document automation, this project is designed to grow with you.

---

## 🧑‍🤝‍🧑 Contributing

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

## 🧭 Roadmap

- [ ] Preserve runs and formatting
- [ ] Translate tables
- [ ] CLI improvements
- [ ] Config file support
- [ ] Multiple translation providers

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

