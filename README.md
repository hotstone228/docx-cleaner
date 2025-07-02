# DOCX Cleaner

A Python script to **remove specified words** and **clear all headers and footers** from one or more `.docx` files. Useful for document anonymization, legal cleanup, or batch content scrubbing.

---

## 🚀 Features

- 🔍 Remove custom words from paragraphs
- 🧹 Clear all headers and footers (including first/even pages)
- 📂 Batch-process all `.docx`/`.doc` files in the current directory
- 📝 Input list of words via a plain `words_to_remove.txt` file

---

## 📦 Requirements

- Python 3.7+
- [`python-docx`](https://python-docx.readthedocs.io/en/latest/)

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## 📁 Usage

1. **Prepare a list of words** to remove:

   ```text
   words_to_remove.txt
   --------------------
   confidential
   JohnDoe
   password
   ```

2. **Place `.docx` files** in the same directory as the script.

3. **Run the script**:

   ```bash
   python docx_cleaner.py
   ```

4. **Output**: New files prefixed with `new_` will be saved, with words removed and headers/footers cleared.

---

## 🧠 How It Works

- Removes words using regular expressions (whole-word matching).
- Collapses extra spaces and punctuation after removals.
- Strips all headers and footers by modifying the internal XML structure of each section.

---

## 📂 File Structure

```text
├── docx_cleaner.py          # Main script
├── words_to_remove.txt      # List of words to redact
├── example.docx             # Example input file (not included)
├── new_example.docx         # Output (after cleaning)
└── requirements.txt         # Dependency list
```

---

## ⚠️ Notes

- This does **not** remove text from headers/footers based on content — it removes them completely.
- Be sure to review cleaned documents before distributing them.

---

## 📄 License

MIT License — free to use, adapt, and distribute.

---

## 🙌 Contributions Welcome

Feel free to open an issue or submit a PR for improvements like:

- CLI argument parsing (e.g., for custom folders)
- Regex support for advanced replacements
- GUI wrapper (Tkinter, etc.)
