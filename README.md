# md2docx — Convert Markdown to Word (DOCX) from the Command Line

`md2docx.py` is a simple Python script that converts a Markdown file (`.md`) into a Word document (`.docx`) directly from the command line.  
It uses [pypandoc](https://pypi.org/project/pypandoc/) (with [Pandoc](https://pandoc.org/)) for high-quality conversion and falls back to running `pandoc` directly if `pypandoc` is not installed or fails.

---

## Features

- Convert Markdown files into Word DOCX with one command.
- Optional **reference DOCX** support to apply custom Word styles.
- Works even if `pypandoc` is missing (uses `pandoc` CLI as a fallback).
- Lightweight and easy to integrate into automation pipelines.

---

## Requirements

- Python 3.8 or later  
- [Pandoc](https://pandoc.org/) installed and available in your `PATH`  
  (recommended for full feature support).  
- Python packages:
  ```bash
  pip install pypandoc
  # Consigliato: installa anche pandoc (necessario per la conversione completa)
  # macOS: brew install pandoc
  # Ubuntu/Debian: sudo apt-get install pandoc
  # Windows (choco): choco install pandoc
  ```

---

## Installation

Download or clone this repository:

```bash
git clone https://github.com/yourusername/md2docx.git
cd md2docx
```

Make the script executable (optional):

```bash
chmod +x md2docx.py
```

---

## Usage

### Basic conversion

```bash
python md2docx.py document.md
```

Creates a `document.docx` file in the same folder.

### Specify output file

```bash
python md2docx.py document.md -o output/my_report.docx
```

### Use a custom Word style template

If you have a reference DOCX (Word file that defines heading styles, fonts, spacing, etc.):

```bash
python md2docx.py document.md --refdoc template.docx
```

---

## Example

```bash
python md2docx.py README.md -o README.docx
```

Output:

```
Created: README.docx
```

---

## Troubleshooting

* **`pypandoc` not installed**  
  Install it with:
  ```bash
  pip install pypandoc
  ```

* **`pandoc` not found**  
  Install Pandoc manually:
   * macOS: `brew install pandoc`
   * Ubuntu/Debian: `sudo apt-get install pandoc`
   * Windows: download from [pandoc.org/installing.html](https://pandoc.org/installing.html)

* **Conversion failed**  
  The script will print the reason (missing Pandoc, Python error, etc.). Make sure Pandoc is correctly installed and accessible from the terminal.

---

## License

GPL 3 License — you are free to use, modify, and share this script under the terms of the GNU General Public License v3.0.

---

## Credits

Developed by Emanuel Demetrescu. Relies on [pypandoc](https://pypi.org/project/pypandoc/) and [Pandoc](https://pandoc.org/) for Markdown-to-DOCX conversion.