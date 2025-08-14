# MSG Converter

A command-line tool to batch convert Outlook `.msg` files to `.pdf`, `.eml`, or `.mbox` formats. Built by the McClatchy Journalism AI team with the help of ChatGPT.

## Features

- Converts individual `.msg` files to:
  - 📄 PDF (with preserved formatting and embedded attachments)
  - ✉️ EML (raw email format with attachments)
  - 📦 MBOX (bulk archive via intermediate EML generation)

- Supports batch processing of folders
- Handles attachments and inline images
- Creates clean, readable PDFs with structured metadata

## Installation

### Requirements

- Python 3.8+
- Dependencies (install via pip):

```bash
pip install -r requirements.txt
```

*Note:* WeasyPrint requires additional system dependencies:

On *macOS*, using [Homebrew](https://brew.sh/):

```bash
brew install cairo pango gdk-pixbuf libffi
```

On *Ubuntu/Debian*:

```bash
sudo apt install libpango-1.0-0 libpangocairo-1.0-0 libcairo2 libffi-dev shared-mime-info
```

### Installing using a virtual environment

```bash
# Clone the repository to your local machine
git clone https://github.com/mcclatchyai/msg_converter.git
# Navigate to new directory
cd msg_converter
# Initialize a new virtual environment to install requirements and activate
python3 -m venv venv
source venv/bin/activate

# Within virtual environment, install dependencies
pip install --upgrade pip
pip install -r requirements.txt
pip install extract-msg weasyprint pytest
brew install cairo pango gdk-pixbuf libffi
```

## Usage

```bash
python main.py --input <input_folder> --output <output_folder> --format <pdf|eml|mbox>
```

### Examples
Convert all `.msg` files in a folder to PDF:

```bash
python main.py --input ./inbox --output ./pdf_out --format pdf
```

Convert to EML:

```bash
python main.py --input ./inbox --output ./eml_out --format eml
```

Convert to MBOX:

```bash
python main.py --input ./inbox --output ./mbox_out --format mbox
```

When converting to MBOX, the tool first converts each .msg file to .eml in a temporary folder, then compiles them into a single output.mbox file.

## Testing

```bash
pytest tests --msg-file "tests/sample/sample_message_a1.msg"    
```