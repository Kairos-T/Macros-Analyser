# Macros-Analyser

This is a Python-based utility tool to analyse VBA macros within Microsoft Office documents. It is designed to be used
to automate the process of extracting and analysing VBA macros from Microsoft Office documents (Word, Excel, and
PowerPoint). It uses the `oletools` library to detect, extract, deobfuscate, and analyse VBA source code.

## Features

- Extract VBA macros from MS Office documents.
- Analyse VBA source code for malicious content.
- Deobfuscate VBA source code.

## Installation and Usage

1. Clone the repository:

    ```bash
    git clone https://github.com/Kairos-T/Macros-Analyser.git
    ```

2. Install the required packages:

    ```bash
    pip install -r requirements.txt
    ```

3. Place the document(s) to be analysed in the `document` directory.
4. Run the script:

    ```bash
    python main.py
    ```