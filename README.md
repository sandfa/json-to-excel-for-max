# JSON to Excel Converter

This project converts JSON files from a `data` directory into a single Excel file.

## Setup

1.  **Create a virtual environment:**
    ```bash
    python3 -m venv .venv
    source .venv/bin/activate
    ```

2.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1.  **Place JSON files:**
    Put your JSON files in the `data/` directory.

2.  **Run the conversion:**
    ```bash
    python3 main.py
    ```

3.  **Check Output:**
    The generated Excel file will be located at `output/result.xlsx`.
