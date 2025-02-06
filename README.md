# Pro Systemer Order Confirmation PDF Table Extractor

This script is specifically designed to extract structured data from order confirmation PDF files issued by the company ProSystemer. Using the `pdfplumber` library, it processes specific pages within these PDFs, extracts text lines, identifies data sections and tables, and saves the extracted data into CSV files for further analysis.

## Features

- Processes order confirmation PDFs from ProSystemer.
- Extracts and organizes text content into readable lines.
- Identifies structured sections and tables within the document.
- Outputs each section's table data into individual CSV files.
- Combines all extracted sections into a single CSV for comprehensive analysis.

## Requirements

- Python 3.x
- Python packages listed in `requirements.txt`

Your `requirements.txt` should include:
```bash
pdfplumber
pandas
```

To install the required packages, run:

```bash
pip install -r requirements.txt
```

## Usage

Execute the script from the command line, specifying the PDF file path, the start page, and the end page:

```bash
python pdf_table_extractor.py /path/to/prosystemer_order_confirmation.pdf --start_page 1 --end_page 10
```

### Command Line Arguments

- `pdf_path`: The path to the ProSystemer order confirmation PDF file to process.
- `--start_page`: (Optional) The first page to start extraction from. Defaults to 1.
- `--end_page`: (Optional) The last page to stop extraction at. Defaults to three pages before the end of the PDF if not specified. Use `-1` to apply the default behavior.

## Script Explanation

- **extract_lines_from_pdf**: Opens the PDF, crops pages to exclude any unwanted content like headers or footers, extracts words, and organizes these into lines of text. Handles the range of pages dynamically based on user input or defaults.

- **concatenate_words_to_lines**: Combines extracted words into coherent lines based on their vertical position, helping maintain the document's original structure.

- **sanitize_filename**: Cleans filename strings by replacing invalid characters with underscores, ensuring safe and compatible file creation.

- **parse_pdf_structure**: Analyzes extracted lines to identify various document sections and tabular data. Structures the table data into `pandas` DataFrames for ease of manipulation and export.

- **process_pdf**: Integrates extraction and parsing functions to produce organized section data from the PDF.

- **main**: The main execution block that configures argument parsing, manages processing, and saves the resulting data into CSV files.

## Example Output

Running the script generates separate CSV files for each section's data table extracted from the PDF. Additionally, all section data is compiled into a comprehensive `combined_data.csv` file for overall analysis.
