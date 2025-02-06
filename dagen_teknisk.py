import os
import pdfplumber
import pandas as pd
import argparse

def extract_lines_from_pdf(pdf_path, start_page, end_page):
    """
    Extracts words from a PDF, concatenates them into lines, and returns the line data.

    Args:
        pdf_path (str): The path to the PDF file to be processed.
        start_page (int): The first page to start extraction from.
        end_page (int): The last page to stop extraction at.

    Returns:
        dict: A dictionary with page numbers as keys and lists of lines as values.
    """
    line_data = {}
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        print(f"PDF has {total_pages} pages.")

        # Adjust end_page to be skip the last pages of terms if it's not explicitly set
        if end_page == -1:
            end_page = max(1, total_pages - 3)

        for page_number in range(start_page, end_page):
            if page_number >= total_pages:
                print(f"Page number {page_number} is outside the total number of pages.")
                continue
            
            page = pdf.pages[page_number]
            cropped_page = page.crop((0, 130, page.width, page.height))
            words = cropped_page.extract_words()
            lines = concatenate_words_to_lines(words)
            line_data[page_number] = lines
        print(line_data)

    return line_data

def concatenate_words_to_lines(words):
    """
    Concatenates words into lines based on their y-coordinate on the page.

    Args:
        words (list): A list of words extracted from a page.

    Returns:
        list: A list of lines.
    """
    lines = []
    current_line = []
    last_top = None

    for word in words:
        if last_top is None or abs(word['top'] - last_top) > 5:
            if current_line:
                line_text = ' '.join([w['text'] for w in current_line])
                lines.append(line_text)
            current_line = [word]
            last_top = word['top']
        else:
            current_line.append(word)
    
    if current_line:
        line_text = ' '.join([w['text'] for w in current_line])
        lines.append(line_text)

    return lines

def sanitize_filename(name):
    """
    Sanitizes a filename by replacing invalid characters with underscores.

    Args:
        name (str): The original filename.

    Returns:
        str: The sanitized filename.
    """
    return ''.join(c if c.isalnum() or c in (' ', '_') else '_' for c in name).strip()

def parse_pdf_structure(line_data):
    """
    Parses the structured data from extracted lines, identifying sections and tables.

    Args:
        line_data (dict): A dictionary containing page numbers and lines of text.

    Returns:
        list: A list of dictionaries representing sections with metadata and tables.
    """
    sections = []
    current_section = None
    in_table = False
    table_data = []
    
    def end_current_section():
        """
        Ends the current section, finalizing its data and appending it to the sections list.
        """
        nonlocal current_section, table_data, in_table
        if current_section and not current_section.get('finalized', False):
            if table_data:
                current_section['tables'].append(pd.DataFrame(table_data, columns=["Pos", "Antall", "Navn"]))
            sections.append(current_section)
            for s in sections:
                s['finalized'] = True
        current_section = None
        in_table = False
        table_data = []

    for page_number, lines in line_data.items():
        for line in lines:
            line = line.strip()
            if line.startswith("Jobb navn:"):
                end_current_section()
                section_name = line.replace("Jobb navn:", "").strip()
                current_section = {
                    "section_name": section_name,
                    "start_date": "",
                    "return_date": "",
                    "brukerdager": "",
                    "tables": [],
                    "totals": {}
                }
                in_table = False
                table_data = []
            elif current_section:
                if line.startswith("Start dato"):
                    start_date = line.replace("Start dato", "").strip()
                    current_section['start_date'] = start_date
                elif line.startswith("Retur dato"):
                    return_date = line.replace("Retur dato", "").strip()
                    current_section['return_date'] = return_date
                elif line.startswith("Brukerdager"):
                    brukerdager = line.replace("Brukerdager", "").strip()
                    current_section['brukerdager'] = brukerdager
                elif line.startswith("Total utstyr:"):
                    current_section['totals']['total_utstyr'] = line.replace("Total utstyr:", "").strip()
                elif line.startswith("Total eks.mva"):
                    current_section['totals']['total_eks_mva'] = line.replace("Total eks.mva", "").strip()
                elif line.startswith("Pos") and "Antall" in line and "Navn" in line:
                    if table_data:
                        current_section['tables'].append(pd.DataFrame(table_data, columns=["Pos", "Antall", "Navn"]))
                        table_data = []
                    in_table = True
                elif in_table:
                    split_line = line.split(maxsplit=2)
                    if len(split_line) == 3:
                        table_data.append(split_line)
                    else:
                        if table_data:
                            table_data[-1][-1] += f" {line.strip()}"

    end_current_section()
    return sections

def process_pdf(pdf_path, start_page, end_page):
    """
    Processes the PDF to extract and parse its data.

    Args:
        pdf_path (str): Path to the PDF file.
        start_page (int): The first page to start extraction from.
        end_page (int): The last page to stop extraction at.

    Returns:
        list: A list of parsed sections with their data.
    """
    line_data = extract_lines_from_pdf(pdf_path, start_page, end_page)
    sections = parse_pdf_structure(line_data)
    return sections

def main():
    """
    Main function to execute script functionality.
    Parses command-line arguments, processes the PDF, and saves results to CSVs.
    """
    parser = argparse.ArgumentParser(description="Process a PDF file and extract data.")
    parser.add_argument('pdf_path', type=str, help='Path to the PDF file to be processed.')
    parser.add_argument('--start_page', type=int, default=1, help='The first page to start extraction from (default: 1).')
    parser.add_argument('--end_page', type=int, default=-1, help='The last page to stop extraction at. Defaults to 3 less than the total page count.')
    args = parser.parse_args()

    sections = process_pdf(args.pdf_path, args.start_page, args.end_page)
    print(f"Extracted {len(sections)} sections.")
    
    for section in sections:
        if section['tables']:
            section_df = pd.concat(section['tables'], ignore_index=True)
            section_filename = sanitize_filename(f"{section['section_name']}_data") + ".csv"
            print(section_filename)
            file_exists = os.path.isfile(section_filename)
            
            mode = 'a' if file_exists else 'w'
            header = not file_exists

            section_df.to_csv(section_filename, mode=mode, header=header, index=False, encoding='utf-8-sig')
            print(f"Section data saved to {section_filename}")

    all_tables = [pd.concat(section['tables'], ignore_index=True) for section in sections if section['tables']]
    combined_df = pd.concat(all_tables, ignore_index=True)
    combined_filename = "combined_data.csv"
    combined_df.to_csv(combined_filename, index=False, encoding='utf-8-sig')
    print(f"All sections combined data saved to {combined_filename}")

if __name__ == "__main__":
    main()
