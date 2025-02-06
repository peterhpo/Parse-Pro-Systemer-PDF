import os
import pdfplumber
import pandas as pd

start_page = 1
end_page = 15

# Function to extract words and concatenate them into lines
def extract_lines_from_pdf(pdf_path):
    line_data = {}
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        print(f"PDF has {total_pages} pages.")
        
        for page_number in range(start_page, end_page + 1):
            if page_number >= total_pages:
                print(f"Page number {page_number} is outside the total number of pages.")
                continue
            
            page = pdf.pages[page_number]
            cropped_page = page.crop((0, 130, page.width, page.height))  # Crop the page
            words = cropped_page.extract_words()
            lines = concatenate_words_to_lines(words)
            line_data[page_number] = lines

    return line_data

# Function to concatenate words into lines based on their y-coordinate
def concatenate_words_to_lines(words):
    lines = []
    current_line = []
    last_top = None

    for word in words:
        if last_top is None or abs(word['top'] - last_top) > 5:  # Adjust threshold as needed
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

# Function to sanitize filenames
def sanitize_filename(name):
    return ''.join(c if c.isalnum() or c in (' ', '_') else '_' for c in name).strip()

# Function to group lines into sections and tables
def parse_pdf_structure(line_data):
    sections = []
    current_section = None
    in_table = False
    table_data = []
    
    def end_current_section():
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
                        # Handling case where additional multiline text might be encountered
                        if table_data:
                            table_data[-1][-1] += f" {line.strip()}"

    end_current_section()

    return sections

def process_pdf():
    pdf_path = '24-0254 Dagen@IFI Ordrebekreftelse v2.pdf'
    
    line_data = extract_lines_from_pdf(pdf_path)
    sections = parse_pdf_structure(line_data)

    return sections

def main():
    sections = process_pdf()
    print(f"Extracted {len(sections)} sections.")
    
    for section in sections:
        if section['tables']:
            section_df = pd.concat(section['tables'], ignore_index=True)
            section_filename = sanitize_filename(f"{section['section_name']}_data")+".csv"
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
