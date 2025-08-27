import re
import shutil
import numpy as np
import io
import pdfplumber
import pandas as pd
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter
import os
import base64
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

def concatenate_excel_files(input_folder=None, output_file=None):
    path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
            r'Social\Teste001\tabela Código da Natureza de Despesa.xlsx')
    ref_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                r'Social\Teste001\atT00001 (3) (1).xlsx')

    df = pd.read_excel(path)
    print("Main DF columns:", df.columns.tolist())  # Debug print

    ref_df = pd.read_excel(ref_path, header=None)
    print("Reference DF columns:", ref_df.columns.tolist())  # Debug print

    ref_df = ref_df.applymap(lambda x: str(x).strip().lower() if pd.notna(x) else '')

    unique_values_col_b = ref_df[1].unique()
    print("Unique values in reference:", unique_values_col_b)  # Debug print

    try:
        # first occurrence index
        tipo_mask = [str(x).strip() == 'tipo' for x in unique_values_col_b]
        unique_idx = np.where(tipo_mask)[0]
        if len(unique_idx) == 0:
            raise ValueError("'TIPO' not found in reference file's column B")
        unique_idx = unique_idx[0]
        unique_values_temp = unique_values_col_b[unique_idx+1:]
    except Exception as e:
        print(f"Error finding 'TIPO' marker: {str(e)}")
        print("Values found:", unique_values_col_b)
        return

    unique_values = []
    for val in unique_values_temp:
       if isinstance(val, (str, int, float)):
            if not val or val.startswith('<built-in method'):
                continue
            if '/' in val:
                unique_values.extend([v.strip() for v in val.split('/') if v.strip()])
            else:
                unique_values.append(val.lower())
    print(f"Unique values {unique_values}")


    desc_col = next((col for col in df.columns if 'descrição' in str(col).lower()), None)

    if not desc_col:
        raise ValueError("Column 'Descrição' not found in the main DataFrame")

    for value in unique_values:
        # Print separator and DataFrame
        print('\n' + '=' * 80 + '\n')  # Black line separator (using = for visibility)
        print(f"Results for value: '{value}'\n")

        safe_value = re.escape(value)
        filtered_df = df[df['Descrição'].str.contains(safe_value, case=False, regex=True, na=False)]

        # Print DataFrame with formatting
        pd.set_option('display.max_rows', None)
        with pd.option_context('display.max_columns', None):
            print(filtered_df.to_string(index=False))
        print('\n' + '=' * 80 + '\n')  # Black line separator


def separate_nuclei():
    def aid_func(filename):
        pattern = r'Chamadas\s(.*?)(?:\.pdf|_compressed\.pdf|\.xlsx\.pdf|_compressed_compressed\.pdf)'
        match = re.search(pattern, filename, re.IGNORECASE)
        if match:
            nucleus_name = match.group(1).strip()
            return nucleus_name.replace('_.', '')
        return None

    dir_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                r'Social\SNEAELIS - Termo Fomento Inst. Léo Moura\Frequencia 926508')

    for root, dirs, files in os.walk(dir_path):
        if 'Frequencias 910782' in dirs:
            dirs.remove('Frequencias 910782')
        for filename in files:
            if not filename.lower().endswith('.pdf'):
                continue
            nucleus = aid_func(filename)

            if nucleus:
                new_dir_path = os.path.join(root, nucleus)
                if not os.path.exists(new_dir_path):
                    print(f"Creating directory: {new_dir_path}")
                    os.makedirs(new_dir_path)
                source_file = os.path.join(root, filename)
                destination_file = os.path.join(new_dir_path, filename)

                print(f"Moving file: {filename} to {new_dir_path}")
                shutil.move(source_file, destination_file)


def pdf_to_xlsx():
    def write_log(log_file, message):
        """Write messages to log file"""
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(message + '\n')

    def extract_tables_with_pdfplumber(pdf_source):
        """Extract tables using pdfplumber with focus on area below 'Nº'"""
        tables = []

        with pdfplumber.open(pdf_source) as pdf:
            for page in pdf.pages:
                # Find the 'Nº' marker position
                y_start = None
                for word in page.extract_words():
                    if word['text'].strip().upper() == 'Nº':
                        y_start = word['top']
                        break

                if y_start is None:
                    continue  # Skip if no Nº found

                # Define crop area (everything below Nº)
                crop_area = (0, y_start - 5, page.width, page.height)  # x0, top, x1, bottom

                # Extract table from this region
                cropped_page = page.crop(crop_area)
                table = cropped_page.extract_table()

                if table and len(table) > 1:  # Ensure we have at least header + one row
                    tables.append(table)

        return tables

    dir_path = (r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência "
                r"Social\SNEAELIS - Termo Fomento Inst. Léo Moura")

    # Create log file
    log_file = os.path.join(dir_path, f"conversion_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")

    for root, dirs, files in os.walk(dir_path):
        if 'Frequencias 910782' in dirs:
            dirs.remove('Frequencias 910782')
        for filename in files:
            if not filename.lower().endswith('.pdf'):
                continue

            pdf_path = os.path.join(root, filename)
            xlsx_filename = filename.replace('.pdf', '.xlsx')
            xlsx_path = os.path.join(root, xlsx_filename)

            log_message = f"\nProcessing {filename}..."
            print(log_message)
            write_log(log_file, log_message)

            try:
                # Try with original PDF first
                try:
                    tables = extract_tables_with_pdfplumber(pdf_path)
                    if tables:
                        log_message = "Extracted tables with pdfplumber"
                        print(log_message)
                        write_log(log_file, log_message)
                    else:
                        raise ValueError("No tables found in initial extraction")
                except Exception as e:
                    log_message = f"Initial extraction failed, trying cleaned version: {str(e)[:200]}"
                    print(log_message)
                    write_log(log_file, log_message)

                    # Clean PDF in memory
                    pdf_reader = PdfReader(pdf_path)
                    pdf_writer = PdfWriter()

                    for page in pdf_reader.pages:
                        if "/Annots" in page:
                            del page["/Annots"]
                        pdf_writer.add_page(page)

                    # Create in-memory PDF bytes
                    pdf_bytes = io.BytesIO()
                    pdf_writer.write(pdf_bytes)
                    pdf_bytes.seek(0)

                    # Try extraction again with cleaned PDF
                    tables = extract_tables_with_pdfplumber(pdf_bytes)
                    if not tables:
                        raise ValueError("All extraction methods failed")

                # Save to Excel
                with pd.ExcelWriter(xlsx_path) as writer:
                    for i, table in enumerate(tables):
                        try:
                            # Convert table data to DataFrame
                            df = pd.DataFrame(table[1:], columns=table[0])
                            df.to_excel(writer, sheet_name=f"Table_{i + 1}", index=False)
                        except Exception as e:
                            log_message = f"Error saving table {i + 1}: {str(e)[:200]}"
                            print(log_message)
                            write_log(log_file, log_message)
                            continue

                log_message = f"Successfully converted {filename} to {xlsx_filename}"
                print(log_message)
                write_log(log_file, log_message)

            except Exception as e:
                log_message = f"Failed to process {filename}: {type(e).__name__} - {str(e)[:200]}"
                print(log_message)
                write_log(log_file, log_message)
                continue


def html_to_pdf(html_file_path, pdf_file_path=None):
    """
    Convert an HTML file to PDF using Chrome browser.

    Args:
        html_file_path (str): Path to the input HTML file
        pdf_file_path (str): Path for the output PDF file (optional)

    Returns:
        bool: True if conversion was successful, False otherwise
    """
    # If no output path is provided, use the same name with .pdf extension
    if pdf_file_path is None:
        pdf_file_path = os.path.splitext(html_file_path)[0] + '.pdf'

    # Set up Chrome options
    chrome_options = Options()
    chrome_options.add_argument('--headless')  # Run in background
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--window-size=1920,1080')  # Set window size for consistent rendering

    try:
        # Initialize the driver
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)

        try:
            # Convert file path to URL (handle Windows paths)
            absolute_path = os.path.abspath(html_file_path)
            # Replace backslashes with forward slashes for Windows
            absolute_path = absolute_path.replace('\\', '/')
            file_url = f"file:///{absolute_path}"

            # Navigate to the HTML file
            driver.get(file_url)

            # Wait for page to load (more reliable than implicit wait for PDF generation)
            time.sleep(2)  # Adjust this based on your content complexity

            # Print to PDF with better options
            print_options = {
                'landscape': False,
                'displayHeaderFooter': False,
                'printBackground': True,
                'preferCSSPageSize': True,
                'marginTop': 0.5,
                'marginBottom': 0.5,
                'marginLeft': 0.5,
                'marginRight': 0.5,
                'pageRanges': '',  # Print all pages
            }

            # Execute the print command
            result = driver.execute_cdp_cmd('Page.printToPDF', print_options)

            # Ensure the output directory exists
            os.makedirs(os.path.dirname(pdf_file_path), exist_ok=True)

            # Save the PDF
            with open(pdf_file_path, 'wb') as f:
                f.write(base64.b64decode(result['data']))

            print(f"✓ Successfully converted: {html_file_path} -> {pdf_file_path}")
            return True

        except Exception as e:
            print(f"✗ Error processing {html_file_path}: {str(e)}")
            return False

        finally:
            driver.quit()

    except Exception as e:
        print(f"✗ Failed to initialize browser for {html_file_path}: {str(e)}")
        return False


def convert_all_html_files(root_directory='.', output_directory=None, pattern=('.html', '.htm')):
    """
    Recursively find and convert all HTML files in a directory tree.

    Args:
        root_directory (str): Root directory to search for HTML files
        output_directory (str): Custom output directory (optional)
        pattern (tuple): File extensions to search for
    """
    success_count = 0
    total_count = 0

    print(f"Searching for HTML files in: {os.path.abspath(root_directory)}")

    for root, dirs, files in os.walk(root_directory):
        for file in files:
            if file.lower().endswith(pattern):
                total_count += 1
                html_file_path = os.path.join(root, file)

                # Determine output PDF path
                if output_directory:
                    # Maintain directory structure in output directory
                    relative_path = os.path.relpath(root, root_directory)
                    pdf_dir = os.path.join(output_directory, relative_path)
                    pdf_file_path = os.path.join(pdf_dir, os.path.splitext(file)[0] + '.pdf')
                else:
                    # Save in same directory as HTML file
                    pdf_file_path = None

                # Convert the file
                if html_to_pdf(html_file_path, pdf_file_path):
                    success_count += 1

    print(f"\n📊 Conversion Summary:")
    print(f"   Total HTML files found: {total_count}")
    print(f"   Successfully converted: {success_count}")
    print(f"   Failed: {total_count - success_count}")

    return success_count, total_count


def split_parecer():
    try:
        df = pd.read_excel(r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                           r'Social\Teste001\Sofia\Pareceres_SEi\processos_DB.xlsx', dtype=str)
        for i, r in df.iterrows():
            partial_txt = r.iloc[1]
            txt = partial_txt.split('(')[-1].strip()
            txt = txt.replace(')', '')
            df.at[i,1] = txt
        df.to_excel(r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                           r'Social\Teste001\Sofia\Pareceres_SEi\processos_DB.xlsx',
                    index=False,)
    except Exception as e:
        print(f"✗ Failed to initialize browser for {type(e).__name__}: {str(e)[:100]}")
        return False



if __name__ == "__main__":
    func = int(input("Choose a function: "))

    if func == 1:
        concatenate_excel_files()
    elif func == 2:
        separate_nuclei()
    elif func == 3:
        pdf_to_xlsx()
    elif func == 4:
        root_dir = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                    r'Social\Teste001\Sofia\Pareceres_SEi')
        convert_all_html_files(root_directory=root_dir)
    elif func == 5:
        split_parecer()