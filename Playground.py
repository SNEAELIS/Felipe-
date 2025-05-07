import os

from fuzzywuzzy import fuzz
import pandas as pd
import json
from itertools import islice
import random
import string

from numpy.matlib import empty


def pad_number(value):
    if isinstance(value, str):
        parts = value.split('/')
        if len(parts) > 1:
            number_part = parts[0]
            if len(number_part) == 4:
                padded_number = '00' + number_part
                return padded_number + '/' + '/'.join(parts[1:])
            elif len(number_part) == 5:
                padded_number = '0' + number_part
                return padded_number + '/' + '/'.join(parts[1:])
            else:
                return value
    else:
        return value

# Function to fix the domain part of the email
def fix_email(email):
    valid_domains = [
        'mds.gov.br',
        'esporte.gov.br'
    ]
    if not pd.isna(email) and '@' in email:
        # Split the email into local part and domain part
        local_part, domain = email.split('@', 1)
        # Add "@" back to the domain
        domain = '@' + domain

        # Check if the domain is valid
        if domain not in valid_domains:
            # Find the closest valid domain
            best_match = max(valid_domains, key=lambda d:fuzz.ratio(domain,d))
            return local_part + '@' + best_match

        # If there's no "@", assume the domain is missing
        else:
            return email + '@esporte.gov.br'
    return email

# Function to count occurrences of local parts
def count_local_partes(emails):
    local_parts_count = {}
    for email in emails:
        if not pd.isna(email) and '@' in email:
            local_part = email.split('@', 1)[0]
            local_parts_count[local_part] = local_parts_count.get(local_part, 0) +1
    return local_parts_count

# Function to correct typos in the local part
def fix_local_parts(email, local_part_count):
    if not pd.isna(email) and '@' in email:
        local_part, domain = email.split('@', 1)
        # If the local part appears only once, it might be a typo
        if local_part_count.get(local_part, 0) == 1:
            # Find the most similar local part in the dictionary
            best_match = max(local_part_count.keys(), key=lambda x:fuzz.ratio(local_part, x))
            # If the similarity is high enough, replace the local part
            if fuzz.ratio(local_part, best_match) > 80: # Adjust the threshold as needed
                return best_match + '@' + domain
    return email

# Planilhas Meritos e Custos DEAELIS/CGC
def exe():
    try:
        # Define the file path
        sheet_address = (r'C:\Users\felipe.rsouza\OneDrive - '
                         r'Ministério do Desenvolvimento e Assistência Social\Teste001')

        # Load the sheets
        sheet_a = pd.read_excel(sheet_address + r'\Controle Análise de Custos(Planilhas de Custo)'
                                                r'.xlsx', sheet_name=0, header=1)
        sheet_b = pd.read_excel(sheet_address + r'\Controle Análise de Custos(Planilhas de Custo)'
                                                r'.xlsx', sheet_name=1, header=1)

        # Define the column names
        num_propose_col = 'Num Proposta'
        email_col = 'Técnico Responsável '
        instrumento_col = 'Instrumento'

        # Clean and pad the 'Num Proposta' column in both sheets
        print('Cleaning columns')
        sheet_b[num_propose_col] = sheet_b[num_propose_col].str.strip()
        sheet_b[email_col] = sheet_b[email_col].str.strip()
        sheet_a[num_propose_col] = sheet_a[num_propose_col].str.strip()
        sheet_a[num_propose_col] = sheet_a[num_propose_col].apply(pad_number)
        sheet_b[num_propose_col] = sheet_b[num_propose_col].apply(pad_number)
        sheet_b[instrumento_col] = sheet_b[instrumento_col].str.strip()

        #Fix the domain part of the emails
        print('Fixing e-mails and e-mails domains')
        sheet_a['E-mail Técnico Custos'] = sheet_a['E-mail Técnico Custos'].apply(fix_email)
        local_part_count = count_local_partes(sheet_a['E-mail Técnico Custos'])
        sheet_a['E-mail Técnico Custos'] = sheet_a['E-mail Técnico Custos'].apply(lambda x:
                            fix_local_parts(x, local_part_count))

        # Add the '@esporte.gov.br' suffix to the email column in sheet_b
        sheet_b['Técnicos Ludmila'] = sheet_b[email_col].astype(str) + "@esporte.gov.br"

        # Create a mapping dictionary from 'Num Proposta' to 'Técnico Responsável'
        print('Mapping sheet b to sheet a')
        email_map = sheet_b.set_index(num_propose_col)['Técnicos Ludmila'].to_dict()

        # Add the 'Técnicos Ludmila' column to sheet_a by mapping 'Num Proposta'
        print('Adding column')
        sheet_a['Técnicos Ludmila'] = sheet_a[num_propose_col].map(email_map)

        # Identify rows in DF2 where 'Num Proposta' does not exist in DF1
        rows_to_append = sheet_b[~sheet_b[num_propose_col].isin(sheet_a[num_propose_col])]

        # Extract only the 'Num Proposta' and 'Técnicos Ludmila' columns from these rows
        rows_to_append = rows_to_append[['Instrumento', num_propose_col, email_col, 'Técnicos Ludmila']].rename(
            columns={'Instrumento': 'Tipo Instrumento:', email_col: 'Responsável  Análise Custos'}   )

        # Append these rows to the end of DF1
        print('Appending data')
        sheet_a = pd.concat([sheet_a, rows_to_append], ignore_index=True)
        sheet_a['E-mail Técnico Custos'] = sheet_a['E-mail Técnico Custos'].fillna('nomail')

        # Drop NaN values
        print('Droping NaN')
        sheet_a.fillna(0)

        robot_df = sheet_address + r'\Dataframe Custos e Méritos.xlsx'
        sheet_a.to_excel(excel_writer= robot_df, sheet_name='Custos_Méritos_DF', index=False)

        print("Process completed successfully!")
    except Exception as e:
        print(f"Process failed. Error: {e}")

arquivo_log = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento'
                   r' e Assistência Social\Automações SNEAELIS\Resultados Robô '
               r'Mérito-Custos( back_end )\arquivo_log.json')

def generate_password(length=12, use_letters=True, use_numbers=True, use_symbols=True):
    """
    Generate a random password.

    Parameters:
        length (int): Length of the password (default: 12).
        use_letters (bool): Include letters (default: True).
        use_numbers (bool): Include numbers (default: True).
        use_symbols (bool): Include symbols (default: True).

    Returns:
        str: Generated password.
    """
    # Define character sets
    letters = string.ascii_letters if use_letters else ""
    numbers = string.digits if use_numbers else ""
    symbols = string.punctuation if use_symbols else ""

    # Combine character sets based on user preferences
    all_characters = letters + numbers + symbols

    # Ensure at least one character type is selected
    if not all_characters:
        raise ValueError("At least one character type (letters, numbers, symbols) must be enabled.")

    # Generate the password
    password = "".join(random.choice(all_characters) for _ in range(length))
    return print("Generated Password:", password)

def filter_sheet(file_path):
    try:
        # Load the sheet
        df = pd.read_excel(file_path, sheet_name=0)
        # Step 1: Filter rows where 'MODO' is 'Convênio' or 'Termo de Fomento'
        filtered_df = df[df['MODO'].isin(['Convênio', 'Termo de Fomento'])]
        # Step 2: Remove rows with incompatible date structure
        # Assuming the date column is named 'DateColumn' (replace with the actual column name)
        # Check if the value is exactly 6 numeric digits
        filtered_df = filtered_df[filtered_df['Nº CONVÊNIO'].astype(str).str.match(r'^\d{6}$')]

        # Save the filtered DataFrame to a new Excel file
        output_file_path = file_path.replace('.xlsx', '_filtered.xlsx')
        filtered_df.to_excel(output_file_path, index=False)

        print(f"Filtered sheet saved to: {output_file_path}")
        return filtered_df

    except Exception as e:
        print(f"An error occurred: {e}")


def test_for():
    test_list = [
        3, 22, 29, 36, 42, 65, 71, 79, 85, 89, 107, 113, 114, 121, 128, 129, 132, 137, 139, 140, 142, 153, 159,
        161, 164, 167, 172, 177,183, 190, 193, 199, 224, 248, 259
             ]
    update_test_list = [i-2 for i in test_list]
    file_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social'
                 r'\Teste001\CGAC_2024_dataSource_filtered.xlsx')
    df = pd.read_excel(file_path)
    num_prop = df.loc[update_test_list, 'Nº CONVÊNIO']
    print(num_prop, '\n', len(test_list))


def categoryze_folder(directory_path: str):
    """
     Categorize folders by item count and return separate lists for:
     - Folders with exactly 20 items
     - Empty folders

     Args:
         directory_path (str): Path to the directory containing folders to check

     Returns:
         tuple: (folders_with_20_items, empty_folders)
     """
    folders_with_20 = []
    empty_folders = []
    total_processed = 0

    # Loop through each item in the directory
    for folder_name in os.listdir(directory_path):
        folder_path = os.path.join(directory_path, folder_name)
        # Check if it's a directory
        if os.path.isdir(folder_path):
            # Get all items excluding '.' and '..'
            items = [
                name for name in os.listdir(folder_path)
                if os.path.isfile(os.path.join(folder_path, name)) or
                   os.path.isdir(os.path.join(folder_path, name))
                     ]
            item_count = len(items)

            if item_count == 20:
                folders_with_20.append(folder_name)
            elif item_count == 0:
                empty_folders.append(folder_name)

            total_processed += 1

    return folders_with_20, empty_folders


def main_folders():
    target_directory = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento'
                        r' e Assistência Social\Automações SNEAELIS\Sel PAC DIE')
    xlsx_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento'
                 r' e Assistência Social\Teste001\tabela_Respostas_Ordenadas.xlsx')
    df = pd.read_excel(xlsx_path, header=0)

    twenty_item_folders, empty_folders = categoryze_folder(target_directory)
    print(f"\nEmpty folders ({len(empty_folders)}):")

    df_nots = df[df['Status'] == 'Não']
    df_not = df_nots['Nº Reservado PAC'].tolist()
    normal_empty_folders = [i.replace('_', '/') for i in empty_folders]
    diff = list(set(normal_empty_folders) - set(df_not))
    print(diff, '\n', len(empty_folders), '\n', len(df_not))


def has_png():
    root_dir = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento'
                r' e Assistência Social\Automações SNEAELIS\Sel PAC DIE')

    folder_counter = 0
    total_folders = 0

    for root, dirs, files in os.walk(root_dir):
        png = any(file.lower().endswith('.png') for file in files)
        total_folders += len(dirs)
        if png:
            folder_counter += 1
    print(f'Number of folders that has PNG file is: {folder_counter}.\nTotal number of folders:'
          f' {total_folders}.\nDifference: {total_folders - folder_counter}.\n'
          f'Ratio: {folder_counter/total_folders:2f}')



def use_filter_sheet():
    file_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social'
                 r'\Teste001\CGAC_2024_dataSource_filtered.xlsx')
    filter_sheet(file_path)


if __name__ == "__main__":
    #use_filter_sheet()
    #exe()
    #test_for()
    #password = generate_password(length=8, use_letters=True, use_numbers=True, use_symbols=True)
    #main_folders()
    #has_png()
    print(list(i for i in range(19, 27)))

