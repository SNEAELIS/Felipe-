from datetime import datetime
import pandas as pd
import time
import re
import win32com.client as win32



def update_xlsx(target_file: str, source_file: str):
    # Read both spreadsheets into DataFrames
    df_source = pd.read_excel(source_file)
    df_target = pd.read_excel(target_file)

    # Ensure columns are aligned and both have the same columns for update
    column_to_update = df_source.columns

    updated_count = 0  # Counter for updated rows

    # Iterate over rows in source where column D (index 3) is not null/empty
    for _, src_row in df_source[df_source.iloc[:, 3].notnull()].iterrows():
        key_value = src_row.iloc[0]  # value in column A
        # Find matching row in target where column A matches
        match = df_target[df_target.iloc[:, 0] == key_value]
        if not match.empty:
            idx = match.index[0]
            # Update all columns in the target row with source row's data
            df_target.loc[idx, column_to_update] = src_row.values
            updated_count += 1  # Increment counter

    # Save the updated target file
    df_target.to_excel(target_file, index=False)
    print(f"Updated file saved as: {target_file}")
    print(f"Number of rows updated: {updated_count}")


def send_emails_from_excel(excel_path,):
    """
    Prepares emails from Excel data (Columns A, B, C) and opens them in Outlook for review.

    Args:
        excel_path (str): Path to Excel file.
        attachment_path (str, optional): Attachment file path.
        sender (str, optional): Email sender address.
    """

    def prepare_outlook_email(email, subject, html_body):
        """Creates and displays an Outlook email (without sending)."""
        try:
            if not all(isinstance(x, str) for x in [email, subject, html_body]):
                raise TypeError("Recipient, subject, and body must be strings")

            outlook = win32.Dispatch('outlook.application')
            time.sleep(2)
            e_mail = outlook.CreateItem(0)

            e_mail.To = email
            e_mail.Subject = subject
            e_mail.HTMLBody = html_body

            e_mail.Display()  # Opens email for review (instead of .Send())
            print(f"📧 Email prepared for: {email}")
            return True
        except Exception as e:
            print(f"❌ Failed to prepare email for {email}: \n{e}")
            return False

    def generate_email_body(extra_data_list):
        """Generates HTML body from Excel data only."""
        today = datetime.now().strftime("%d/%m/%Y")
        subject = f"Atualização das diligencias. Data {today}"
        html_body = """<!DOCTYPE html>
        <html>
        <head>
            <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        </head>
        <body style="font-family: 'Courier New', monospace; margin: 0; padding: 0;">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
                <td>"""
        for data in extra_data_list:
           ''' html_body += (
                f"<p style='white-space: pre ; font-family:Courier New, monospace>Código do Plano de "
                f"Ação: {data['A']}      "
                f"Responsável:"
                f" {data['B']}    "
                f"Data/Hora: {data['C']}    Situação: {data['D']}</p>"
                "<br><br></p>"  # 2 empty lines between entries
            )'''
           html_body += f"""
                       <div style="white-space: pre; margin-bottom: 20px;">
           Código do Plano de Ação:    {data['A']}
           Responsável:               {data['B']}
           Data/Hora:                 {data['C']}
           Situação:                  {data['D']}
                       </div>
                       """

           # Close all tags
           html_body += """</td>
               </tr>
           </table>
           </body>
           </html>"""
        return html_body, subject

    def read_excel_data():
        """Reads Excel and extracts Columns A, B, C if C has data."""
        try:
            df = pd.read_excel(excel_path, dtype=str)
            email = "sofia.souza@esporte.gov.br"
            extra_data_list = []  # Stores {A, B, C} dicts

            for _, row in df.iterrows():
                # Check if Column C (index 2) has data
                if pd.notna(row.iloc[3]):
                    extra_data = {
                        'A': row.iloc[0],  # Column A
                        'B': row.iloc[1],  # Column B
                        'C': row.iloc[2],  # Column C
                        'D': row.iloc[3]  # Column D
                    }
                    extra_data_list.append(extra_data)

            print(f'Número de planos de ação no email: {len(extra_data_list)}')

            return email, extra_data_list

        except Exception as e:
            print(f"❌ Failed to read Excel file: \n{e}")
            return [], []

    # Read data
    email, extra_data_list = read_excel_data()


    # Prepare emails
    html_body, subject = generate_email_body(extra_data_list)  # Pass single entry
    prepare_outlook_email(email, subject, html_body)


def align_xlsx(source_path, output_path):
    def delete_empty_rows(df):
        # Remove rows where the first column is NaN or a blank string
        cleaned_df = df[df.iloc[:, 0].notna() & (df.iloc[:, 0].astype(str).str.strip() != "")]
        return cleaned_df

    def extract_numbers_after_sei(df):
        def extract_sei_number(cell):
            if isinstance(cell, str):
                # Search for 'SEI' followed by any whitespace and then numbers (and optional punctuation)
                match = re.search(r'SEI\s*([0-9\.\-\/]+)', cell)
                if match:
                    return match.group(1)
            return cell

        df.iloc[:, 1] = df.iloc[:, 1].apply(extract_sei_number)
        return df

    def fill_column2_from_column1(df):
        mask = df.iloc[:, 2].isna() | (df.iloc[:, 2].astype(str).str.strip() == "")
        df.loc[mask, df.columns[2]] = df.loc[mask, df.columns[1]]
        return df

    def remove_parecer(cell):

        if isinstance(cell, str):
            return cell.replace('Parecer', '')
        return cell


    # Read file
    df = pd.read_excel(source_path)

    # Iterate over DataFrame
    for idx, row in df.iterrows():
        if pd.notna(row.iloc[0]):  # If column A has data
            block_start = idx
            # Scan following rows for empty col A (same block)
            next_idx = idx + 1
            while next_idx < len(df) and pd.isna(df.iloc[next_idx, 0]):
                # If there's data in B or C, move it up
                for col in [1, 2]:  # columns B and C (0-based index)
                    if pd.notna(df.iloc[next_idx, col]):
                        if pd.isna(df.iloc[block_start, col]):
                            df.iloc[block_start, col] = df.iloc[next_idx, col]
                        else:
                            # If multiple lines, concatenate (optional, can change)
                            df.iloc[block_start, col] = (
                                    str(df.iloc[block_start, col]) + "\n" + str(df.iloc[next_idx, col])
                            )
                        # Clear the moved cell
                        df.iloc[next_idx, col] = None
                next_idx += 1
    df = delete_empty_rows(df)
    df = extract_numbers_after_sei(df)
    df = fill_column2_from_column1(df)
    df = df.map(remove_parecer)
    df = df.drop('Processo SEI', axis=1)

    # Save result
    df.to_excel(output_path, index=False)
    print(f"Done! Saved as: {output_path[:55]}")


def align_meta_and_fields_with_next_custeio(output_filename=None):
    # Read the Excel file
    filename = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                r'Social\Teste001\Sofia\Emendas pix 2025\Emendas pix 2025_2_ciclo - Copia.xlsx')

    df = pd.read_excel(filename)

    # The fields to move together with "Meta"
    columns_to_move = [
        'Meta',
        'Descrição',
        'Unidade de Medida',
        'Quantidade',
        'Meses Previstos'
    ]

    # Check columns exist
    for col in columns_to_move + ['Categoria']:
        if col not in df.columns:
            raise ValueError(f"The DataFrame must contain the column: {col}")

    # Buffer for the values to move
    buffer = None
    last_custeio_idx = None

    for idx, row in df.iterrows():
        if row['Categoria'] == 'Custeio':
            last_custeio_idx = idx
        if row['Categoria'] == 'Investimento' and pd.notna(row['Meta']) and str(row['Meta']).strip():
            # Buffer the entire set of fields
            buffer = {col: row[col] for col in columns_to_move}
            # Clear these fields from the Investimento row
            for col in columns_to_move:
                df.at[idx, col] = ""
                df.at[last_custeio_idx, col] = buffer[col]
            buffer = None  # Clear the buffer

    # Output file name
    if output_filename is None:
        if filename.lower().endswith('.xlsx'):
            output_filename = f'{filename[:-13]}_{datetime.now():%d_%m_%Y}.xlsx'

    # Save the aligned DataFrame
    df.to_excel(output_filename, index=False)
    print(f"Aligned file saved as: {output_filename}")


def find_empty_cell_on_column(file_path, column: int=10):
    """
    Processes entire Excel file:
    1. For each row where column is empty:
        - Use 0 index numeration to find the desired column eg: k => 0
        - Copies column A value to the end of file
    2. Deletes all empty rows between index markers
    3. Preserves original formatting and sheets
    """
    # Read Excel with header=None to treat all as data
    df = pd.read_excel(file_path, header=None, engine='openpyxl')

    # Track rows to delete and values to append
    rows_to_delete = set()
    values_to_append = []

    # First pass: Identify processing targets
    i = int(0)
    print(f"Beginning loop. DF size = {len(df)}...".center(50))
    while i < len(df):
        if pd.isna(df.iloc[i, column]) and pd.notna(df.iloc[i, 0]):  # Column K (0-based index 10)
            # Record value from column A (index 0)
            values_to_append.append(df.iloc[i, 0])
            df.iloc[i, 0] = ""

            # Mark this row and subsequent empty A cells for deletion
            j = (i + 1)
            while j < len(df) and pd.isna(df.iloc[j, 0]):
                rows_to_delete.add(j)
                j += 1
            i = j  # Skip ahead
        else:
            i += 1

    print("Dropping rows...".center(50))

    # Perform deletions (in reverse order)
    df = df.drop(index=sorted(rows_to_delete, reverse=True)).reset_index(drop=True)

    print("Appending values...".center(50))

    # Append new rows
    if values_to_append:
        new_rows = pd.DataFrame({0: values_to_append})  # Column A
        df = pd.concat([df, new_rows], ignore_index=True)

    # Write back to Excel preserving original format
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, header=False)

    print(f"Processed {len(values_to_append)} appends and {len(rows_to_delete)} deletions.")


def drop_empty_rows_and_save(file_path):
    """
    Reads an Excel file, drops completely empty rows, and saves back to the same file.

    Args:
        file_path (str): Path to the Excel file (.xlsx or .xls)
    """
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Drop rows where ALL columns are empty
    df_cleaned = df.dropna(how='all')

    # Save back to the same file
    df_cleaned.to_excel(file_path, index=False)

    print(f"Removed {len(df) - len(df_cleaned)} empty rows. File saved to {file_path}")




if __name__ == "__main__":
    inp_ini = int(input('What to run?\n 1 to update "enviado para analises"; 2 to "align metas field; 3 '
                        'to find empty cell on column: 4 to drop_empty_rows_and_save"\n'))
    if inp_ini == 1:
        inp = input('Run all? Y/n\n')
        if inp == 'Y':
            update_xlsx(source_file=
            r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
            r'Social\Teste001\Sofia\PT_SNEAELIS_env_para_analise\enviados_para_analise - Copia.xlsx',
            target_file=
            r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
            r'Social\Teste001\Sofia\PT_SNEAELIS_env_para_analise\enviados_para_analise - Copia - Copia.xlsx'
            )

        r'''
        align_xlsx(source_path=
                   r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                   r'Social\Teste001\Sofia\Pareceres_SEi\processos_DB.xlsx',
                   output_path=
                   r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                   r'Social\Teste001\Sofia\Pareceres_SEi\processos_DB - Copia.xlsx')
        '''
        send_emails_from_excel(r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                r'Social\Teste001\Sofia\PT_SNEAELIS_env_para_analise\enviados_para_analise - Copia.xlsx')
    elif inp_ini == 2:
        align_meta_and_fields_with_next_custeio()
    elif inp_ini == 3:
        xlsx_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                     r'Social\Teste001\Sofia\Sofia Emendas PIX 2021 - Copia - Copia.xlsx')
        find_empty_cell_on_column(xlsx_path)
    elif inp_ini == 4:
        drop_empty_rows_and_save(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência "
                      r"Social\Teste001\Sofia\Emissão Parecer 2024 (732 PT) - Copia.xlsx")

