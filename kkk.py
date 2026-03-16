import os, glob, pandas as pd
from pathlib import Path

def merge_results_files(dir_path: str, output_file_name: str="merged_results.xlsx"):
    file_pattern = 'resultado_aba_esclarecimento_?-?.xlsx'

    search_pattern = os.path.join(dir_path, file_pattern)
    matching_files = glob.glob(search_pattern)

    if not matching_files:
        print(f"No files found matching pattern: {file_pattern}")
        return None

    matching_files.sort()
    dataframes = []

    print(f'found {len(matching_files)} files')
    for file_path in matching_files:
        try:
            df = pd.read_excel(file_path)
            dataframes.append(df)
            print(f"  ✓ Read: {os.path.basename(file_path)} - {len(df)} rows")

        except Exception as e:
            print(f"  ✗ Error reading {os.path.basename(file_path)}: {e}")
    
    if not dataframes:
        print("No data could be read from the files.")
        return None

    merged_dfs = pd.concat(dataframes, ignore_index=False)

    print(f"\nMerged {len(dataframes)} files into one dataframe with {len(merged_dfs)} total rows")

    output_file_path = os.path.join(dir_path, output_file_name)
    merged_dfs.to_excel(output_file_path, index=False)

    print(f"Merged file saved to: {output_file_path}")


def count_process_numbers(file_path: str, dir_path: str):
    df = pd.read_excel(file_path, dtype=str)
    df_filter = df['Número da Proposta'].to_list()
    base_dir = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Propostas_Extraidas_filtered.xlsx"
    base_df = pd.read_excel(base_dir)
    base_df_filter = base_df['Nº Proposta'].to_list()

    missing_props_path = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\missing_props.xlsx"
    my_set = set(base_df_filter) - set(df_filter)

    missing_df = pd.DataFrame(list(my_set), columns=["Nº Proposta"])
    missing_df.to_excel(missing_props_path)


if __name__ == '__main__':
    dir_path = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001"
    file_path = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\merged_results.xlsx"
    #merge_results_files(dir_path=dir_path)
    count_process_numbers(file_path=file_path, dir_path=dir_path)