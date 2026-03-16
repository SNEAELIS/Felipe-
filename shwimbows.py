import pandas as pd
import re


def normalize_proposal(prop):
    """Convert any proposal format to standard 6-digit format"""
    if pd.isna(prop):
        return None

    prop_str = str(prop).strip()

    if '/' in prop_str:
        parts = prop_str.split('/')
        if len(parts) == 2:
            first_digits = re.sub(r'\D', '', parts[0])
            second_digits = re.sub(r'\D', '', parts[1])

            if first_digits and second_digits:
                first_padded = f"{int(first_digits):06d}"
                return f"{first_padded}/{second_digits}"

    return prop_str


def clean_process_data():
    caminho_base = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Propostas_Extraidas_filtered.xlsx"
    caminho_origem = (
        r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - "
        r"webscraping\Resultado scraping Aba Dados\resultado_aba_dados.xlsx")

    # Read files
    print("Reading files...")
    df_origem = pd.read_excel(caminho_origem)
    df_base = pd.read_excel(caminho_base)

    # Get proposals
    source_raw = df_origem['Número da Proposta'].astype(str).tolist()
    base_raw = df_base['Nº Proposta'].astype(str).tolist()

    # Create normalized sets
    source_norm = {normalize_proposal(p) for p in source_raw if normalize_proposal(p)}
    base_norm = {normalize_proposal(p) for p in base_raw if normalize_proposal(p)}

    # Find proposals in base but not in source
    extra_in_base = base_norm - source_norm

    print(f"\n=== RESULTS ===")
    print(f"Proposals in base but NOT in source: {len(extra_in_base)}")

    if extra_in_base:
        # Create DataFrame with just the proposal numbers
        df_extra = pd.DataFrame(sorted(list(extra_in_base)), columns=['Nº Proposta'])

        print(f"\nFirst 5 rows:")
        print(df_extra.head())

        # Save to Excel
        output_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                       r'Social\Teste001\propostas_extras_na_base.xlsx')
        df_extra.to_excel(output_path, index=False)

        print(f"\n✓ File saved successfully: {output_path}")
        print(f"✓ Total proposals exported: {len(df_extra)}")

        return df_extra

    else:
        print("\nNo extra proposals found in base file.")
        return None


if __name__ == "__main__":
    df_extra = clean_process_data()