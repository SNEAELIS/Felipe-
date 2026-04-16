import pandas as pd
import re
import sys


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


def confere_Sei():
    caminho_base = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\propostas_SEi.xlsx"
    caminho_origem_1 = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - webscraping\Consulta_SEi\Consultas parciais\consulta_direcao_sei_1-2.xlsx"
    caminho_origem_2 = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - webscraping\Consulta_SEi\Consultas parciais\consulta_direcao_sei_2-2.xlsx"

    master_file = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - webscraping\Consulta_SEi\consulta_direcao_sei_final.xlsx"

    # Read files
    print("Reading files...")
    df_origem_1 = pd.read_excel(caminho_origem_1)
    df_origem_2 = pd.read_excel(caminho_origem_2)

    df_base = pd.read_excel(caminho_base)

    df_master = pd.read_excel(master_file)


    # Get proposals
    source_raw_1 = set(df_origem_1['processo'].astype(str).tolist())
    source_raw_2 = set(df_origem_2['processo'].astype(str).tolist())
    base_raw = df_base['processo'].astype(str).tolist()
    master_raw = df_base['processo'].astype(str).tolist()


    output_path = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - webscraping\Consulta_SEi\Consultas parciais\nao_feitos.xlsx"
    source_raw = source_raw_1 | source_raw_2

    print(len(set(master_raw)) - len(set(base_raw)))

    print(len(set(master_raw)) - len(source_raw))

    not_iter = set(base_raw).difference(source_raw)
    df_not_iter = pd.DataFrame(columns=df_origem_1.columns)
    df_not_iter[df_not_iter.columns[0]] = list(not_iter)

    df_not_iter.to_excel(output_path, index=False)


def cruzar_instrumentos_func(arquivo_aspar, arquivo_busca, coluna_chave_aspar, coluna_chave_busca, coluna_retorno):
    """
    Cruzar dados entre dois arquivos Excel usando número de instrumento como chave.

    Parâmetros:
    - arquivo_aspar: caminho do arquivo Aspar (fonte)
    - arquivo_busca: caminho do arquivo de busca (resultado_aba_dados)
    - coluna_chave_aspar: nome da coluna com o número do instrumento no Aspar
    - coluna_chave_busca: nome da coluna com o número do instrumento no arquivo de busca
    - coluna_retorno: nome da coluna que contém o application number

    Retorna:
    - DataFrame do Aspar com a coluna de application number adicionada
    """

    # Carregar os arquivos
    df_aspar = pd.read_excel(arquivo_aspar, sheet_name='Aspar')
    df_busca = pd.read_excel(arquivo_busca)
    print(f'{len(df_aspar)}\n{len(df_busca)}')
    # Converter a coluna chave para string (para evitar problemas de tipo)
    # Isso padroniza os dados, tratando int e texto como strings
    df_aspar[coluna_chave_aspar] = df_aspar[coluna_chave_aspar].astype(str)
    df_busca[coluna_chave_busca] = df_busca[coluna_chave_busca].astype(str)

    # Remover espaços extras e caracteres invisíveis
    df_aspar[coluna_chave_aspar] = df_aspar[coluna_chave_aspar].str.strip()
    df_busca[coluna_chave_busca] = df_busca[coluna_chave_busca].str.strip()

    # Criar um dicionário para busca mais eficiente
    dicionario_busca = dict(zip(df_busca[coluna_chave_busca], df_busca[coluna_retorno]))

    # Adicionar a coluna de application number ao DataFrame do Aspar
    df_aspar['application_number'] = df_aspar[coluna_chave_aspar].map(dicionario_busca)

    # Substituir NaN por "Não Encontrado"
    df_aspar['application_number'] = df_aspar['application_number'].fillna("Não Encontrado")

    return df_aspar


def cruzar_instrumentos_main():
    df_resultado = cruzar_instrumentos_func(
        arquivo_aspar=r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Obras Maranhão - ASPAR-SNEAELIS.xlsx",
        arquivo_busca=r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\resultado_aba_dados_1-2.xlsx",
        coluna_chave_aspar='Instrumento',  # nome da coluna no Aspar
        coluna_chave_busca='Código do Instrumento',  # nome da coluna no arquivo de busca
        coluna_retorno='Número da Proposta'  # nome da coluna que você quer trazer
    )

    # Visualizar o resultado
    print(df_resultado.head())

    # Salvar o resultado em um novo arquivo Excel
    df_resultado.to_excel(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Obras Maranhão - ASPAR-SNEAELIS_resultado_cruzamento.xlsx", index=False)


def convert_pdf_to_xlsx():
    import pandas as pd
    import camelot

    pdf_file = r"C:\Users\felipe.rsouza\Downloads\Frequências e Ponto Normal 2022 Bruna_compressed PARTE 02 (2).pdf"
    xlsx_file = r"C:\Users\felipe.rsouza\Downloads\Frequências e Ponto.xlsx"
    tables = camelot.read_pdf(pdf_file, pages='all', flavor='lattice')

    combined_df = pd.DataFrame()
    for table in tables:
        df = table.df
        combined_df = pd.concat([combined_df, df], ignore_index=True)
    combined_df = combined_df.dropna(how='all').drop_duplicates()
    combined_df.to_excel(xlsx_file, sheet_name='Sheet1', index=False)

    print("PDF converted to Excel successfully!")


def merge_transf_spec():
    base_path = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Transferencias_Especiais\export_PlanoAcao_1775069545490 - Copia - Copia.xlsx"
    extention_path = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Transferencias_Especiais\export_PlanoAcao_1776249380121.xlsx"

    df_base = pd.read_excel(base_path, dtype=str)
    df_extent = pd.read_excel(extention_path, dtype=str)

    base = df_base.iloc[:, 0].dropna().to_list()
    extent = df_extent.iloc[:, 1].dropna()

    base_set = set(base)
    extent_set = set(extent)

    filtered = extent_set - base_set

    for i in filtered:
        print(i)
    print(len(filtered))


def test_table():
    pd.read_excel()


if __name__ == "__main__":
    #clean_process_data()
    #confere_Sei()
    #cruzar_instrumentos_main()
    #convert_pdf_to_xlsx()
    merge_transf_spec()
