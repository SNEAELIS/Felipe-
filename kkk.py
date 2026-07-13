import pandas as pd
import glob
import os, sys, re
from pathlib import Path
os.system('cls' if os.name == 'nt' else 'clear')

def diff_controle_prop_sourcer():
    acomp = pd.read_excel(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\dashboard-nodejs\DATA\Acompanhamento\Controle SNEAELIS - 2026.xlsx", dtype=str)
    tgov = pd.read_excel(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\webscraping\Resultado scraping Aba Dados\resultado_aba_dados.xlsx", dtype=str)
    acomp = acomp[~acomp['Nº Proposta'].str.contains('sem proposta', na=False)]
    tgov[['Número do Processo']] = tgov[['Número do Processo']].apply(lambda col: col.str.replace(r'^0+', '', regex=True))
    mask = tgov['Número do Processo'].apply(lambda x: bool(re.search(r'/(2026)$', str(x))) if pd.notna(x) else False)
    print(tgov[mask]['Número do Processo'].tolist())
    tgov_filtrado = acomp[~acomp['Nº Proposta'].isin(tgov['Número do Processo'])]
    print()

    tgov_filtrado = tgov_filtrado.reset_index(drop=True)
    #tgov_filtrado.to_excel(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\webscraping\Resultado scraping Aba Dados\diff\tgov_filtrado.xlsx", index=False)

    # Resultado

    print(f"{tgov_filtrado['Nº Proposta'].tolist()}\n {(len(tgov_filtrado['Nº Proposta'].tolist()))}")


def find_lost_props():
    df_diff = pd.read_excel(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\webscraping\Consulta_SEi\Consultas parciais\sei_diff.xlsx", dtype=str)

    tgov = pd.read_excel(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\webscraping\Resultado scraping Aba Dados\resultado_aba_dados.xlsx", dtype=str)
    
    
    df_diff = df_diff[~df_diff['texto_link'].str.contains('Processo não possui andamentos abertos.', na=False)].drop_duplicates()
    print(f"\Filtered df_diff: {len(df_diff)}\n {(df_diff.head())}")

    tgov = tgov[tgov['Número do Processo'].isin(df_diff['processo'])]
    print(f"\Filtered tgov: {len(tgov)}\n {(tgov.head())}")
    tgov.to_excel(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\webscraping\Resultado scraping Aba Dados\diff\tgov_filtrado.xlsx", index=False)

    # Resultado


def props_tgov_filtered():
    df_diff = pd.read_excel(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\webscraping\Resultado scraping Aba Dados\diff\tgov_filtrado.xlsx", dtype=str)

    tgov = pd.read_excel(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Propostas_Extraidas_filtradas.xlsx", dtype=str)

    res_aba_dados = pd.read_excel(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\webscraping\Resultado scraping Aba Dados\resultado_aba_dados.xlsx", dtype=str)

        
    tgov = tgov[~tgov['Nº Proposta'].isin(df_diff['Número da Proposta'])]
    print(f"\Filtered tgov: {len(tgov)}\n {(tgov.head())}")

    res_aba_dados = res_aba_dados[~res_aba_dados['Número da Proposta'].isin(tgov['Nº Proposta'])]

    res_aba_dados.to_excel(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\webscraping\Resultado scraping Aba Dados\diff\resultado_aba_dados.xlsx", index=False)
    #tgov[['Nº Proposta']].to_excel(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Propostas_Extraidas_filtradas.xlsx", index=False)



if __name__ == "__main__":
    select_ = input('Digite o número da função que deseja executar:\n1 - diff_controle_prop_sourcer\n2 - props_tgov_filtered\n3 - find_lost_props\n')
    if select_ == '1':
        diff_controle_prop_sourcer()
    elif select_ == '2':
        props_tgov_filtered()
    elif select_ == '3':
        find_lost_props()