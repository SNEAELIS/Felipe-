import pandas as pd
import os

def load_celeb_pt_data_custom(xlsx_path):
    sheets = [
        'Certidões', 'Declarações', 'Comprovantes de Execução',
        'Outros', 'Histórico Requisitos', 'Lista Anexos Proposta', 'Lista Anexos Execução'
    ]
    KEY_COL  = "N° da proposta"
    RANK_COL = "__row_rank__"   # temporary helper; removed at the end

    # ------------------------------------------------------------------
    # Strategy: cumcount + merge on (KEY_COL, RANK_COL)
    #
    # Each sheet may have N rows for the same proposal.
    # cumcount() assigns 0, 1, 2 … to each row within its proposal group.
    # Merging on (KEY_COL, RANK_COL) aligns rows side-by-side:
    #
    #   row 0 of Certidões  ↔  row 0 of Declarações  → same output row
    #   row 1 of Certidões  ↔  row 1 of Declarações  → same output row
    #   ...
    #
    # If one sheet has fewer rows for a proposal, the extra rows from the
    # other sheet simply get NaN in that sheet's columns — no data is lost.
    # ------------------------------------------------------------------

    print("--- Reading and aligning sheets ---")
    df_result = None

    for sheet in sheets:
        print(f"  Processing sheet: {sheet}")
        df_sheet = pd.read_excel(xlsx_path, sheet_name=sheet, dtype=str)

        # Skip sheets that don't have the key column at all
        if KEY_COL not in df_sheet.columns:
            print(f"    ⚠ Key column not found in '{sheet}', skipping.")
            continue

        # Drop rows where the proposal key is missing
        df_sheet = df_sheet.dropna(subset=[KEY_COL])

        # Prefix all data columns with the sheet name to avoid collisions
        df_sheet = df_sheet.rename(columns={
            col: f"{sheet}_{col}"
            for col in df_sheet.columns if col != KEY_COL
        })

        # Assign sequential rank within each proposal group (0, 1, 2, …)
        df_sheet[RANK_COL] = df_sheet.groupby(KEY_COL).cumcount()

        if df_result is None:
            df_result = df_sheet
        else:
            # outer join on (proposal key + row rank) → side-by-side alignment
            df_result = pd.merge(
                df_result, df_sheet,
                on=[KEY_COL, RANK_COL],
                how="outer"
            )

    # Remove the temporary rank column
    if df_result is not None and RANK_COL in df_result.columns:
        df_result = df_result.drop(columns=[RANK_COL])

    # Sort so rows for the same proposal are grouped together
    df_result = df_result.sort_values(by=KEY_COL).reset_index(drop=True)

    print(f"\nDone: {df_result.shape[0]} rows × {df_result.shape[1]} columns.")
    return df_result


# Clear screen and execute
os.system('cls' if os.name == 'nt' else 'clear')

df_final = load_celeb_pt_data_custom(
    r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\webscraping\demanda_andrei\Requisitos_Celebracao_e_Plano_de_Trabalho.xlsx"
)
df_final.to_excel(
    r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\webscraping\demanda_andrei\Requisitos_Celebracao_e_Plano_de_Trabalho_custom.xlsx",
    index=False
)
print("Saved to Excel.")