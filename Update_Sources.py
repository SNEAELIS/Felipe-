import re
import shutil

from pathlib import Path

from datetime import datetime


def copy_todays_files_cmof(source_dir:Path, dest_dir:Path):
    """
    Copy files with today's date from current month folder to destination.
    Maintains same folder structure, creates folders if they don't exist.
    """
    source_path = Path(source_dir)
    dest_path = Path(dest_dir)
    today = datetime.now()

    # Month mapping
    months = {
        1: 'JANEIRO', 2: 'FEVEREIRO', 3: 'MARCO', 4: 'ABRIL',
        5: 'MAIO', 6: 'JUNHO', 7: 'JULHO', 8: 'AGOSTO',
        9: 'SETEMBRO', 10: 'OUTUBRO', 11: 'NOVEMBRO', 12: 'DEZEMBRO'
    }

    current_month_num = today.month
    current_month_name = months[current_month_num]

    # Find the folder for current month (format: "4 - ABRIL")
    month_folder = None
    for item in source_path.iterdir():
        if item.is_dir():
            # Check if folder name matches current month
            pattern = rf'^0?{current_month_num}\s*-\s*{current_month_name}$'
            if re.match(pattern, item.name.upper()):
                month_folder = item
                break

    if not month_folder:
        print(f"❌ Folder not found: {current_month_num} - {current_month_name}")
        return

    print(f"✅ Found folder: {month_folder.name}")

    # Create destination folder (same name as source folder)
    dest_month_folder = dest_path / month_folder.name
    dest_month_folder.mkdir(parents=True, exist_ok=True)

    if dest_month_folder.exists():
        print(f"✅ Destination folder ready: {dest_month_folder}")

    # Look for files with today's date
    today_str = today.strftime('%d-%m-%Y')
    date_pattern = re.compile(r'(\d{2})-(\d{2})-(\d{4})')

    copied_count = 0

    # Recursively iterate through all files in the month folder
    for file_path in month_folder.rglob('*'):
        if not file_path.is_file():
            continue

        # Check if filename contains today's date
        match = date_pattern.search(file_path.name)
        if match:
            file_date = f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
            if file_date == today_str:
                # Maintain same folder structure in destination
                relative_path = file_path.relative_to(month_folder)
                dest_file_path = dest_month_folder / relative_path

                # Create subfolders if needed
                dest_file_path.parent.mkdir(parents=True, exist_ok=True)

                # Copy the file
                shutil.copy2(file_path, dest_file_path)
                print(f"  ✓ Copied: {file_path.name}")
                copied_count += 1

    # Summary
    print(f"\n{'=' * 50}")
    print(f"SUMMARY")
    print(f"{'=' * 50}")
    print(f"Date: {today_str}")
    print(f"Source: {month_folder}")
    print(f"Destination: {dest_month_folder}")
    print(f"Files copied: {copied_count}")


def copy_tgov_files(source_dir:Path, dest_dir:Path):
    dirs_mapping = {
        'Resultado scraping Aba Dados': 'Aba Dados',
        'Consulta_SEi': 'SEI',
        'Extração_Requisitos_Celebracao_e_Plano_de_Trabalho': 'Requisitos',
        'Extração_TA_Prorrogas_Instrumentos': 'TAs'
    }

    for source_subdir, dest_subdir in dirs_mapping.items():
        # Create full paths
        source_subdir_path = source_dir / source_subdir
        dest_subdir_path = dest_dir / dest_subdir

        # Check if source directory exists
        if not source_subdir_path.exists():
            print(f"❌ Source directory not found: {source_subdir_path}")
            continue

        # Check if destination directory exists
        if not dest_subdir_path.exists():
            print(f"❌ Destination directory not found: {dest_subdir_path}")
            continue

        # Get the only file in source directory
        source_files = [f for f in source_subdir_path.iterdir() if f.is_file()]

        if len(source_files) == 0:
            print(f"⚠️  No files found in: {source_subdir}")
            continue
        elif len(source_files) > 1:
            print(f"⚠️  Multiple files found in {source_subdir}, using first one: {source_files[0].name}")

        source_file = source_files[0]  # Get the first (or only) file

        # Get the only file in destination directory (to know the name to overwrite)
        dest_files = [f for f in dest_subdir_path.iterdir() if f.is_file()]

        if len(dest_files) == 0:
            print(f"⚠️  No files found in destination: {dest_subdir}")
            # Just copy with a default name or skip?
            dest_file_path = dest_subdir_path / source_file.name
        else:
            # Use the existing destination file name (assuming you want to keep the name)
            dest_file_path = dest_files[0]

        # Copy and overwrite
        try:
            shutil.copy2(source_file, dest_file_path)
            print(f"✓ Copied: {source_file.name} → {dest_file_path}")
        except Exception as e:
            print(f"✗ Error copying {source_file.name}: {e}")



if __name__ == "__main__":
    # =========================================================================
    # PATHS para as pastas de destino
    # =========================================================================
    destiny_path_cmof = Path(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\dashboard-nodejs\Orcamento")

    destiny_path_tgov = Path(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\dashboard-nodejs\DATA\TGov")

    destiny_path_acompanhamento = Path(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\dashboard-nodejs\DATA\Acompanhamento\Controle SNEAELIS - 2026.xlsx")

    # =========================================================================
    # PATHs para as pastas fonte
    # =========================================================================
    path_CMOF = Path(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Arquivos de Andre Luiz de Oliveira Santos - 2026")

    path_acompanhamento = Path(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Assessoria\2026\001 - Controle\Controle SNEAELIS - 2026.xlsx")

    path_tgov = Path(r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\webscraping"
)


    # =========================================================================
    # --- CMOF ---
    # =========================================================================
    copy_todays_files_cmof(source_dir=path_CMOF, dest_dir=destiny_path_cmof)

    # =========================================================================
    # --- Acompanhamento ---
    # =========================================================================
    shutil.copy2(path_acompanhamento, destiny_path_acompanhamento)

    # =========================================================================
    # --- TGov ---
    # =========================================================================
    copy_tgov_files(source_dir=path_tgov, dest_dir=destiny_path_tgov)