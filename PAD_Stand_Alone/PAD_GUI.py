import threading
import os
import json
import sys

import PAD_Exec

import tkinter as tk

from tkinter import filedialog as fd
from tkinter import ttk, messagebox as mb

# Global variables
captcha_resume_btn = None
automation_instance = None


def handle_captcha_event(event_type):
    """This function gets called by the automation class when CAPTCHA is detected"""
    global captcha_resume_btn
    if event_type == 'captcha_detected':
        window.after(0, show_captcha_warning)


def show_captcha_warning():
    """Show CAPTCHA warning and enable resume button"""
    mb.showwarning(
        "CAPTCHA Detectado",
        "Por favor, resolva o CAPTCHA manualmente no navegador.\n\n"
        "Após resolver, clique em 'Continuar' para retomar a automação."
    )
    # Show and enable the resume button
    captcha_resume_btn.pack(pady=10)
    captcha_resume_btn.config(state=tk.NORMAL)
    # Disable other controls during CAPTCHA
    run_script_btn.config(state=tk.DISABLED)
    select_btn.config(state=tk.DISABLED)
    clear_bnt.config(state=tk.DISABLED)
    print("⏸️ Script pausado - aguardando resolução do CAPTCHA...")


def resume_after_captcha():
    """Resume automation after user confirms CAPTCHA is solved"""
    global automation_instance
    if automation_instance:
        automation_instance.resume_after_captcha()
    # Hide resume button and re-enable controls
    captcha_resume_btn.pack_forget()
    run_script_btn.config(state=tk.NORMAL)
    select_btn.config(state=tk.NORMAL)
    clear_bnt.config(state=tk.NORMAL)
    print("▶️ Retomando execução do script...")


def select_file():
    """Open a file dialog and get the selected file path."""
    global selected_files

    file_types = (('Excel files', '*.xlsx'),)

    paths = fd.askopenfilenames(
        title='Selecionar arquivo Excel',
        initialdir='/',
        filetypes=file_types
    )

    if not paths:
        return

    if all(p.lower().endswith('.xlsx') for p in paths):
        file_path.set(';'.join(paths))
        run_script_btn.config(state=tk.NORMAL)

        files = file_path.get().split(';') if file_path.get() else []
        selected_files = files

        if not files:
            display = ''
            print_msg = 'Nenhum arquivo selecionado.'
        else:
            first_basename = os.path.basename(files[0])
            if len(files) == 1:
                display = first_basename
                print_msg = f'✅ Arquivo selecionado: {first_basename}.Você pode rodar o script agora.'
            else:
                display = f'{first_basename} (+{len(files)-1} outros)'
                print_msg = (f'✅ {len(files)} arquivos selecionados. Primeiro arquivo: {first_basename}. '
                             f'Você pode rodar o script agora')

        # Update the entry to show just the filename for cleaner look
        file_path_entry.config(state=tk.NORMAL)
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, display)
        file_path_entry.config(state='readonly')

        print(print_msg)

    else:
        # If the user cancels, keep the run button disabled
        mb.showerror(
            "Wrong File Type",
            "Please select a valid Excel file (.xlsx) to continue."
        )
        run_script_btn.config(state=tk.DISABLED)
        print("❌ No file selected.")


# Runs the automatic assistant to fill in the PAD
def run_assistant():
    """Runs the rest of the script after the user confirms their login."""
    global automation_instance

    # Button configuration
    run_script_btn.config(state=tk.DISABLED)
    select_btn.config(state=tk.DISABLED)  # Disable select during run
    clear_bnt.config(state=tk.DISABLED) # Disable clear during run

    files = selected_files

    # Checks if there files
    if not files:
        mb.showwarning("Aviso", "Por favor, escolha um arquivo primeiro.")
        run_script_btn.config(state=tk.NORMAL)
        select_btn.config(state=tk.NORMAL)
        clear_bnt.config(state=tk.NORMAL)
        return

    missing_files = [mf for mf in files if not os.path.exists(mf)]

    # Double-check that file exists
    if missing_files:
        mb.showerror("Erro", "Os seguintes arquivos não existem:\n" +
                     "\n".join(os.path.basename(m) for m in missing_files))
        run_script_btn.config(state=tk.NORMAL)
        select_btn.config(state=tk.NORMAL)
        clear_bnt.config(state=tk.NORMAL)
        return

    # Initialize automation with callback
    automation_instance = PAD_Exec.Robo(gui_callback=handle_captcha_event)
    try:
        def target_wrapper():
            try:
                automation_instance.main(files=files)
                # Re-enable buttons after completion
                window.after(0, lambda: run_script_btn.config(state=tk.NORMAL))
                window.after(0, lambda: select_btn.config(state=tk.NORMAL))
                window.after(0, lambda: clear_bnt.config(state=tk.NORMAL))

                print("✅ Automação concluída com sucesso!")

            except Exception as e:
                print(f'❌ Erro durante a automação: {type(e).__name__}: {str(e)[:100]}')
                window.after(0, lambda: run_script_btn.config(state=tk.NORMAL))
                window.after(0, lambda: select_btn.config(state=tk.NORMAL))
                window.after(0, lambda: clear_bnt.config(state=tk.NORMAL))

        automation_thread = threading.Thread(
            target=target_wrapper,
            daemon=True
        )
        automation_thread.start()

    except Exception as e:
        print(f'❌ An unexpected error has occurred {type(e).__name__}\n {str(e)[:100]}')
        run_script_btn.config(state=tk.NORMAL)
        select_btn.config(state=tk.NORMAL)
        clear_bnt.config(state=tk.NORMAL)



def exit_app():
    """Close the application window and browser."""
    if automation_instance:
        try:
            print("🧹 Cleaning up resources before exit...")

            automation_instance.close()

        except Exception as e:
            print(f"⚠️ Cleanup warning: {str(e)[:100]}")

    window.destroy()
    print("👋 Application closed")


def clear_selection():
    """Clear the current file selection"""
    run_script_btn.config(state=tk.DISABLED)
    file_path.set('')
    file_path_entry.config(state=tk.NORMAL)
    file_path_entry.delete(0, tk.END)
    file_path_entry.config(state='readonly')
    file_path_entry.config(state=tk.DISABLED)
    select_btn.focus()
    print("🗑️ Seleção de arquivo limpa")


def update_status(*args):
    if file_path.get():
        status_label.config(text=f'Pronto para executar: {os.path.basename(file_path.get())}',
                            foreground='green')
    else:
        status_label.config(text='Selecione um arquivo Excel para começar')


# --- Main Application Window --
window = tk.Tk()
window.title("Assistente de preenchimento de PAD ")
window.resizable(False, False)

# Set the theme for a retro look
style = ttk.Style()
style.theme_use('clam')
# make widget backgrounds match the window so no "boxed" areas appear
bg = window.cget('background')
style.configure('TFrame', background=bg)
style.configure('TLabel', background=bg)
style.configure('TEntry', fieldbackground=bg, background=bg)
style.configure('Flat.TButton', background=bg, border_width=0)
style.map('Flat.TButton',
          background=[('active', bg), ('pressed', bg), ('!disabled', bg)],
          foregorund=[('active', 'black'), ('pressed', 'black'), ('!disabled', 'black')]
          )

# Center the window on the screen
window_width = 750
window_height = 550
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
center_x = int(screen_width / 2 - window_width / 2)
center_y = int(screen_height / 2 - window_height / 2)
window.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

# Create a central frame to hold the buttons
main_frame = tk.Frame(window, padx=20, pady=20)
main_frame.pack(expand=True, fill='both')

# Title label
title_lable = ttk.Label(main_frame,
                        text='Assistente de Preenchimento do PAD',
                        font=('Courier New', 16, 'bold'))
title_lable.pack(pady=(3, 6))

# Status Label
status_label = ttk.Label(main_frame,
                         text='Selecione um arquivo Excel para começar',
                         font=('Courier New', 10)
                        )
status_label.pack(pady=(0,10))

# Variable to hold the file path
selected_files = []
file_path = tk.StringVar()
file_path.trace_add('write', update_status)

# Entry for file path
file_path_entry = tk.Entry(main_frame, textvariable=file_path, width=70, state='readonly')
file_path_entry.pack(pady=10)

# Frame to hold the file operation buttons
file_buttons_frame = tk.Frame(main_frame)
file_buttons_frame.pack(pady=10)

# Button for file selection, excel(.xlsx) only
select_btn = tk.Button(file_buttons_frame, text="Escolher Arquivo", command=select_file)
select_btn.pack(padx=5, pady=(0, 6))

# Clear selection button
clear_bnt = tk.Button(file_buttons_frame, text="🗑️ Limpar", command=clear_selection)
clear_bnt.pack(padx=5)

# Button to start automation
run_script_btn = tk.Button(main_frame, text="Executar assistente", command=run_assistant, state=tk.DISABLED)
run_script_btn.pack(pady=(70, 0))

# Resume button (use ttk for consistency)
# Don't pack it yet - we'll show it only when CAPTCHA is detected
captcha_resume_btn = ttk.Button(
    main_frame,
    text='▶️ Continuar após CAPTCHA',
    command=resume_after_captcha,  # No parentheses!
    state=tk.DISABLED
)

# Enable an exit button
exit_frame = ttk.Frame(window)
exit_frame.pack(side='bottom', fill='x')
exit_btn = ttk.Button(exit_frame, text='SAIR', command=exit_app)
exit_btn.pack(side='right', padx=10, pady=5)

window.mainloop()