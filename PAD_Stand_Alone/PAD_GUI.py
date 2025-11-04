import PAD_Exec
import threading
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
    print("▶️ Retomando execução do script...")

def select_file():
    """Open a file dialog and get the selected file path."""
    file_types = (('Excel files', '*.xlsx'),)
    file_name = fd.askopenfilename(
        title='Abrir um arquivo',
        initialdir='/',
        filetypes=file_types
    )
    if file_name.endswith('.xlsx'):
        file_path.set(file_name)
        run_script_btn.config(state=tk.NORMAL)
        print("✅ File selected. You can now run the script.")
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
    run_script_btn.config(state=tk.DISABLED)
    select_btn.config(state=tk.DISABLED)  # Disable select during run
    file_path_source = file_path.get()
    # Initialize automation with callback
    automation_instance = PAD_Exec.Robo(gui_callback=handle_captcha_event)
    try:
        def target_wrapper():
            automation_instance.main(file_path_source)
            # Re-enable buttons after completion
            window.after(0, lambda: run_script_btn.config(state=tk.NORMAL))
            window.after(0, lambda: select_btn.config(state=tk.NORMAL))
            print("✅ Automation completed.")

        automation_thread = threading.Thread(
            target=target_wrapper,
            daemon=True
        )
        automation_thread.start()
    except Exception as e:
        print(f'❌ An unexpected error has occurred {type(e).__name__}\n {str(e)[:100]}')
        run_script_btn.config(state=tk.NORMAL)
        select_btn.config(state=tk.NORMAL)

def exit_app():
    """Close the application window and browser."""
    window.destroy()

# --- Main Application Window --
window = tk.Tk()
window.title("Assistente de preenchimento de PAD ")

# Set the theme for a retro look
style = ttk.Style()
style.theme_use('clam')

# Center the window on the screen
window_width = 750
window_height = 550
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
center_x = int(screen_width / 2 - window_width / 2)
center_y = int(screen_height / 2 - window_height / 2)
window.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

# Create a central frame to hold the buttons
main_frame = tk.Frame(window)
main_frame.pack(expand=True, fill='both')

# Variable to hold the file path
file_path = tk.StringVar()

# Create and place widgets
file_path_entry = ttk.Entry(main_frame, textvariable=file_path, width=60)
file_path_entry.pack(pady=(80, 0))

# Button for file selection, excel(.xlsx) only
select_btn = ttk.Button(main_frame, text="Escolher Arquivo", command=select_file)
select_btn.pack(side='top', pady=5)

# Button to start automation
run_script_btn = ttk.Button(main_frame, text="Executar assistente", command=run_assistant, state=tk.DISABLED)
run_script_btn.pack(pady=(70, 0))

# Resume button (use ttk for consistency)
# Don't pack it yet - we'll show it only when CAPTCHA is detected
captcha_resume_btn = ttk.Button(
    main_frame,
    text='Continuar após CAPTCHA',
    command=resume_after_captcha,  # No parentheses!
    state=tk.DISABLED
)

# Enable an exit button
exit_frame = ttk.Frame(window)
exit_frame.pack(side='bottom', fill='x')
exit_btn = ttk.Button(exit_frame, text='SAIR', command=exit_app)
exit_btn.pack(side='right', padx=10, pady=5)

window.mainloop()