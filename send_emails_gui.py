import threading
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from mailer import send_campaign, cfg
import pandas as pd

log = logging.getLogger('gui')
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')

class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        self.text_widget.after(0, self.text_widget.insert, tk.END, msg + '\n')
        self.text_widget.after(0, self.text_widget.see, tk.END)


class ProgressHandler(logging.Handler):
    def __init__(self, progress, status_var, total):
        super().__init__()
        self.progress = progress
        self.status_var = status_var
        self.total = total

    def emit(self, record):
        msg = record.getMessage()
        if msg.startswith(('sent', 'scheduled', 'DRY-RUN')):
            self.progress.after(0, self.progress.step, 1)
            value = int(float(self.progress['value'])) + 1
            self.progress.after(0, self.status_var.set, f"{value} / {self.total}")

def run_campaign(opts, progress, status_var, total):
    ph = ProgressHandler(progress, status_var, total)
    logging.getLogger().addHandler(ph)
    try:
        send_campaign(**opts)
    except Exception as e:
        log.exception('Campaign error: %s', e)
    finally:
        logging.getLogger().removeHandler(ph)
        progress.after(0, status_var.set, f"{total} / {total}")

def start_send(entries, dry_var, progress, status_var):
    opts = {
        'subject_line': entries['subject'].get(),
        'excel_path': entries['leads'].get() or None,
        'template_base': entries['template_base'].get() or None,
        'sheet_name': entries['sheet'].get() or None,
        'send_at': entries['send_at'].get() or 'now',
        'account': entries['account'].get() or None,
        'language_column': entries['language_column'].get() or 'language',
        'dry_run': bool(dry_var.get()),
    }

    leads_file = opts['excel_path'] or cfg['defaults']['default_leads_file']
    sheet = opts['sheet_name'] or cfg['defaults']['sheet_name']
    total = 0
    if leads_file:
        try:
            df = pd.read_excel(leads_file, sheet_name=sheet)
            total = len(df)
        except Exception as e:
            messagebox.showerror('Leads Error', f'Failed to read leads file: {e}')
            return
    progress['value'] = 0
    progress['maximum'] = max(total, 1)
    status_var.set(f"0 / {total}")
    threading.Thread(target=run_campaign, args=(opts, progress, status_var, total), daemon=True).start()

def browse_file(entry):
    path = filedialog.askopenfilename(title='Select Excel file', filetypes=[('Excel','*.xlsx *.xls')])
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

def main():
    root = tk.Tk()
    root.title('Outspamer GUI')

    fields = [
        ('Subject', 'subject'),
        ('Leads file', 'leads'),
        ('Template base', 'template_base'),
        ('Sheet name', 'sheet'),
        ('Send at', 'send_at'),
        ('Account', 'account'),
        ('Language column', 'language_column'),
    ]
    entries = {}
    for idx, (label, key) in enumerate(fields):
        ttk.Label(root, text=label).grid(row=idx, column=0, sticky='e', padx=5, pady=2)
        ent = ttk.Entry(root, width=40)
        ent.grid(row=idx, column=1, sticky='w', padx=5, pady=2)
        entries[key] = ent
        if key == 'leads':
            btn = ttk.Button(root, text='Browse', command=lambda e=ent: browse_file(e))
            btn.grid(row=idx, column=2, padx=5)

    dry_var = tk.IntVar(value=0)
    ttk.Checkbutton(root, text='Dry run', variable=dry_var).grid(row=len(fields), column=1, sticky='w', pady=4)

    progress = ttk.Progressbar(root, length=250, mode='determinate')
    progress.grid(row=len(fields)+1, column=0, columnspan=2, padx=5, pady=4, sticky='we')
    status_var = tk.StringVar(value='0 / 0')
    ttk.Label(root, textvariable=status_var).grid(row=len(fields)+1, column=2)

    send_btn = ttk.Button(root, text='Send campaign', command=lambda: start_send(entries, dry_var, progress, status_var))
    send_btn.grid(row=len(fields)+2, column=1, pady=4)

    log_text = tk.Text(root, height=15, width=60)
    log_text.grid(row=len(fields)+3, column=0, columnspan=3, padx=5, pady=5)
    th = TextHandler(log_text)
    logging.getLogger().addHandler(th)

    root.mainloop()

if __name__ == '__main__':
    main()
