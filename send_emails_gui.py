import threading
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from mailer import send_campaign

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

def run_campaign(opts):
    try:
        send_campaign(**opts)
    except Exception as e:
        log.exception('Campaign error: %s', e)

def start_send(entries, dry_var):
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
    threading.Thread(target=run_campaign, args=(opts,), daemon=True).start()

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
        tk.Label(root, text=label).grid(row=idx, column=0, sticky='e', padx=5, pady=2)
        ent = tk.Entry(root, width=40)
        ent.grid(row=idx, column=1, sticky='w', padx=5, pady=2)
        entries[key] = ent
        if key == 'leads':
            btn = tk.Button(root, text='Browse', command=lambda e=ent: browse_file(e))
            btn.grid(row=idx, column=2, padx=5)

    dry_var = tk.IntVar(value=0)
    tk.Checkbutton(root, text='Dry run', variable=dry_var).grid(row=len(fields), column=1, sticky='w', pady=4)

    send_btn = tk.Button(root, text='Send campaign', command=lambda: start_send(entries, dry_var))
    send_btn.grid(row=len(fields)+1, column=1, pady=4)

    log_text = tk.Text(root, height=15, width=60)
    log_text.grid(row=len(fields)+2, column=0, columnspan=3, padx=5, pady=5)
    th = TextHandler(log_text)
    logging.getLogger().addHandler(th)

    root.mainloop()

if __name__ == '__main__':
    main()
