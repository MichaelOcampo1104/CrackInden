import tkinter as tk
from tkinter import filedialog
import subprocess
from tkinter import simpledialog

def browse_input():
    input_path.set(filedialog.askopenfilename(title='Select Input PDF', filetypes=[('PDF files', '*.pdf')]))

def browse_output():
    output_path.set(filedialog.asksaveasfilename(title='Select Output Excel', defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx')]))

def run_program():
    input_pdf = input_path.get()
    output_excel = output_path.get()
    if input_pdf and output_excel:
        print('Running the main program...')
        print('Input PDF:', input_pdf)
        print('Output Excel:', output_excel)
        subprocess.run(['python', r'C:\Users\hp\Documents\Python Scripts\crackinden\mainV.01.py', '--input', input_pdf, '--output', output_excel])
    else:
        print('Please select both input and output files.')

root = tk.Tk()
root.title('PDF to Excel Converter')
input_path = tk.StringVar()
output_path = tk.StringVar()

tk.Label(root, text='Input PDF:').grid(row=0, column=0, sticky='e')
tk.Entry(root, textvariable=input_path).grid(row=0, column=1)
tk.Button(root, text='Browse', command=browse_input).grid(row=0, column=2)

tk.Label(root, text='Output Excel:').grid(row=1, column=0, sticky='e')
tk.Entry(root, textvariable=output_path).grid(row=1, column=1)
tk.Button(root, text='Browse', command=browse_output).grid(row=1, column=2)

tk.Button(root, text='Run', command=run_program).grid(row=2, columnspan=3)

root.mainloop()