import os
import sys
import tabula
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

jar_path = os.path.join(os.path.dirname(__file__), 'tabula-1.0.5-jar-with-dependencies.jar')

# Function to process PDF to Excel with progress
def convert_pdf_to_excel(pdf_file, output_file, progress_bar, progress_label):
    column_names = ['Inv.No', 'F.Dist', 'TWC.No', 
                    'Init.Proc.Date', 'Vehicle ID Number(WMI)', 
                    'Vehicle ID Number(VDS)', 'Vehicle ID Number(C/D)', 
                    'Vehicle ID Number(VIS)', 
                    'Amount Approved(Parts)', 'Amount Approved(Labor)', 
                    'Amount Approved(Sublet)', 'Amount Approved(Total)', 
                    'TAX Approved', 'Amount Claimed', 'Reason Code']

    # Read the PDF
    tables = tabula.read_pdf(pdf_file, pages='all', 
                             multiple_tables=True,
                             relative_area=True, 
                             relative_columns=True,
                             area=[30, 0, 100, 100],
                             columns=[9, 12.5, 16, 21.5, 23, 26.5, 27.5, 32.5, 40, 50, 55, 62.5, 70, 75.5],
                             guess=False, 
                             pandas_options={'header': None})

    df_combined = pd.DataFrame(columns=column_names)
    
    total_tables = len(tables)
    
    # Process each table and update progress
    for idx, table in enumerate(tables):
        df_table = pd.DataFrame(table)
        
        # Filter valid rows
        valid_row_pattern = r"^[A-Z0-9]{7}"
        df_table = df_table[df_table[0].astype(str).str.match(valid_row_pattern)]
        
        # Ensure missing columns
        if len(df_table.columns) < len(column_names):
            for col in range(len(df_table.columns), len(column_names)):
                df_table[col] = pd.NA
        
        # Assign column names
        df_table.columns = column_names
        
        # Convert numeric columns to proper numeric format
        numeric_columns = ['F.Dist', 'TWC.No', 'Init.Proc.Date', 'Amount Approved(Parts)', 
                           'Amount Approved(Labor)', 'Amount Approved(Sublet)', 
                           'Amount Approved(Total)', 'Amount Claimed']
        df_table[numeric_columns] = df_table[numeric_columns].replace({',': ''}, regex=True)
        df_table[numeric_columns] = df_table[numeric_columns].apply(pd.to_numeric, errors='coerce')

        df_combined = pd.concat([df_combined, df_table], ignore_index=True)
        
        # Update progress bar and label
        progress = ((idx + 1) / total_tables) * 100
        progress_bar['value'] = progress
        progress_label.config(text=f"Processing... {int(progress)}%")
        root.update_idletasks()

    # Save to Excel
    df_combined.to_excel(output_file, index=False)
    messagebox.showinfo("Success", f"PDF converted to Excel and saved as {output_file}")

# Function to browse PDF files and trigger conversion
def browse_files():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                   filetypes=[("Excel files", "*.xlsx")])
        if output_file:
            progress_bar['value'] = 0
            progress_label.config(text="Processing... 0%")
            convert_pdf_to_excel(file_path, output_file, progress_bar, progress_label)

# Create the main window
root = tk.Tk()
root.title("PDF to Excel Converter")

# Set window size
root.geometry("400x300")

# Add label and button to browse files
label = tk.Label(root, text="Convert PDF to Excel", font=("Arial", 14))
label.pack(pady=20)

button = tk.Button(root, text="Browse PDF", command=browse_files, font=("Arial", 12))
button.pack(pady=10)

# Add a progress bar
progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode='determinate')
progress_bar.pack(pady=20)

# Add a label to show percentage
progress_label = tk.Label(root, text="Select a File.", font=("Arial", 12))
progress_label.pack()

# Start the GUI loop
root.mainloop()
