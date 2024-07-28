import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl import load_workbook

def process_excel(file_path):
    try:
        xlsx = pd.read_excel(file_path, sheet_name=['Report', 'Criteria'])
        Report_df = xlsx['Report']
        Criteria_df = xlsx['Criteria']

        Report_df = Report_df.drop(Report_df.tail(5).index)
        Report_df['Accepted'] = ""
        Report_df['Rejected'] = ""
        Report_df['Reason for Rejection'] = ""

        accepted_loans = ['Term Loan', 'Term Loan A', 'Term Loan B', 'Term Loan C']
        Report_df.loc[Report_df['Tranche Type'].isin(accepted_loans), 'Accepted'] = 'X'
        Report_df.loc[~Report_df['Tranche Type'].isin(accepted_loans), 'Rejected'] = 'X'

        def copy_reason(row):
            if row['Rejected'] == 'X':
                return row['Tranche Type']
            else:
                return ""

        Report_df['Reason for Rejection'] = Report_df.apply(copy_reason, axis=1)

        Accepted_df = Report_df[Report_df['Accepted'] == 'X']
        Rejected_df = Report_df[Report_df['Rejected'] == 'X']

        columns_to_keep = ['Tranche Active Date', 'Tranche Maturity Date', 'Tranche Amount (m)', 'Tranche Currency', 'Base Rate & Margin (bps)']
        Swaps_df = Accepted_df[columns_to_keep]

        # Prompt the user to choose a save location
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        
        if output_file:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                Criteria_df.to_excel(writer, sheet_name='Criteria', index=False)
                Report_df.to_excel(writer, sheet_name='Workings', index=False)
                Accepted_df.to_excel(writer, sheet_name='Accepted', index=False)
                Swaps_df.to_excel(writer, sheet_name='SWAPS', index=False)
                Rejected_df.to_excel(writer, sheet_name='Rejected', index=False)

            messagebox.showinfo("Success", f"Processed file saved as '{output_file}'")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def open_file_dialog():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        process_excel(file_path)

def main():
    root = tk.Tk()
    root.title("Loan Interest Rate Processor")

    canvas = tk.Canvas(root, width=300, height=300)
    canvas.pack()

    label = tk.Label(root, text="Drag and drop an Excel file here")
    label.pack()

    button = tk.Button(root, text="Browse", command=open_file_dialog)
    button.pack()

    def drop(event):
        file_path = event.data
        process_excel(file_path)

    root.mainloop()

if __name__ == "__main__":
    main()

