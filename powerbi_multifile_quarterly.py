#!/usr/bin/env python3
"""
Power BI Data Transformer - Multi-File Quarterly Edition (Fixed for text in revenue)
Import 4 separate Excel files and let user SELECT which columns to use
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
from datetime import datetime
import os

class MultiFileQuarterlyTransformer:
    def __init__(self, root):
        self.root = root
        self.root.title("Power BI Multi-File Quarterly - Column Selector")
        self.root.geometry("1200x850")
        
        self.file_data = {}
        self.file_paths = {'Q1': tk.StringVar(), 'Q2': tk.StringVar(), 'Q3': tk.StringVar(), 'Q4': tk.StringVar()}
        self.output_df = None
        self.year_var = tk.StringVar(value='2025')
        
        # Column selections
        self.equipment_col = tk.StringVar()
        self.revenue_col = tk.StringVar()
        self.desc_col = tk.StringVar()
        
        self.create_widgets()
    
    def create_widgets(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Tab 1: Files
        tab1 = ttk.Frame(notebook)
        notebook.add(tab1, text='1. Select Files')
        
        ttk.Label(tab1, text="Select 4 Quarterly Excel Files", font=('Arial', 14, 'bold')).pack(pady=10)
        
        year_frame = ttk.Frame(tab1)
        year_frame.pack(fill='x', padx=10, pady=5)
        ttk.Label(year_frame, text="Year:").pack(side=tk.LEFT, padx=5)
        ttk.Combobox(year_frame, textvariable=self.year_var, values=['2024','2025','2026'], width=10, state='readonly').pack(side=tk.LEFT)
        
        file_frame = ttk.LabelFrame(tab1, text="Files", padding=10)
        file_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        for idx, (q, period) in enumerate([('Q1','Jan-Mar'),('Q2','Apr-Jun'),('Q3','Jul-Sep'),('Q4','Oct-Dec')]):
            ttk.Label(file_frame, text=f"{q} {period}:", font=('Arial',10,'bold'), foreground='blue').grid(row=idx, column=0, sticky='w', padx=5, pady=5)
            ttk.Entry(file_frame, textvariable=self.file_paths[q], width=50).grid(row=idx, column=1, padx=5, pady=5)
            ttk.Button(file_frame, text="Browse", command=lambda quarter=q: self.browse(quarter)).grid(row=idx, column=2, padx=5, pady=5)
        
        ttk.Button(tab1, text="Load Files →", command=self.load_files).pack(pady=10)
        
        self.info1 = scrolledtext.ScrolledText(tab1, height=10)
        self.info1.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Tab 2: Columns
        tab2 = ttk.Frame(notebook)
        notebook.add(tab2, text='2. Select Columns')
        
        ttk.Label(tab2, text="Which Columns to Use?", font=('Arial', 14, 'bold')).pack(pady=10)
        ttk.Label(tab2, text="Tell us which columns contain your data", foreground='blue').pack()
        
        col_frame = ttk.LabelFrame(tab2, text="Column Mapping (REQUIRED)", padding=20)
        col_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(col_frame, text="Equipment ID Column:", font=('Arial',10,'bold'), foreground='red').grid(row=0, column=0, sticky='w', padx=10, pady=10)
        self.equip_combo = ttk.Combobox(col_frame, textvariable=self.equipment_col, width=40, state='readonly')
        self.equip_combo.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(col_frame, text="Revenue Column:", font=('Arial',10,'bold'), foreground='red').grid(row=1, column=0, sticky='w', padx=10, pady=10)
        self.rev_combo = ttk.Combobox(col_frame, textvariable=self.revenue_col, width=40, state='readonly')
        self.rev_combo.grid(row=1, column=1, padx=10, pady=10)
        
        ttk.Label(col_frame, text="Description Column:", font=('Arial',10)).grid(row=2, column=0, sticky='w', padx=10, pady=10)
        self.desc_combo = ttk.Combobox(col_frame, textvariable=self.desc_col, width=40, state='readonly')
        self.desc_combo.grid(row=2, column=1, padx=10, pady=10)
        
        ttk.Button(tab2, text="Continue →", command=self.validate_columns).pack(pady=10)
        
        self.info2 = scrolledtext.ScrolledText(tab2, height=15)
        self.info2.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Tab 3: Transform
        tab3 = ttk.Frame(notebook)
        notebook.add(tab3, text='3. Combine')
        
        ttk.Label(tab3, text="Combine Quarterly Data", font=('Arial', 14, 'bold')).pack(pady=10)
        ttk.Button(tab3, text="Combine & Transform", command=self.transform).pack(pady=5)
        
        self.info3 = scrolledtext.ScrolledText(tab3, height=25)
        self.info3.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Tab 4: Export
        tab4 = ttk.Frame(notebook)
        notebook.add(tab4, text='4. Export')
        
        ttk.Label(tab4, text="Export to Excel", font=('Arial', 14, 'bold')).pack(pady=10)
        ttk.Button(tab4, text="Export", command=self.export).pack(pady=5)
        
        self.info4 = scrolledtext.ScrolledText(tab4, height=25)
        self.info4.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Status
        self.status = tk.StringVar(value="Ready")
        ttk.Label(self.root, textvariable=self.status, relief=tk.SUNKEN, anchor=tk.W).pack(side=tk.BOTTOM, fill=tk.X)
    
    def browse(self, quarter):
        f = filedialog.askopenfilename(title=f"Select {quarter} File", filetypes=[("Excel","*.xlsx *.xls")])
        if f:
            self.file_paths[quarter].set(f)
    
    def load_files(self):
        missing = [q for q in ['Q1','Q2','Q3','Q4'] if not self.file_paths[q].get()]
        if missing:
            messagebox.showwarning("Missing Files", f"Please select: {', '.join(missing)}")
            return
        
        try:
            self.status.set("Loading...")
            self.file_data = {}
            all_cols = set()
            
            for q in ['Q1','Q2','Q3','Q4']:
                df = pd.read_excel(self.file_paths[q].get())
                df.columns = df.columns.str.strip()
                self.file_data[q] = df
                all_cols.update(df.columns)
            
            cols = sorted(list(all_cols))
            self.equip_combo['values'] = [''] + cols
            self.rev_combo['values'] = [''] + cols
            self.desc_combo['values'] = [''] + cols
            
            # Auto-select if possible
            for col in cols:
                cl = col.lower()
                if not self.equipment_col.get() and any(x in cl for x in ['equipment','code','item','id']):
                    self.equipment_col.set(col)
                if not self.revenue_col.get() and any(x in cl for x in ['revenue','price','amount','total']):
                    self.revenue_col.set(col)
                if not self.desc_col.get() and any(x in cl for x in ['description','desc','name']):
                    self.desc_col.set(col)
            
            self.info1.delete('1.0', tk.END)
            info = "FILES LOADED ✅\n\n"
            for q in ['Q1','Q2','Q3','Q4']:
                info += f"{q}: {len(self.file_data[q])} rows\n"
            self.info1.insert('1.0', info)
            
            self.info2.delete('1.0', tk.END)
            self.info2.insert('1.0', f"AVAILABLE COLUMNS IN YOUR FILES:\n\n" + "\n".join(f"{i}. {c}" for i,c in enumerate(cols,1)))
            
            messagebox.showinfo("Success", "Files loaded!\nGo to 'Select Columns' tab")
            self.status.set("Files loaded")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load:\n{e}")
            self.status.set("Error")
    
    def validate_columns(self):
        if not self.equipment_col.get():
            messagebox.showerror("Required", "Equipment ID column is REQUIRED!")
            return
        if not self.revenue_col.get():
            messagebox.showerror("Required", "Revenue column is REQUIRED!")
            return
        
        self.info3.delete('1.0', tk.END)
        preview = f"READY TO COMBINE\n\nEquipment Column: {self.equipment_col.get()}\n"
        preview += f"Revenue Column: {self.revenue_col.get()}\n"
        if self.desc_col.get():
            preview += f"Description Column: {self.desc_col.get()}\n"
        preview += f"\nClick 'Combine & Transform' to proceed"
        self.info3.insert('1.0', preview)
        
        messagebox.showinfo("Ready", "Columns selected!\nGo to 'Combine' tab")
        self.status.set("Ready to combine")
    
    def transform(self):
        if not self.file_data:
            messagebox.showwarning("Warning", "Load files first")
            return
        
        equip = self.equipment_col.get()
        rev = self.revenue_col.get()
        desc = self.desc_col.get()
        
        if not equip or not rev:
            messagebox.showerror("Error", "Equipment and Revenue columns required!")
            return
        
        try:
            self.status.set("Combining...")
            
            all_equip = set()
            for df in self.file_data.values():
                if equip in df.columns:
                    all_equip.update(df[equip].dropna().unique())
            
            year = self.year_var.get()
            records = []
            
            for eq in all_equip:
                rec = {'Equipment Code': eq, 'Description': ''}
                
                for qi, q in enumerate(['Q1','Q2','Q3','Q4'], 1):
                    df = self.file_data[q]
                    
                    if equip not in df.columns:
                        rec[f'{year} Q{qi} Revenue'] = 0.0
                        continue
                    
                    eq_data = df[df[equip] == eq]
                    
                    if len(eq_data) == 0:
                        rec[f'{year} Q{qi} Revenue'] = 0.0
                        continue
                    
                    # Get description
                    if not rec['Description'] and desc and desc in eq_data.columns:
                        val = eq_data[desc].iloc[0]
                        if pd.notna(val):
                            rec['Description'] = str(val)
                    
                    # Get revenue - CONVERT TO NUMERIC FIRST
                    revenue = 0.0
                    if rev in eq_data.columns:
                        # Convert to numeric, replacing any non-numeric with 0
                        revenue_series = pd.to_numeric(eq_data[rev], errors='coerce').fillna(0)
                        revenue = float(revenue_series.sum())
                    
                    # Now we can safely round
                    rec[f'{year} Q{qi} Revenue'] = round(revenue, 2)
                
                rec[f'{year} Revenue'] = sum(rec.get(f'{year} Q{i} Revenue', 0) for i in range(1,5))
                rec['Revenue'] = rec[f'{year} Revenue']
                records.append(rec)
            
            self.output_df = pd.DataFrame(records)
            self.output_df = self.output_df.sort_values(f'{year} Revenue', ascending=False)
            
            self.info3.delete('1.0', tk.END)
            stats = f"COMBINED! ✅\n\nEquipment Items: {len(self.output_df)}\n\n"
            for qi in range(1,5):
                col = f'{year} Q{qi} Revenue'
                if col in self.output_df.columns:
                    stats += f"Q{qi} Revenue: ${self.output_df[col].sum():,.2f}\n"
            stats += f"\nTotal: ${self.output_df[f'{year} Revenue'].sum():,.2f}"
            self.info3.insert('1.0', stats)
            
            self.info4.delete('1.0', tk.END)
            prev_cols = ['Equipment Code','Description'] + [f'{year} Q{i} Revenue' for i in range(1,5)] + [f'{year} Revenue']
            prev_cols = [c for c in prev_cols if c in self.output_df.columns]
            self.info4.insert('1.0', self.output_df[prev_cols].head(20).to_string())
            
            messagebox.showinfo("Success", "Data combined!\nGo to 'Export' tab")
            self.status.set("Combined successfully")
            
        except Exception as e:
            messagebox.showerror("Error", f"Transform failed:\n{e}")
            import traceback
            traceback.print_exc()
            self.status.set("Transform error")
    
    def export(self):
        if self.output_df is None:
            messagebox.showwarning("Warning", "Transform data first")
            return
        
        year = self.year_var.get()
        fname = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel","*.xlsx")],
            initialfile=f"PowerBI_Quarterly_{year}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )
        
        if fname:
            try:
                with pd.ExcelWriter(fname, engine='openpyxl') as writer:
                    title = f"Quarterly Data {year} - {datetime.now().strftime('%Y-%m-%d')}"
                    title_df = pd.DataFrame([[title] + ['']*(len(self.output_df.columns)-1)])
                    title_df.to_excel(writer, sheet_name='Quarterly Data', index=False, header=False)
                    self.output_df.to_excel(writer, sheet_name='Quarterly Data', startrow=1, index=False)
                
                messagebox.showinfo("Success", f"Exported!\n\n{os.path.basename(fname)}\n{len(self.output_df)} items\n${self.output_df[f'{year} Revenue'].sum():,.2f}")
                self.status.set(f"Exported: {fname}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Export failed:\n{e}")

def main():
    root = tk.Tk()
    app = MultiFileQuarterlyTransformer(root)
    root.mainloop()

if __name__ == "__main__":
    main()
