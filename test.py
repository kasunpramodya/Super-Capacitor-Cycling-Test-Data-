import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import re
import os

# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
#                 BEGIN DATA PROCESSING LOGIC
# This function encapsulates your entire pandas script.
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---

def process_battery_data(excel_file_path):
    """
    Processes the battery data from the given Excel file.
    
    This function contains the logic from your script to find
    CC charge/discharge cycles and calculate ESR.
    
    Args:
        excel_file_path (str): The path to the .xlsx file.

    Returns:
        tuple: (merged_df, average_esr, error_message)
               On success, (DataFrame, float, None)
               On failure, (None, None, str)
    """
    try:
        df = pd.read_excel(excel_file_path, engine="openpyxl")
    except FileNotFoundError:
        return None, None, f"Error: File not found at {excel_file_path}"
    except Exception as e:
        return None, None, f"Error reading Excel file: {e}"

    # --- Normalize column names ---
    df.columns = df.columns.str.strip().str.lower()

    # --- Identify useful columns automatically (with error handling) ---
    try:
        cycle_col = [c for c in df.columns if "cycle" in c and "index" in c][0]
        voltage_col = [c for c in df.columns if "voltage" in c][0]
        current_col = [c for c in df.columns if "current" in c][0]
        step_col = [c for c in df.columns if "step" in c and "type" in c][0]
    except IndexError:
        error_msg = ("Error: Could not find required columns.\n"
                     "Please ensure your file has columns containing:\n"
                     "- 'cycle' and 'index'\n"
                     "- 'voltage'\n"
                     "- 'current'\n"
                     "- 'step' and 'type'")
        return None, None, error_msg

    # --- Ensure numeric for safety ---
    df[voltage_col] = pd.to_numeric(df[voltage_col], errors="coerce")
    df[current_col] = pd.to_numeric(df[current_col], errors="coerce")
    
    # Drop rows where conversion failed
    df = df.dropna(subset=[voltage_col, current_col])

    # =========================
    # CC CHARGE
    # =========================
    mask_cc_chg = df[step_col].astype(str).str.contains(
        r'(?:\bcc\b.*(?:chg|charge))|(?:constant\s*current\s*charge)',
        flags=re.IGNORECASE, regex=True
    )
    df_cc_chg = df[mask_cc_chg]
    
    if df_cc_chg.empty:
        return None, None, "Error: No 'Constant Current Charge' steps found."

    max_v_cc_chg = df_cc_chg.groupby(cycle_col)[voltage_col].max().reset_index()
    max_i_cc_chg = df_cc_chg.groupby(cycle_col)[current_col].max().reset_index()

    # ============================
    # CC DISCHARGE
    # ============================
    mask_cc_dchg = df[step_col].astype(str).str.contains(
        r'(?:\bcc\b.*(?:dchg|disch|discharge))|(?:constant\s*current\s*discharge)',
        flags=re.IGNORECASE, regex=True
    )
    df_cc_dchg = df[mask_cc_dchg]

    if df_cc_dchg.empty:
        return None, None, "Error: No 'Constant Current Discharge' steps found."

    max_v_cc_dchg = df_cc_dchg.groupby(cycle_col)[voltage_col].max().reset_index()
    max_i_cc_dchg = df_cc_dchg.groupby(cycle_col)[current_col].max().reset_index()

    # === Merge all into one DataFrame ===
    merged = (max_v_cc_chg
              .merge(max_v_cc_dchg, on=cycle_col, how="outer", suffixes=("_chg", "_dchg_v"))
              .merge(max_i_cc_chg, on=cycle_col, how="outer")
              .merge(max_i_cc_dchg, on=cycle_col, how="outer", suffixes=("_chg", "_dchg_i")))
    
    # Rename columns clearly
    merged.columns = [
        "Cycle Index",
        "Max Voltage CC-Chg (V)",
        "Max Voltage CC-DChg (V)",
        "Max Current CC-Chg (A)",
        "Max Current CC-DChg (A)"
    ]
    
    merged = merged.sort_values("Cycle Index").dropna()
    
    if merged.empty:
        return None, None, "Error: No matching charge/discharge cycles found."

    # === Compute Average ESR ===
    dv = (merged["Max Voltage CC-Chg (V)"] - merged["Max Voltage CC-DChg (V)"]).mean()
    di = (merged["Max Current CC-Chg (A)"] - merged["Max Current CC-DChg (A)"]).mean()

    if di == 0:
        average_esr = float('nan') # Avoid ZeroDivisionError
    else:
        average_esr = dv / di

    # === Compute per-cycle ESR ===
    # Avoid division by zero
    merged["ESR (Ω)"] = (merged["Max Voltage CC-Chg (V)"] - merged["Max Voltage CC-DChg (V)"]) / \
                      (merged["Max Current CC-Chg (A)"] - merged["Max Current CC-DChg (A)"])
    
    # Round for display
    merged = merged.round(6)

    return merged, average_esr, None

# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
#                  BEGIN TKINTER GUI
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---

class BatteryAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Battery ESR Analyzer")
        self.root.geometry("800x600")
        self.root.minsize(500, 400)

        # --- Style ---
        self.style = ttk.Style(self.root)
        try:
            self.style.theme_use('clam') 
        except tk.TclError:
            self.style.theme_use('default')

        # Configure styles
        self.style.configure("TFrame", background="#f1f1f1")
        self.style.configure("TLabel", background="#f1f1f1", font=("Inter", 10))
        self.style.configure("Title.TLabel", font=("Inter", 16, "bold"))
        self.style.configure("Result.TLabel", font=("Inter", 12, "bold"), foreground="#00529B")
        self.style.configure("TButton", font=("Inter", 10, "bold"), padding=5)
        self.style.configure("Treeview.Heading", font=("Inter", 10, "bold"))
        self.style.configure("Treeview", rowheight=25, font=("Inter", 10))
        self.root.configure(background="#f1f1f1")

        # --- Variables ---
        self.selected_file = tk.StringVar(value="No file selected.")
        self.average_esr = tk.StringVar(value="Average ESR: N/A")
        self.results_df = None # To store the results dataframe for export

        # --- Main Frame ---
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Top Frame (Controls) ---
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=5)

        ttk.Label(top_frame, text="Battery Data Analyzer", style="Title.TLabel").pack(side=tk.LEFT, anchor="w", padx=5)
        
        load_button = ttk.Button(top_frame, text="Load Excel File", command=self.load_file)
        load_button.pack(side=tk.RIGHT, anchor="e", padx=5)

        # --- File Label Frame ---
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=5)
        ttk.Label(file_frame, textvariable=self.selected_file, font=("Inter", 9, "italic")).pack(side=tk.LEFT, padx=5)

        # --- Result Frame ---
        result_frame = ttk.Frame(main_frame, padding=10)
        result_frame.pack(fill=tk.X)
        
        ttk.Label(result_frame, textvariable=self.average_esr, style="Result.TLabel").pack(side=tk.LEFT)
        
        self.export_button = ttk.Button(result_frame, text="Export Results", command=self.export_results, state="disabled")
        self.export_button.pack(side=tk.RIGHT, anchor="e", padx=5)

        # --- Treeview Frame (for the table) ---
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Scrollbars
        scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)

        self.tree = ttk.Treeview(tree_frame, 
                                 yscrollcommand=scroll_y.set, 
                                 xscrollcommand=scroll_x.set, 
                                 show="headings")
        
        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)

        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def load_file(self):
        """
        Open a file dialog to select an Excel file and trigger processing.
        """
        file_path = filedialog.askopenfilename(
            title="Select a Battery Data File",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        
        if not file_path:
            return # User cancelled

        # Update GUI
        self.selected_file.set(f"Processing: {os.path.basename(file_path)}...")
        self.average_esr.set("Average ESR: Calculating...")
        self.export_button.config(state="disabled") # Disable while processing
        self.results_df = None # Clear old results
        self.root.update_idletasks() # Force GUI update

        # Run the processing
        df, avg_esr, error = process_battery_data(file_path)

        # Handle results
        if error:
            messagebox.showerror("Processing Error", error)
            self.selected_file.set(f"Error. Please try again.")
            self.average_esr.set("Average ESR: N/A")
            self.results_df = None # Ensure results are cleared
            self.export_button.config(state="disabled") # Keep disabled
            # Clear treeview on error
            self.populate_treeview(pd.DataFrame()) 
        else:
            self.selected_file.set(f"Loaded: {os.path.basename(file_path)}")
            self.average_esr.set(f"Average ESR: {avg_esr:.6f} Ω")
            self.results_df = df # Store results
            self.export_button.config(state="normal") # Enable button
            self.populate_treeview(df)

    def populate_treeview(self, df):
        """
        Clear and populate the Treeview widget with DataFrame data.
        """
        # Clear existing data
        for i in self.tree.get_children():
            self.tree.delete(i)
        
        if df.empty:
            self.tree["columns"] = []
            return

        # --- Setup new columns ---
        self.tree["columns"] = list(df.columns)
        
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=120)

        # --- Insert data ---
        # THIS LOOP WAS DUPLICATED. It is now fixed.
        for index, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def export_results(self):
        """
        Save the processed results DataFrame to an Excel or CSV file.
        """
        if self.results_df is None or self.results_df.empty:
            messagebox.showwarning("No Data", "There is no data to export. Please load a file first.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Save Results As",
            initialfile="esr_results.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), 
                       ("CSV files", "*.csv"),
                       ("All files", "*.*")]
        )

        if not file_path:
            return # User cancelled

        try:
            if file_path.endswith('.xlsx'):
                self.results_df.to_excel(file_path, index=False, engine='openpyxl')
            elif file_path.endswith('.csv'):
                self.results_df.to_csv(file_path, index=False)
            else:
                # Default to excel if unsure
                self.results_df.to_excel(file_path + ".xlsx", index=False, engine='openpyxl')

            messagebox.showinfo("Export Successful", f"Results successfully saved to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred while saving the file:\n{e}")

# --- Main execution ---
if __name__ == "__main__":
    root = tk.Tk()
    app = BatteryAnalyzerApp(root)
    root.mainloop()
  