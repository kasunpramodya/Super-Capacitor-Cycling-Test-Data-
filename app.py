import pandas as pd
# --- Normalize column names (strip spaces, lowercase) ---
import re
df = pd.read_excel("testdata.xlsx", engine="openpyxl")
df.columns = df.columns.str.strip().str.lower()

# --- Identify useful columns automatically ---
cycle_col   = [c for c in df.columns if "cycle" in c and "index" in c][0]
voltage_col = [c for c in df.columns if "voltage" in c][0]
current_col = [c for c in df.columns if "current" in c][0]
step_col    = [c for c in df.columns if "step" in c and "type" in c][0]

# --- Ensure numeric for safety ---
df[voltage_col] = pd.to_numeric(df[voltage_col], errors="coerce")
df[current_col] = pd.to_numeric(df[current_col], errors="coerce")

# =========================
# CC CHARGE (your original)
# =========================
mask_cc_chg = df[step_col].astype(str).str.contains(
    r'(?:\bcc\b.*(?:chg|charge))|(?:constant\s*current\s*charge)',
    flags=re.IGNORECASE, regex=True
)
df_cc_chg = df[mask_cc_chg]

max_v_cc_chg = df_cc_chg.groupby(cycle_col)[voltage_col].max().reset_index()
max_v_cc_chg.columns = ["Cycle Index", "Max Voltage CC-Chg (V)"]

max_i_cc_chg = df_cc_chg.groupby(cycle_col)[current_col].max().reset_index()
max_i_cc_chg.columns = ["Cycle Index", "Max Current CC-Chg (A)"]

print(max_v_cc_chg)
print(max_i_cc_chg)

# ============================
# CC DISCHARGE (new, as asked)
# ============================
mask_cc_dchg = df[step_col].astype(str).str.contains(
    r'(?:\bcc\b.*(?:dchg|disch|discharge))|(?:constant\s*current\s*discharge)',
    flags=re.IGNORECASE, regex=True
)
df_cc_dchg = df[mask_cc_dchg]

max_v_cc_dchg = df_cc_dchg.groupby(cycle_col)[voltage_col].max().reset_index()
max_v_cc_dchg.columns = ["Cycle Index", "Max Voltage CC-DChg (V)"]

max_i_cc_dchg = df_cc_dchg.groupby(cycle_col)[current_col].max().reset_index()
max_i_cc_dchg.columns = ["Cycle Index", "Max Current CC-DChg (A)"]

print(max_v_cc_dchg)
print(max_i_cc_dchg)

# ============================
# Optional: combine into one table
# ============================
out = (max_v_cc_chg
       .merge(max_i_cc_chg, on="Cycle Index", how="outer")
       .merge(max_v_cc_dchg, on="Cycle Index", how="outer")
       .merge(max_i_cc_dchg, on="Cycle Index", how="outer")
       .sort_values("Cycle Index"))

print("\n=== Summary per Cycle ===")
print(out)

# Save (Colab-friendly)
out.to_excel("cc_charge_discharge_max_by_cycle.xlsx", index=False)
out.to_csv("cc_charge_discharge_max_by_cycle.csv", index=False)
print("Saved: cc_charge_discharge_max_by_cycle.xlsx / .csv")

# === Merge all into one DataFrame ===
merged = (max_v_cc_chg
          .merge(max_v_cc_dchg, on="Cycle Index", how="outer")
          .merge(max_i_cc_chg, on="Cycle Index", how="outer")
          .merge(max_i_cc_dchg, on="Cycle Index", how="outer")
          .sort_values("Cycle Index"))

# Rename columns clearly
merged.columns = [
    "Cycle Index",
    "Max Voltage CC-Chg (V)",
    "Max Voltage CC-DChg (V)",
    "Max Current CC-Chg (A)",
    "Max Current CC-DChg (A)"
]

# === Compute ESR ===
# (Average of (Vchg - Vdchg)) / (Average of (Ichg - Idchg))
dv = (merged["Max Voltage CC-Chg (V)"] - merged["Max Voltage CC-DChg (V)"]).mean()
di = (merged["Max Current CC-Chg (A)"] - merged["Max Current CC-DChg (A)"]).mean()

esr = dv / di

print("\n=== ESR Calculation ===")
print(f"Average ΔV = {dv:.6f} V")
print(f"Average ΔI = {di:.6f} A")
print(f"ESR = {esr:.6f} Ω")

# Optional: add per-cycle ESR if you want
merged["ESR (Ω)"] = (merged["Max Voltage CC-Chg (V)"] - merged["Max Voltage CC-DChg (V)"]) / \
                    (merged["Max Current CC-Chg (A)"] - merged["Max Current CC-DChg (A)"])

print("\n=== Per-cycle ESR values ===")
print(merged[["Cycle Index", "ESR (Ω)"]])

# Save to Excel
merged.to_excel("esr_results.xlsx", index=False)
print("Saved → esr_results.xlsx")
