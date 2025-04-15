#calculates Langmuir Adsorption Isotherm parameters (monolayer capacity, qm and Langmuir constant KL) from data in an excel file (MUST BE PLACED IN THE SAME FOLDER), exports the data to a .csv file for Veusz, then generates a .vsz project file that plots the Langmuir fit with the equation and R^2

import numpy as np
import pandas as pd
from scipy.stats import linregress
import veusz.embed as veusz_embed
import os

# === Settings ===
excel_file = "pp5 data 2025.xlsx"
sheet_name = "Sheet2"
output_csv = "langmuir_data.csv"
veusz_project = "langmuir_plot.vsz"
MW_pp5 = 243.29         # Molar mass in g/mol
volume_L = 0.025        # L
mass_g = 0.2            # g

# === Read Excel ===
df = pd.read_excel(excel_file, sheet_name=sheet_name)
C_e = df.iloc[:, 2].dropna().to_numpy()          # mol/L
C_adsorbed = df.iloc[:, 3].dropna().to_numpy()   # mol/L

# Truncate to same length
min_len = min(len(C_e), len(C_adsorbed))
C_e = C_e[:min_len]
C_adsorbed = C_adsorbed[:min_len]

# Calculate qe in mg/g
q_e_mg_g = (C_adsorbed * volume_L * MW_pp5 * 1000) / mass_g

# Langmuir transformation
inv_Ce = 1 / C_e
inv_qe = 1 / q_e_mg_g

# Linear fit
slope, intercept, r_value, _, _ = linregress(inv_Ce, inv_qe)
q_m = 1 / intercept
K_L = slope / intercept
r_squared = r_value**2

# Export CSV for Veusz
export_df = pd.DataFrame({
    "inv_Ce": inv_Ce,
    "inv_qe": inv_qe,
    "fit_qe": slope * inv_Ce + intercept
})
export_df.to_csv(output_csv, index=False)

# === Create Veusz Plot ===
doc = veusz_embed.Embedded("langmuir_plot")

doc.To("root")
doc.Add("graph", name="langmuir")
doc.To("langmuir")
doc.Add("xy", name="data", autoadd=False)
doc.Set("data", "xData", "inv_Ce")
doc.Set("data", "yData", "inv_qe")
doc.Set("data", "marker", "circle")
doc.Set("data", "PlotLine/width", "0pt")

doc.Add("xy", name="fit", autoadd=False)
doc.Set("fit", "xData", "inv_Ce")
doc.Set("fit", "yData", "fit_qe")
doc.Set("fit", "PlotLine/color", "red")
doc.Set("fit", "PlotLine/width", "1.5pt")
doc.Set("fit", "marker", "none")

doc.Set("x/label", "1 / Ce (L/mol)")
doc.Set("y/label", "1 / qe (g/mg)")
doc.Set("title", "Langmuir Isotherm (Linearized)")

# Add equation and R²
doc.Add("labels", name="annotation")
doc.Set("annotation", "label", f"1/qe = {slope:.4f}(1/Ce) + {intercept:.4f}\nR² = {r_squared:.4f}")
doc.Set("annotation", "xPos", min(inv_Ce))
doc.Set("annotation", "yPos", max(inv_qe))
doc.Set("annotation", "Text/size", "12pt")

# Link to data
doc.SetDataFile(output_csv)

# Export VSZ
doc.Export(veusz_project)
doc.Save(veusz_project)

# Done
print(f"✅ Langmuir plot saved as Veusz project: {veusz_project}")
print(f"→ Monolayer capacity (q_m): {q_m:.2f} mg/g")
print(f"→ Langmuir constant (K_L): {K_L:.2f} L/mg")
print(f"→ R² (fit quality): {r_squared:.4f}")