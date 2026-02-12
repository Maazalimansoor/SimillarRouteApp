import pandas as pd

file_path = r"C:\Users\maaz.mansoor\OneDrive - AlpHa Measurement Solutions\Desktop\SimillarRouteApp\Route File Lean.xlsx"

# Load sheets
apnrn = pd.read_excel(file_path, sheet_name="apnrn")       # part → route
rodetail = pd.read_excel(file_path, sheet_name="rodetail") # route operations

# Pivot operations: one row per route
ro_pivot = rodetail.pivot(index='routeno', columns='opno', values='cycletime')
ro_pivot = ro_pivot.fillna(0)

# Merge with parts
part_ops = apnrn.merge(ro_pivot, on='routeno', how='left')

# Save CSV
output_csv = r"C:\Users\maaz.mansoor\OneDrive - AlpHa Measurement Solutions\Desktop\SimillarRouteApp\PrecomputedOps.csv"
part_ops.to_csv(output_csv, index=False)
print(f"Precomputed CSV saved at: {output_csv}")
