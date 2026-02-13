# Automation-excel-
combine multiple excel into one excel
import pandas as pd
from datetime import datetime

# ✅ Step 1: Load your source Excel file
df = pd.read_excel(r'C:\Users\HarithZainal\OneDrive - iExpertZ [weIzExpertz]\Documents\PROD_ACE_talend_job_io_components_20250723_115809.xlsx')  # adjust path if needed
# ✅ Step 2: Make sure text columns are lowercase (for consistency)
df['Component Type'] = df['Component Type'].str.lower()
# ✅ Step 3: Group source tables
df_input = (
    df[df['Component Type'] == 'input']
    .groupby('Job Name')['Table Name']
    .apply(lambda x: ', '.join(x.dropna().astype(str).unique()))
    .reset_index(name='Source Table(s)')
)
# ✅ Step 4: Group target tables
df_output = (
    df[df['Component Type'] == 'output']
    .groupby('Job Name')['Table Name']
    .apply(lambda x: ', '.join(x.dropna().astype(str).unique()))
    .reset_index(name='Target Table(s)')
)
# ✅ Step 5: Get Job Path (pick first path per job)
df_path = (
    df[['Job Name', 'Job Path']]
    .dropna()
    .drop_duplicates('Job Name')
    .reset_index(drop=True)
)
# ✅ Step 6: Merge all together
df_final = pd.merge(df_input, df_output, on='Job Name', how='outer')
df_final = pd.merge(df_final, df_path, on='Job Name', how='left')
# ✅ Step 7: Save to Excel
filename = datetime.now().strftime("job_summary_%Y%m%d_%H%M%S.xlsx")
output_path = fr'C:\Users\HarithZainal\Downloads\{filename}'
df_final.to_excel(output_path, index=False)
print(f"✅ Success! File saved to:\n{output_path}")
<img width="1333" height="938" alt="image" src="https://github.com/user-attachments/assets/7d1008d9-d170-4e05-ab19-acfd29d4f702" />
