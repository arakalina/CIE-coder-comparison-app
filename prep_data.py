import pandas as pd
import numpy as np
import os

def load_and_melt(filepath):
    if filepath.endswith('.xlsb'):
        df = pd.read_excel(filepath, engine='pyxlsb')
    else:
        df = pd.read_excel(filepath)
    
    # Replace 'undefined' text with NaN
    df = df.replace(r'(?i)^\s*undefined\s*$', np.nan, regex=True)
    
    # Melt into long format
    df_long = df.melt(id_vars=['Document name'], var_name='Category', value_name='Score')
    return df_long

def prepare_log():
    df_A = load_and_melt("Alina_All Scenario_Codes_simple.xlsb")
    df_S = load_and_melt("S_Scores.xlsx")

    df_A = df_A.rename(columns={'Score': 'Score_A'})
    df_S = df_S.rename(columns={'Score': 'Score_S'})

    # Keep only documents that S was assigned to score
    merged = pd.merge(df_A, df_S, on=['Document name', 'Category'], how='inner')

    discrepancies = []

    for index, row in merged.iterrows():
        a = row['Score_A']
        s = row['Score_S']
        
        if pd.isna(a): a = np.nan
        if pd.isna(s): s = np.nan

        # Skip matching scores or both missing
        if a == s or (pd.isna(a) and pd.isna(s)):
            continue
            
        # Ignore 0.5 vs 1
        if (a == 0.5 and s == 1) or (a == 1 and s == 0.5):
            continue

        # Flag Severity
        severity = None
        if (a == 1 and s == 0) or (a == 0 and s == 1):
            severity = "Red"
        elif (a == 0.5 and s == 0) or (a == 0 and s == 0.5):
            severity = "Orange"
        elif (pd.isna(a) and not pd.isna(s)) or (not pd.isna(a) and pd.isna(s)):
            severity = "Missing"

        # Determine Initial Workflow Status:
        # If A has missing data, send to A first ('Pending A').
        # Otherwise, send to S first ('Pending S').
        initial_status = 'Pending A' if pd.isna(a) else 'Pending S'

        if severity:
            discrepancies.append({
                'ID': f"{row['Document name']}_{row['Category']}",
                'Document': row['Document name'],
                'Category': row['Category'],
                'Score_A': "Missing" if pd.isna(a) else a,
                'Score_S': "Missing" if pd.isna(s) else s,
                'Severity': severity,
                'Status': initial_status,
                'Revision_Round': 1,
                'Comment_S': "",
                'Comment_A': "",
                'Final_Score': ""
            })

    new_log = pd.DataFrame(discrepancies)

    if os.path.exists("Discrepancy_Log.csv"):
        old_log = pd.read_csv("Discrepancy_Log.csv", dtype=str)
        new_log = new_log[~new_log['ID'].isin(old_log['ID'])]
        final_log = pd.concat([old_log, new_log], ignore_index=True)
    else:
        final_log = new_log

    final_log.to_csv("Discrepancy_Log.csv", index=False)
    print(f"Data Prep Complete! Log saved with {len(final_log)} discrepancies.")

if __name__ == "__main__":
    prepare_log()