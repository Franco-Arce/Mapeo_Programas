import pandas as pd
import os

def enrich_from_bases():
    csv_file = 'na_unab.csv'
    xls_file = 'TodaslasBases.xls'
    
    if not os.path.exists(xls_file):
        print(f"Error: {xls_file} not found.")
        return
    
    print(f"Reading {csv_file}...")
    try:
        df_na = pd.read_csv(csv_file, sep=';', encoding='utf-8')
    except UnicodeDecodeError:
        df_na = pd.read_csv(csv_file, sep=';', encoding='latin-1')
    
    print(f"Reading {xls_file} (this might take a while due to size)...")
    # Reading necessary columns only to save memory
    df_bases = pd.read_excel(xls_file, usecols=['EMLMAIL', 'Programa interes'])
    
    # Pre-process for matching
    df_na['emlmail_clean'] = df_na['emlmail'].astype(str).str.strip().str.lower()
    df_bases['EMLMAIL'] = df_bases['EMLMAIL'].astype(str).str.strip().str.lower()
    
    # Create mapping
    # Drop duplicates to avoid issues, keep last (more recent potential entry)
    mapping = df_bases.dropna(subset=['EMLMAIL', 'Programa interes']).drop_duplicates('EMLMAIL', keep='last').set_index('EMLMAIL')['Programa interes'].to_dict()
    
    print(f"Enriching empty fields...")
    rows_updated = 0
    
    def fill_program(row):
        nonlocal rows_updated
        if pd.isna(row['programa_unab']) or str(row['programa_unab']).strip() == "":
            email = row['emlmail_clean']
            new_prog = mapping.get(email)
            if new_prog:
                rows_updated += 1
                return new_prog
        return row['programa_unab']
    
    df_na['programa_unab'] = df_na.apply(fill_program, axis=1)
    
    # Cleanup temporary column
    df_na = df_na.drop(columns=['emlmail_clean'])
    
    print(f"Updated {rows_updated} records.")
    
    print(f"Saving {csv_file}...")
    df_na.to_csv(csv_file, sep=';', index=False, encoding='latin-1')
    print("Enrichment complete.")

if __name__ == "__main__":
    enrich_from_bases()
