import pandas as pd
import os

def update_programs():
    csv_file = 'na_unab.csv'
    xlsx_file = 'Admitidos_Consolidado.xlsx'
    
    print(f"Reading {csv_file}...")
    # Read the CSV with semicolon delimiter as seen in the file preview
    df_na = pd.read_csv(csv_file, sep=';')
    
    print(f"Reading {xlsx_file}...")
    # Read the Excel file
    df_admin = pd.read_excel(xlsx_file)
    
    # Pre-process emails for better matching (lowercase and strip)
    df_na['emlmail'] = df_na['emlmail'].astype(str).str.strip().str.lower()
    df_admin['CORREO'] = df_admin['CORREO'].astype(str).str.strip().str.lower()
    
    # Create a mapping from email to program description
    # Taking the first match if there are duplicates
    mapping = df_admin.dropna(subset=['CORREO', 'DESC_PROGRAMA']).drop_duplicates('CORREO').set_index('CORREO')['DESC_PROGRAMA'].to_dict()
    
    print(f"Matching data...")
    # Update the 'programa_unab' column (column E)
    # Map the email to the program description
    def get_program(email):
        # Clean current email if it's like "email1 - email2" or has other issues
        # But for now we use the exact match after cleaning as per the mapping
        return mapping.get(email, None)
    
    # We only update if 'programa_unab' is empty or #N/A (though the CSV showed it as empty)
    # The user said to "agregar", which usually means fill if empty or replace.
    # I'll fill it with the mapping.
    
    # Apply the mapping
    df_na['programa_unab'] = df_na['emlmail'].map(mapping).fillna(df_na['programa_unab'])
    
    print(f"Saving updated {csv_file}...")
    # Save back to CSV keeping the semicolon delimiter
    df_na.to_csv(csv_file, sep=';', index=False)
    print("Done!")

if __name__ == "__main__":
    update_programs()
