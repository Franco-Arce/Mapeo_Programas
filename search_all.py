import pandas as pd

def search_all_files():
    files = ['Admitidos_Consolidado.xlsx', 'ADMITIDOS VIRTUAL 2026-1  PREGRADO CIERRE.xlsx']
    emails_to_check = [
        "udi.fernandez@gmail.com", "alejaojeda@hotmail.com", "juanchonet20@gmail.com",
        "yubedi.21@gmail.com", "delacruztatiana371@gmail.com", "dinayanidpenaruiz@gmail.com"
    ]
    
    for file_name in files:
        print(f"\n--- Checking File: {file_name} ---")
        try:
            xls = pd.ExcelFile(file_name)
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet)
                # Search across all columns
                for col in df.columns:
                    mask = df[col].astype(str).str.contains('|'.join(emails_to_check), na=False, case=False)
                    if mask.any():
                        found = df[mask][col].unique()
                        print(f"Sheet: {sheet} | Column: {col} | Matches: {found}")
        except Exception as e:
            print(f"Error reading {file_name}: {e}")

if __name__ == "__main__":
    search_all_files()
