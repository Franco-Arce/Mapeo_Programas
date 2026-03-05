import pandas as pd

def debug_emails():
    xlsx_file = 'Admitidos_Consolidado.xlsx'
    df_admin = pd.read_excel(xlsx_file)
    admin_emails = df_admin['CORREO'].dropna().astype(str).str.strip().str.lower().tolist()
    
    check_list = [
        "udi.fernandez@gmail.com", "alejaojeda@hotmail.com", "juanchonet20@gmail.com",
        "yubedi.21@gmail.com", "delacruztatiana371@gmail.com", "dinayanidpenaruiz@gmail.com",
        "mayeramos15@yahoo.com", "elisacar2014@gmail.com", "vdarlys10@gmail.com"
    ]
    
    print("Results for some of the listed emails:")
    for email in check_list:
        found = email.lower() in admin_emails
        print(f"{email}: {'FOUND' if found else 'NOT FOUND'}")
        if not found:
            # Check for partial matches
            partial = [ae for ae in admin_emails if email.lower() in ae or ae in email.lower()]
            if partial:
                print(f"  -> Partial matches in Excel: {partial}")

    print("\nSummary of Excel Emails (First 10):")
    print(df_admin['CORREO'].head(10).tolist())

if __name__ == "__main__":
    debug_emails()
