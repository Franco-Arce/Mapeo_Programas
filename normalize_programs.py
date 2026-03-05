import pandas as pd
import difflib
import os

def normalize_programs():
    csv_file = 'na_unab.csv'
    
    # Target list provided by the user
    target_programs = [
        "ADMINISTRACIÓN DE EMPRESAS",
        "CONTADURÍA PUBLICA",
        "DERECHO",
        "ESPECIALIZACIÓN EN COMUNICACIÓN DIGITAL Y MEDIOS INTERACTIVOS",
        "ESPECIALIZACIÓN EN DERECHO LABORAL Y SEGURIDAD SOCIAL",
        "ESPECIALIZACIÓN EN EPIDEMIOLOGIA",
        "ESPECIALIZACIÓN EN GESTIÓN ESTRATÉGICA DE MERCADEO",
        "ESPECIALIZACIÓN EN GESTIÓN HUMANA",
        "ESPECIALIZACIÓN EN SEGURIDAD Y SALUD EN EL TRABAJO",
        "ESPECIALIZACIÓN PRODUCCIÓN Y GESTIÓN DE PROYECTOS AUDIOVISUALES",
        "ESPECIALIZACIÓN TECNOLOGÍA EN FARMACOVIGILANCIA Y TECNOVIGILANCIA",
        "INGENIERÍA INDUSTRIAL",
        "LICENCIATURA EN CIENCIAS SOCIALES",
        "LITERATURA",
        "MAESTRÍA EN CIENCIA DE DATOS",
        "MAESTRÍA EN EDUCACIÓN",
        "MAESTRÍA EN GERENCIA DE PROYECTOS",
        "MAESTRÍA EN GERENCIA PÚBLICA E INNOVACIÓN POLÍTICA",
        "MAESTRÍA EN LITERATURA Y ESCRITURAS CREATIVAS",
        "MAESTRÍAS CIENCIAS BIOMEDICAS",
        "NEGOCIOS INTERNACIONALES",
        "SEGURIDAD Y SALUD EN EL TRABAJO",
        "TECNOLOGÍA EN DESAROLLO DE SOFTWARE",
        "TECNOLOGÍA EN GESTIÓN DE NEGOCIOS",
        "TECNOLOGÍA EN GESTIÓN GASTRONÓMICA",
        "TECNOLOGÍA EN MARKETING",
        "TECNOLOGÍA EN REGENCIA DE FARMACIA",
        "TECNOLOGÍA EN SEGURIDAD Y SALUD TRABAJO",
        "TÉCNICO PROFESIONAL EN PERITAJE AMBIENTAL"
    ]
    
    if not os.path.exists(csv_file):
        print(f"Error: {csv_file} not found.")
        return

    try:
        df = pd.read_csv(csv_file, sep=';', encoding='utf-8')
    except UnicodeDecodeError:
        df = pd.read_csv(csv_file, sep=';', encoding='latin-1')
    
    normalization_report = []

    def get_best_match(original_name):
        if pd.isna(original_name) or str(original_name).strip() == "":
            return original_name
        
        original_name_str = str(original_name).strip()
        
        # Exact match check (case-insensitive)
        for target in target_programs:
            if target.lower() == original_name_str.lower():
                return target
        
        # Fuzzy match using difflib
        matches = difflib.get_close_matches(original_name_str.upper(), target_programs, n=1, cutoff=0.6)
        
        if matches:
            best_match = matches[0]
            # Extra manual adjustments for known abbreviations if needed
            # For example: "TEC EN REGENCIA" -> "TECNOLOGÍA EN REGENCIA DE FARMACIA"
            # But difflib usually handles these well if the threshold is low enough.
            normalization_report.append((original_name_str, best_match))
            return best_match
        
        return original_name

    print("Checking and normalizing entries...")
    df['programa_unab'] = df['programa_unab'].apply(get_best_match)
    
    if normalization_report:
        print("\nNormalization Report:")
        print(f"{'Original':<40} | {'Normalized'}")
        print("-" * 85)
        unique_matches = sorted(list(set(normalization_report)))
        for orig, norm in unique_matches:
            print(f"{orig:<40} | {norm}")
    else:
        print("\nNo changes made.")

    df.to_csv(csv_file, sep=';', index=False, encoding='latin-1')
    print(f"\nSaved updated {csv_file}")

if __name__ == "__main__":
    normalize_programs()
