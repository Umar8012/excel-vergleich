import pandas as pd

# AIMS Datei analysieren
aims_file = 'AIMS_Tajneed_Region Nasir Bagh.xlsx'
aims_xl = pd.ExcelFile(aims_file)

print("=" * 60)
print("AIMS_Tajneed Analyse")
print("=" * 60)
print(f"\nAnzahl Sheets: {len(aims_xl.sheet_names)}")
print(f"Sheet Namen: {aims_xl.sheet_names}")

for sheet_name in aims_xl.sheet_names:
    print(f"\n--- Sheet: {sheet_name} ---")
    df = pd.read_excel(aims_file, sheet_name=sheet_name)
    print(f"Zeilen: {len(df)}, Spalten: {len(df.columns)}")
    print(f"Spalten: {df.columns.tolist()}")
    print(f"\nErste 3 Zeilen:")
    print(df.head(3))

# Tajnied Datei analysieren
tajneed_file = 'Tajnied_Nasir Bagh (Groß-Gerau)_20260412.xlsx'
tajneed_xl = pd.ExcelFile(tajneed_file)

print("\n" + "=" * 60)
print("Tajnied Analyse")
print("=" * 60)
print(f"\nAnzahl Sheets: {len(tajneed_xl.sheet_names)}")
print(f"Sheet Namen: {tajneed_xl.sheet_names}")

for sheet_name in tajneed_xl.sheet_names:
    print(f"\n--- Sheet: {sheet_name} ---")
    df = pd.read_excel(tajneed_file, sheet_name=sheet_name)
    print(f"Zeilen: {len(df)}, Spalten: {len(df.columns)}")
    print(f"Spalten: {df.columns.tolist()}")
    print(f"\nErste 3 Zeilen:")
    print(df.head(3))
    
    # Wenn es eine Stadt-Spalte gibt, zeige die einzigartigen Städte
    for col in df.columns:
        if 'stadt' in col.lower() or 'city' in col.lower() or 'region' in col.lower() or 'majlis' in col.lower():
            print(f"\nEinzigartige Werte in '{col}':")
            print(df[col].unique())
