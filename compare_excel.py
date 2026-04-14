import pandas as pd

# Excel-Dateien lesen
aims_df = pd.read_excel('AIMS_Tajneed_Region Nasir Bagh.xlsx')
tajneed_df = pd.read_excel('Tajnied_Nasir Bagh (Groß-Gerau)_20260412.xlsx')

# IDs aus beiden Dateien extrahieren
aims_ids = set(aims_df['MBRMBRCDE'].dropna().astype(int).astype(str))
tajneed_ids = set(tajneed_df['Jamaat ID'].dropna().astype(int).astype(str))

print(f"AIMS_Tajneed hat {len(aims_ids)} IDs")
print(f"Tajnied hat {len(tajneed_ids)} IDs")

# IDs finden, die in AIMS aber nicht in Tajneed sind
missing_ids = aims_ids - tajneed_ids
print(f"\nFehlende IDs in Tajnied: {len(missing_ids)}")

# Zeilen aus AIMS_Tajneed filtern, die die fehlenden IDs enthalten
missing_rows = aims_df[aims_df['MBRMBRCDE'].astype(str).isin(missing_ids)]

print(f"\nGefundene Zeilen: {len(missing_rows)}")

# Neue Excel-Datei mit den Unterschieden erstellen
output_file = 'Unterschiede_AIMS_Tajneed.xlsx'
missing_rows.to_excel(output_file, index=False)
print(f"\nNeue Datei erstellt: {output_file}")
