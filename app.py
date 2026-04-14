from flask import Flask, render_template, request, send_file, redirect, url_for, after_this_request
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = '/tmp' if os.environ.get('RENDER') else 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# Create uploads directory if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def is_aims_file(filename):
    return True  # Keine Einschränkung mehr für AIMS Dateien

def is_tajneed_file(filename):
    return True  # Keine Einschränkung mehr für Khuddam Software Dateien

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'aims_file' not in request.files or 'tajneed_file' not in request.files:
        return redirect(url_for('index', error='Beide Dateien müssen hochgeladen werden'))
    
    aims_file = request.files['aims_file']
    tajneed_file = request.files['tajneed_file']
    
    if aims_file.filename == '' or tajneed_file.filename == '':
        return redirect(url_for('index', error='Beide Dateien müssen hochgeladen werden'))
    
    if not (allowed_file(aims_file.filename) and allowed_file(tajneed_file.filename)):
        return redirect(url_for('index', error='Nur Excel-Dateien (.xlsx, .xls) sind erlaubt'))
    
    # Check file naming convention
    if not (is_aims_file(aims_file.filename) and is_tajneed_file(tajneed_file.filename)):
        return redirect(url_for('index', error='Dateinamen müssen mit AIMS_Tajneed_ und Tajnied_ beginnen'))
    
    # Save files
    aims_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(aims_file.filename))
    tajneed_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(tajneed_file.filename))
    
    aims_file.save(aims_path)
    tajneed_file.save(tajneed_path)
    
    # Compare files
    try:
        result_file = compare_excel_files(aims_path, tajneed_path)

        # Delete uploaded files immediately after comparison
        try:
            os.remove(aims_path)
            os.remove(tajneed_path)
        except Exception:
            pass

        @after_this_request
        def cleanup(response):
            try:
                os.remove(result_file)
            except Exception:
                pass
            return response

        return send_file(result_file, as_attachment=True, download_name='Unterschiede_AIMS_Tajneed.xlsx')
    except Exception as e:
        try:
            if os.path.exists(aims_path):
                os.remove(aims_path)
            if os.path.exists(tajneed_path):
                os.remove(tajneed_path)
        except Exception:
            pass
        return redirect(url_for('index', error=str(e)))

def compare_excel_files(aims_path, tajneed_path):
    # Read AIMS Excel file with all sheets
    aims_xl = pd.ExcelFile(aims_path)
    aims_dfs = []
    
    for sheet_name in aims_xl.sheet_names:
        df = pd.read_excel(aims_path, sheet_name=sheet_name)
        df['Majlis'] = sheet_name  # Add Majlis column based on sheet name
        aims_dfs.append(df)
    
    # Combine all AIMS sheets
    aims_df = pd.concat(aims_dfs, ignore_index=True)
    
    # Read Tajneed Excel file
    tajneed_df = pd.read_excel(tajneed_path)
    
    # Find column names
    aims_id_col = None
    tajneed_id_col = None
    
    for col in aims_df.columns:
        if 'MBRMBRCDE' in str(col).upper() or 'JAMAAT ID' in str(col).upper():
            aims_id_col = col
            break
    
    for col in tajneed_df.columns:
        if 'JAMAAT ID' in str(col).upper() or 'MBRMBRCDE' in str(col).upper():
            tajneed_id_col = col
            break
    
    if aims_id_col is None:
        raise ValueError('MBRMBRCDE Spalte nicht in AIMS-Datei gefunden')
    if tajneed_id_col is None:
        raise ValueError('Jamaat ID Spalte nicht in Tajneed-Datei gefunden')
    
    # Extract IDs
    aims_ids = set(aims_df[aims_id_col].dropna().astype(int).astype(str))
    tajneed_ids = set(tajneed_df[tajneed_id_col].dropna().astype(int).astype(str))
    
    # Find missing IDs (in AIMS but not in Tajneed)
    missing_ids = aims_ids - tajneed_ids
    
    # Filter rows with missing IDs
    missing_rows = aims_df[aims_df[aims_id_col].astype(str).isin(missing_ids)]
    
    # Save result - split by Majlis (original AIMS sheet names) into sheets
    result_path = os.path.join(app.config['UPLOAD_FOLDER'], 'Unterschiede_AIMS_Tajneed.xlsx')

    def clean_sheet_name(name):
        # Remove invalid Excel characters: \ / * ? : [ ]
        invalid = r'\/*?:[]'
        for ch in invalid:
            name = name.replace(ch, '')
        return name.strip()[:31] or 'Sheet'

    with pd.ExcelWriter(result_path, engine='openpyxl') as writer:
        # First sheet: all missing rows combined
        missing_rows.to_excel(writer, sheet_name='Alle', index=False)

        # Then one sheet per Majlis
        if 'Majlis' in missing_rows.columns:
            for majlis in missing_rows['Majlis'].dropna().unique():
                group = missing_rows[missing_rows['Majlis'] == majlis]
                if len(group) > 0:
                    sheet_name = clean_sheet_name(str(majlis))
                    group.to_excel(writer, sheet_name=sheet_name, index=False)

    return result_path

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    app.run(host='0.0.0.0', port=port, debug=False)
