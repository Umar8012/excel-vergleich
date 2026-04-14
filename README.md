# Excel Vergleich - AIMS Tajneed

Web-Anwendung zum Vergleich von AIMS_Tajneed und Khuddam Software Excel-Dateien.

## Funktionen

- Upload von zwei Excel-Dateien (AIMS_Tajneed und Khuddam Software)
- Automatischer Vergleich der IDs über mehrere Sheets hinweg
- Filterung von Einträgen mit MBRTZMCDE = 'B'
- Download der Ergebnis-Excel-Datei mit fehlenden IDs

## Lokal ausführen

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python3 app.py
```

Die App läuft dann auf http://localhost:5000

## Auf Render.com deployen

1. Gehe zu [render.com](https://render.com) und erstelle einen Account
2. Klicke auf "New +" → "Web Service"
3. Verbinde dein GitHub-Repository
4. Konfiguriere:
   - **Name**: excel-vergleich (oder beliebig)
   - **Region**: Frankfurt (oder Europa)
   - **Branch**: main
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
5. Klicke auf "Create Web Service"

Das Deployment dauert ca. 2-3 Minuten. Danach erhältst du eine URL wie `https://excel-vergleich.onrender.com`

## Wichtige Dateien

- `app.py` - Flask-Backend mit Vergleichslogik
- `templates/index.html` - Frontend UI
- `requirements.txt` - Python-Abhängigkeiten
- `runtime.txt` - Python-Version für Render
