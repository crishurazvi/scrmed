# Generator Streamlit pentru scrisoare medicală cardiologică

Aplicație locală pentru redactarea rapidă a unei scrisori medicale în format asemănător exemplelor analizate. Sugestiile extrase sunt doar ajutor de redactare, nu diagnostic automat și nu recomandare terapeutică.

## Structură proiect

```text
streamlit_scrisoare_medicala/
├── app.py
├── requirements.txt
├── data/
│   └── options.json
└── src/
    ├── anonymizer.py
    ├── letter_utils.py
    └── options_store.py
```

## Instalare locală

```bash
python -m venv .venv
source .venv/bin/activate  # macOS/Linux
# .venv\Scripts\activate  # Windows
pip install -r requirements.txt
streamlit run app.py
```

## Confidențialitate

- Rulează preferabil local pentru date reale de pacienți.
- Nu încărca CNP-uri, nume reale sau documente neanonimizate pe un repo public.
- Dacă publici pe Streamlit Community Cloud, păstrează doar date demo/anonymizate.
- Folosește tabul „Anonimizare” doar ca prim pas. Verificarea manuală rămâne obligatorie.

## Export

- DOCX: export principal, creat cu `python-docx`.
- PDF: export simplu, orientativ, creat cu `reportlab`. Pentru documente oficiale, DOCX rămâne preferabil.
- TXT: util pentru copiere rapidă sau arhivare internă.

## Hosting pe Streamlit Community Cloud

1. Creează un repository GitHub privat sau public fără date reale de pacienți.
2. Pune fișierele proiectului în repository.
3. Pe Streamlit Community Cloud: `New app` → alege repository-ul → `app.py`.
4. Adaugă `requirements.txt` la rădăcina proiectului.
5. Atenție: modificările salvate în `data/options.json` din interfață pot fi temporare pe cloud. Pentru persistență reală folosește Git, Google Sheets, SQLite pe un volum persistent sau o bază de date externă.

## Limitări

- Aplicația nu pune diagnostice automat.
- Aplicația nu recomandă tratamente.
- Medicamentele și diagnosticele din listă provin din exemple și trebuie confirmate clinic.
- Importul direct al fișierelor `.doc` vechi nu este inclus pentru cloud. Pentru extindere, convertește `.doc` în `.docx` înainte de procesare sau rulează local cu unelte dedicate.
