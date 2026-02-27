# Skat Positivliste – ISIN Opslag

## Filer
```
app.py            ← Flask backend (proxyer Excel-filen fra Skat.dk)
index.html        ← Frontend (HTML/CSS/JS)
requirements.txt  ← Python-afhængigheder
```

## Opsætning

### 1. Installer afhængigheder
```bash
pip install -r requirements.txt
```

### 2. Start serveren
```bash
python app.py
```

### 3. Åbn i browser
```
http://localhost:5000
```

## Hvordan det virker

- `app.py` henter Excel-filen fra Skat.dk server-side (ingen CORS-problemer) og cacher den i hukommelsen.
- `/api/search?isin=DK0060336014` returnerer JSON med alle matchende rækker.
- `index.html` kalder denne API og viser resultatet.

## Opdater data

Vil du tvinge en ny download af Excel-filen (f.eks. efter Skat udgiver en ny liste):
```bash
curl -X POST http://localhost:5000/api/reload
```

Eller bare genstart serveren.
