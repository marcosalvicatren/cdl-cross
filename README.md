# Prima Nota Paghe

App web per convertire riepiloghi paghe (PDF) in XML importabile in **GB Software / Wolters Kluwer**.

## Funzionalità

- **Buste Paga**: PDF → voci classificate automaticamente → XLSX → XML
- **F24**: PDF → voci classificate → XLSX → XML  
- **Regole**: motore a regole modificabile da interfaccia, persistente su GitHub

Le voci non coperte da nessuna regola vengono presentate con conto vuoto — 
compilabile a mano direttamente nell'XLSX, senza toccare le regole.

---

## Setup — prima installazione

### 1. Carica il codice su GitHub

```bash
git init
git add .
git commit -m "Prima Nota Paghe"
git branch -M main
git remote add origin https://github.com/TUO_UTENTE/prima-nota-paghe.git
git push -u origin main
```

### 2. Crea un GitHub Token

1. Vai su https://github.com/settings/tokens
2. **Fine-grained tokens** → Generate new token
3. Seleziona il repository `prima-nota-paghe`
4. Permessi: **Contents → Read and Write**
5. Copia il token (inizia con `github_pat_...`)

### 3. Pubblica su Streamlit Cloud

1. Vai su https://share.streamlit.io
2. **New app** → scegli il repo → branch `main` → file `app.py`
3. Prima di cliccare Deploy, vai su **Advanced settings → Secrets**
4. Incolla questo contenuto (con i tuoi valori reali):

```toml
GITHUB_TOKEN = "github_pat_xxxxxxxxxxxx"
GITHUB_REPO  = "tuo-utente/prima-nota-paghe"
GITHUB_BRANCH = "main"
```

5. Clicca **Deploy**

---

## Come funzionano le regole

Ogni regola ha:
- **Contiene**: parole chiave che devono essere presenti nella descrizione PDF (tutte)
- **NON contiene**: parole chiave che NON devono esserci
- **Conto**: codice conto contabile
- **D/A**: Dare o Avere
- **Descrizione XML**: testo che apparirà nel file XML
- **Conto riga 2** (opzionale): per voci che generano due righe (es. METASALUTE, TFR fondo)

Le regole vengono valutate in ordine — vince la prima che corrisponde.

Quando si salva dalla sezione Regole, il file viene scritto direttamente 
nel repository GitHub tramite API. Al prossimo avvio dell'app, tutti 
i collaboratori vedono le regole aggiornate.

---

## Struttura file

```
app.py                          ← applicazione
requirements.txt                ← dipendenze
.streamlit/secrets.toml.example ← template configurazione
regole/                         ← creata automaticamente al primo salvataggio
  buste_paga.json
  f24.json
```
