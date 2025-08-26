# TerrApp

**Excel → Word Tabel Generator**
Een Flask-app die Excel-bestanden inleest (sheet *"Tabel"*) en automatisch een Word-document (`.docx`) genereert met twee tabellen:

1. **Samenstelling analysemonsters**
2. **Samenvatting toetsing milieuhygiënische kwaliteit grond**

---

## 🚀 Deployments

* **Production**: [https://terrapp-production.herokuapp.com](https://terrapp-production.herokuapp.com)

---

## 📂 Projectstructuur

```
TerrApp/
├─ app.py               # Flask app (routes + upload handling)
├─ exporter.py          # Exporteert samples naar .docx met python-docx
├─ parser.py            # Parse Excel bestanden naar sample dicts (pandas + openpyxl)
├─ requirements.txt     # Python dependencies
├─ Procfile             # Heroku startinstructie
├─ runtime.txt          # Python versie (Heroku)
├─ templates/
│  └─ index.html        # Frontend HTML
└─ static/
   └─ style.css         # Frontend styling
```

---

## ⚙️ Installatie (lokaal)

1. Repo clonen:

   ```bash
   git clone https://github.com/jojoJochem/TerrApp.git
   cd TerrApp
   ```

2. Virtuele omgeving:

   ```bash
   python -m venv venv
   source venv/bin/activate   # Mac/Linux
   venv\Scripts\activate      # Windows
   ```

3. Dependencies:

   ```bash
   pip install -r requirements.txt
   ```

4. Run lokaal:

   ```bash
   flask run
   # of: python app.py
   ```

5. Open in browser:
   [http://localhost:5000](http://localhost:5000)

---

## ☁️ Deployment op Heroku

### Eénmalig

```bash
# login
heroku login

# staging app maken
heroku create terrapp-staging-40b877f15eb6

# production app maken
heroku create terrapp-production
```

### Dyno instellen (basic)

```bash
# staging
heroku ps:type web=basic -a terrapp-staging-40b877f15eb6
heroku ps:scale web=1 -a terrapp-staging-40b877f15eb6

# production
heroku ps:type web=basic -a terrapp-production
heroku ps:scale web=1 -a terrapp-production
```

### Deploy

```bash
git push heroku main
# of specifiek:
git push https://git.heroku.com/terrapp-production.git main
```

---

## 🔄 CI/CD via GitHub Actions

* **Push naar `main`** → automatische deploy naar staging → smoke test → promotie naar production.
* Secrets ingesteld in GitHub repo:

  * `HEROKU_API_KEY`
  * `HEROKU_EMAIL`
  * `HEROKU_STAGING_APP=terrapp-staging-40b877f15eb6`
  * `HEROKU_PROD_APP=terrapp-production`

Workflows staan in `.github/workflows/`.

---

## ⏪ Rollback

Als een nieuwe versie problemen veroorzaakt:

```bash
# toon releases
heroku releases -a terrapp-production

# rollback naar vorige release
heroku releases:rollback -a terrapp-production

# of naar een specifieke release
heroku releases:rollback v123 -a terrapp-production
```

---

## 📜 Licentie

Dit project is privé voor intern gebruik.

