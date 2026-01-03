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
├─ parser.py            # Parse Excel bestanden naar dicts
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
   source venv/bin/activate
   venv\Scripts\activate
   ```

3. Dependencies:

   ```bash
   pip install -r requirements.txt
   ```

4. Run lokaal:

   ```bash
   flask run
   ```

5. Open in browser:
   [http://localhost:5000](http://localhost:5000)

---