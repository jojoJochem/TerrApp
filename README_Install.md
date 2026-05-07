# TerrApp deployment — instructies voor beheerder

Korte handleiding voor de eenmalige inrichting van TerrApp (interne Flask-applicatie) op TER-RDS01.

## Overzicht

De setup bestaat uit twee fases:

1. **Fase 1 (beheerder, eenmalig)** — software installeren en omgeving voorbereiden
2. **Fase 2 (beheerder, eenmalig)** — `setup-terrapp.ps1` script draaien om service en rechten in te richten

Daarna kan de eindgebruiker (TERRASCAN\jochemvangaalen) zelfstandig code-updates en service-restarts uitvoeren zonder admin-rechten.

---

## Fase 1 — Software installeren

Alle vier handmatig, vereist admin:

### 1. Python 3.12.10 (64-bit)

- Download: https://www.python.org/downloads/windows/
- Installer-opties:
  - **Add Python to PATH & Use Admin privileges** aanvinken -> dan customize installation kiezen
  - Standaard features (pip, tcl/tk, documentation) aan laten
  - **Install for all users** aanvinken (komt in `C:\Program Files\Python312\`)

Verificatie in een nieuwe cmd:
```
python --version
pip --version
```

### 2. Git for Windows

- Download: https://git-scm.com/download/win waarschijnlijk de x64-bit versie
- Installer-defaults zijn prima

### 3. NSSM

- Download: https://nssm.cc/download (nssm 2.24-101-g897c7ad)
- Pak de ZIP uit, kopieer `win64\nssm.exe` naar `C:\tools\nssm.exe`
- Map `C:\tools\` aanmaken als die nog niet bestaat

### 4. App-map aanmaken

```powershell
New-Item -ItemType Directory -Path C:\apps\terrapp
```

---

## Tussenfase — eindgebruiker bouwt de venv

Voordat het setup-script gedraaid kan worden, moet de eindgebruiker eerst:

1. De applicatiecode klonen naar `C:\apps\terrapp\` (`git clone <repo> .`)
2. Het virtual environment aanmaken en dependencies installeren:
   ```
   cd C:\apps\terrapp
   python -m venv venv
   venv\Scripts\activate
   pip install -r requirements.txt
   ```

Hiervoor heeft de eindgebruiker tijdelijk schrijfrechten nodig op `C:\apps\terrapp\`. Dit kan via één van twee routes:

- **A.** Beheerder geeft tijdelijk Modify-rechten (worden later door het script formeel ingericht), of
- **B.** Beheerder doet de git clone en venv-aanmaak zelf (commando's hierboven)

Optie B is sneller en houdt rechten-management consistent.

---

## Fase 2 — Setup script draaien

Het script `setup-terrapp.ps1` doet de rest:

- Maakt de logs-map aan
- Geeft `TERRASCAN\jochemvangaalen` Modify-rechten op `C:\apps\terrapp\`
- Installeert de Windows service `TerrApp` via NSSM
- Geeft `TERRASCAN\jochemvangaalen` Start/Stop/Query rechten op de service (via `sc sdset`)
- Opent firewall poort 8000 (alleen voor Domain en Private netwerken)
- Start de service

### Draaien

Open PowerShell **als Administrator** en draai:

```powershell
cd C:\path\where\script\is
.\setup-terrapp.ps1
```

Het script is **idempotent** — als iets al bestaat wordt het overgeslagen, niet overschreven. Veilig om opnieuw te draaien als er iets misging.

### Configuratie

Standaard-instellingen staan bovenin het script onder `# CONFIG`. Pas aan indien nodig (bv. andere poort, andere gebruiker, ander pad).

---

## Verificatie

Na het script:

1. Service draait: `Get-Service TerrApp` → Status `Running`
2. App bereikbaar vanaf andere PC in het netwerk: `http://TER-RDS01:8000`
3. Logs schrijven: `C:\apps\terrapp\logs\stdout.log` groeit

---

## Service-account (optioneel maar aanbevolen)

Standaard draait de service onder `LocalSystem` (volledige systeemrechten). Voor productie wordt aangeraden een dedicated low-privilege account te gebruiken. Twee opties:

**A. Virtual service account (eenvoudigst)**
```powershell
nssm set TerrApp ObjectName "NT SERVICE\TerrApp" ""
```
Geen wachtwoord nodig, account wordt automatisch beheerd door Windows.

**B. Domain service account**
Aanmaken in AD, daarna:
```powershell
nssm set TerrApp ObjectName "TERRASCAN\svc-terrapp" "<password>"
```

In beide gevallen daarna controleren dat het service-account Read & Execute heeft op `C:\apps\terrapp\` en Modify op `C:\apps\terrapp\logs\`.

---

## Wat de eindgebruiker daarna zelf kan

Zonder admin-rechten:

```
cd C:\apps\terrapp
git pull
sc stop TerrApp
sc start TerrApp
type logs\stderr.log
```

Voor wijzigingen in service-config, nieuwe poorten, of Python-versies blijft beheerder nodig.

---

## Vragen / contact

[jouw naam en contact]