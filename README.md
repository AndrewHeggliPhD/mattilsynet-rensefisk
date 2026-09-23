# Mattilsynet – Rensefisk API til Excel 🐟

Dette prosjektet ble laget for å løse et praktisk problem for et familie-medlem. De trengte data fra Mattilsynets nye side om rensefisk, men APIer tilbyr bare .json filer til nedlastning, og de fleste utenfor IT-industrien er mer vant med Excel. Jeg lagde først et skript som henter rådata om rensefisk automatisk fra Mattilsynets åpne API, flater ut de nøstede JSON-strukturene, og lagrer dataene i en ryddig Excel-fil som er klar for analyse. Etter de som skulle bruke dette programmet ikke er kjent med VS Code laget jeg et skript til for en web applikasjon som er enklere å bruke. Man laster bare ned .json filen man ønsker fra Mattilsynet og dropper det eller laster det opp på websiden. Applikasjonen kan man finne på https://rensefisk-konverterer.streamlit.app

---

## 🚀 Slik kjører du prosjektet

Hvis det er første gang du kjører dette programmet på din maskin, må du installere noen nødvendige tilleggspakker (`pandas` og `openpyxl`). 

Følg disse stegene:

### 1. Åpne terminalen i VS Code
Gå til toppmenyen i VS Code og velg **Terminal** -> **New Terminal**.


### 2. Installer nødvendige pakker
Kopier koden under, lim den inn i terminalen som dukket opp i bunnen av skjermen, og trykk **Enter**:

```bash
pip install -r requirements.txt
```

### 3. Kjør programmet

Når installasjonen er ferdig, kan du kjøre skriptet ved å trykke på Play-knappen øverst til høyre i VS Code, eller ved å skrive følgende i terminalen og trykke **Enter**:
```bash
python script.py
```

## 📂 Filer i prosjektet
script.py - Hovedkoden til programmet.

requirements.txt - Liste over eksterne Python-pakker som kreves.

README.md - Denne veiledningen.