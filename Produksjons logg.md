# Utfordinger som ble løst under produksjon

## Generell informasjon om produksjons prosessen

Dette prosjektet ble laget for å løse et praktisk problem for et familie-medlem. De trengte data fra Mattilsynets nye [side](https://www.mattilsynet.no/fisk-og-akvakultur/rensefisk/mattilsynet-publiserer-data-om-rensefisk) om [rensefisk](https://akvakultur-offentlig-api.fisk.mattilsynet.io/docs/#/Rensefiskrapportering/get_api_rensefisk_v1_rapporteringer), men APIer tilbyr bare .json filer til nedlastning, og de fleste utenfor IT-industrien er mer vant med Excel. Dermed tok jeg på meg oppgaven å kunne lage en metode for de som ikke er kjent med .json filer å bare laste ned filen fra Mattilsynet, og "drag and drop" filen for å konvertere til Excel.

 Ettersom dette var en oppgave for å løse et praktisk problem for noen, og ikke for å trene Python språket, så har jeg tillatt meg å bruke Gemini til å gi forslag til kode. Jeg har alikevel foretatt noen avgjørelser underveis som viser en forståelse av hvordan et slikt prosjekt blir til, som er uavhengig av programmeringspråk. Her er logg med flere utfordringer som ble løst:

## Spesifikke utfordringer

1. ChatGPT ble forsøkt brukt til å endre JSON-filene til Excel-filer, uten hell. 

Da brukeren (familie-medlemet) prøvde å laste opp JSON-filen i Excel fungerte det ikke, og ChatGPT klarte heller ikke konvertere større JSON-filer. Jeg tenkte det burde være mulig å lage et enkelt program for å utføre dette, beskrev situasjonen til Gemini, og den ga meg denne koden: 

```Python
import json
import pandas as pd

# 1. Last inn JSON-filen
med_filnavn = 'response_1782810866953.json'
with open(med_filnavn, 'r', encoding='utf-8') as f:
    data = json.load(f)

# 2. Flat ut dataene automatisk med pandas
df = pd.json_normalize(
    data,
    record_path=['produksjonsenheter', 'arter'],
    meta=[
        'organisasjonsnavn', 
        'lokalitetsnavn', 
        'lokalitetsnummer', 
        'år', 
        'måned', 
        'rapporteringstidspunkt',
        'tørrforKg'
    ],
    sep='_'
)

# 3. Definer en ryddig rekkefølge på kolonnene
kolonner = [
    'organisasjonsnavn', 'lokalitetsnavn', 'lokalitetsnummer', 'år', 'måned', 
    'merdId', 'artsnavn', 'artskode', 'opprinnelse', 'beholdningVedForrigeMånedsslutt',
    'utsett_antallFlyttet', 'utsett_antallNy', 
    'uttak_antallAvlivetSykdom', 'uttak_antallAvlivetSkader', 'uttak_antallAvlivetAvmagret',
    'uttak_antallAvlivetForeståendeHåndteringAvLaksen', 'uttak_antallAvlivetForeståendeUgunstigLevemiljø',
    'uttak_antallAvlivetSkalIkkeBrukes', 'uttak_antallSelvdød', 'uttak_antallFlyttetUt', 
    'uttak_antallKanIkkeGjøresRedeFor', 'tørrforKg', 'rapporteringstidspunkt'
]

# Sorter kolonnene og opprett Excel-filen
df = df.reindex(columns=kolonner)
df.to_excel('rensefisk_rapport.xlsx', index=False)

print("Suksess! Filen 'rensefisk_rapport.xlsx' er opprettet i mappen din.")
```

Det første problemet her var at koden krevde at JSON-filen man skal konvertere ligger i samme mappe som hovedskriptet. En smidigere løsning for generelle brukere er å la de laste opp en fil fra sin egen utforsker. Dette fortalte jeg til Gemini, og den ga meg denne koden: 

```Python
import json
import os
from tkinter import Tk, filedialog
import pandas as pd


def hent_og_flat_ut_json():
    # 1. Skjul det tomme hovedvinduet til Tkinter (vi vil bare ha fildialogen)
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)  # Gjør at vinduet legger seg foran VS Code

    # 2. Åpne maskinens filutforsker og la brukeren velge en .json-fil
    print("Åpner filutforsker...")
    filsti = filedialog.askopenfilename(
        title="Velg JSON-filen fra Mattilsynet",
        filetypes=[("JSON-filer", "*.json")],
    )

    # Hvis brukeren trykker "Avbryt" i filutforskeren, stopper vi programmet pent
    if not filsti:
        print("Ingen fil ble valgt. Avbryter.")
        return

    # 3. Last inn filen brukeren valgte
    print(f"Laster inn: {os.path.basename(filsti)}")
    with open(filsti, "r", encoding="utf-8") as f:
        data = json.load(f)

    # 4. Flat ut dataene automatisk med pandas
    print("Prosesserer og flater ut data...")
    df = pd.json_normalize(
        data,
        record_path=["produksjonsenheter", "arter"],
        meta=[
            "organisasjonsnavn",
            "lokalitetsnavn",
            "lokalitetsnummer",
            "år",
            "måned",
            "rapporteringstidspunkt",
            "tørrforKg",
        ],
        sep="_",
    )

    # Definer en ryddig rekkefølge på kolonnene
    kolonner = [
        "organisasjonsnavn",
        "lokalitetsnavn",
        "lokalitetsnummer",
        "år",
        "måned",
        "merdId",
        "artsnavn",
        "artskode",
        "opprinnelse",
        "beholdningVedForrigeMånedsslutt",
        "utsett_antallFlyttet",
        "utsett_antallNy",
        "uttak_antallAvlivetSykdom",
        "uttak_antallAvlivetSkader",
        "uttak_antallAvlivetAvmagret",
        "uttak_antallAvlivetForeståendeHåndteringAvLaksen",
        "uttak_antallAvlivetForeståendeUgunstigLevemiljø",
        "uttak_antallAvlivetSkalIkkeBrukes",
        "uttak_antallSelvdød",
        "uttak_antallFlyttetUt",
        "uttak_antallKanIkkeGjøresRedeFor",
        "tørrforKg",
        "rapporteringstidspunkt",
    ]
    df = df.reindex(columns=kolonner)

    # 5. Finn ut hvor brukeren vil lagre den ferdige Excel-filen
    lagringssti = filedialog.asksaveasfilename(
        title="Hvor vil du lagre Excel-filen?",
        defaultextension=".xlsx",
        filetypes=[("Excel-filer", "*.xlsx")],
        initialfile="rensefisk_rapport.xlsx",
    )

    if lagringssti:
        df.to_excel(lagringssti, index=False)
        print(f"Suksess! Excel-filen ble lagret her: {lagringssti}")
    else:
        print("Lagring avbrutt.")


# Kjør funksjonen
if __name__ == "__main__":
    hent_og_flat_ut_json()
```

2. For å kunne kjøre det nye skriptet så måtte jeg installere Python via terminalen, men det er bedre å ha et dokument hvor alle nøvendige pakker kan lastes samtidig i tilfelle andre skal bruke eller prøve å kjøre prosjektet hos seg. 

For å løse dette lagde jeg en "requirements.txt" fil og la til alle pakkene som trengtes etterhvert som jeg jobbet videre. 

3. Ikke alle er kjent med VS Code

I tilfelle noen som ikke var kjent med VS Code skulle prøve å kjøre prosjektet så lagde jeg også en README.md fil som forklarte (ved hjelp av Gemini) hva som trengs å gjøres, og i hvilken rekkefølge for å kjøre prosjektet. 

Etter dette prøvdekjørte jeg koden, og den funket, men den var hardkodet.

4. Hardkodet skript fungerer dårlig hvis det skjer en endring i inputen

Dette fungerer dårlig hvis Mattilsynet skulle endre fremtidig rensefisk filer, så jeg ba Gemini om å endre skriptet til å ikke inkludere hardkoding. Den ga meg da denne koden: 

```Python
# 4. Flat ut dataene automatisk uten å definere kolonnene på forhånd
    print("Prosesserer og flater ut data...")
    df = pd.json_normalize(
        data,
        record_path=['produksjonsenheter', 'arter'],
        meta=[
            'organisasjonsnavn', 
            'lokalitetsnavn', 
            'lokalitetsnummer', 
            'år', 
            'måned', 
            'rapporteringstidspunkt',
            'tørrforKg'
        ],
        sep='_'
    )

    # I stedet for å bruke en fast liste, sorterer vi bare kolonnene alfabetisk 
    # slik at metadata havner samlet, og fjerner tomme kolonner hvis de finnes.
    df = df.reindex(sorted(df.columns), axis=1)
```

"meta" var da fortsatt hardkodet, så jeg ba Gemini endre det, da fikk jeg denne koden: 

```Python
# 4. Fullstendig dynamisk utflating (Ingen hardkodede kolonnenavn!)
    print("Prosesserer og flater ut data dynamisk...")
    
    flate_rader = []

    for rapport in data:
        # Lag en kopi av alt på toppnivået (orgnavn, år, måned, tørrfor, etc.)
        # Vi fjerner 'produksjonsenheter' midlertidig så den ikke blir med som rå-JSON
        toppnivå = {k: v for k, v in rapport.items() if k != 'produksjonsenheter'}
        
        # Sjekk om det i det hele tatt finnes produksjonsenheter (merder)
        if 'produksjonsenheter' in rapport and rapport['produksjonsenheter']:
            for merd in rapport['produksjonsenheter']:
                # Lag en kopi av alt på merd-nivået (f.eks. merdId)
                merd_data = {k: v for k, v in merd.items() if k != 'arter'}
                
                if 'arter' in merd and merd['arter']:
                    for art in merd['arter']:
                        # Bruk pandas sin innebygde json_normalize til å flate ut 
                        # de dypeste strukturene (utsett og uttak) automatisk
                        art_flat = pd.json_normalize(art, sep='_').to_dict(orient='records')[0]
                        
                        # Slå sammen alt til én eneste flat rad: Toppnivå + Merd + Art
                        ny_rad = {**toppnivå, **merd_data, **art_flat}
                        flate_rader.append(ny_rad)
                else:
                    # Hvis en merd mot formodning ikke har noen arter registrert
                    ny_rad = {**toppnivå, **merd_data}
                    flate_rader.append(ny_rad)
        else:
            # Hvis rapporten er helt tom for merder, tar vi bare med toppnivået
            flate_rader.append(toppnivå)

    # Gjør om den flate listen med ordbøker til en DataFrame
    df = pd.DataFrame(flate_rader)

    # Sorter kolonnene alfabetisk så det ser ryddig ut i Excel
    df = df.reindex(sorted(df.columns), axis=1)

if __name__ == "__main__":
    hent_og_flat_ut_json()    
``` 

Det er fortsatt forbedringspotensiale mtp å få koden til å være enda mer robust (altså mindre hardkodet), men jeg sa meg fornøyd for å være effektiv med tiden min. 

5. Brukeren trenger å kunne bruke programmet uten programmeringsverktøy

Etter som brukeren skal kunne bruke dette uten VS Code eller andre programmeringsverktøy, tenkte jeg først at vi kunne lage et program som kan installeres på maskinen til brukeren, men det er ikke mulig å lage *et* installeringsprogram for alle datasystemer samtidig, og i tillegg kan det være begrensninger på programmer man kan installere på jobb. Dermed landet jeg på at programmet kunne gjøres til en web applikasjon. 

6. Programmet må kunne brukes som en web applikasjon, av alle som trenger det

Gemini foreslo å bruke streamlit, jeg la til pakken i "requirements.txt", og lagde en ny .py skript som Gemini sendte: 

```Python
import streamlit as st
import pandas as pd
import json

# Sett opp sidetittel og ikon i nettleser-fanen
st.set_page_config(page_title="Mattilsynet JSON-Konverterer", page_icon="🐟")

st.title("🐟 Mattilsynet – Rensefisk API til Excel")
st.write("Dra og slipp JSON-filen din under for å flate ut dataene til et ryddig Excel-ark.")

# 1. Fil-opplaster på nettsiden
opplastet_fil = st.file_uploader("Velg JSON-fil", type=["json"])

if opplastet_fil is not None:
    try:
        # 2. Last inn JSON-dataen
        data = json.load(opplastet_fil)
        st.success(f"Filen '{opplastet_fil.name}' ble lastet inn!")

        # 3. Fullstendig dynamisk utflating (nøyaktig lik din gamle kode)
        flate_rader = []
        for rapport in data:
            toppnivå = {k: v for k, v in rapport.items() if k != 'produksjonsenheter'}
            
            if 'produksjonsenheter' in rapport and rapport['produksjonsenheter']:
                for merd in rapport['produksjonsenheter']:
                    merd_data = {k: v for k, v in merd.items() if k != 'arter'}
                    
                    if 'arter' in merd and merd['arter']:
                        for art in merd['arter']:
                            art_flat = pd.json_normalize(art, sep='_').to_dict(orient='records')[0]
                            ny_rad = {**toppnivå, **merd_data, **art_flat}
                            flate_rader.append(ny_rad)
                    else:
                        ny_rad = {**toppnivå, **merd_data}
                        flate_rader.append(ny_rad)
            else:
                flate_rader.append(toppnivå)

        df = pd.DataFrame(flate_rader)
        df = df.reindex(sorted(df.columns), axis=1)

        # 4. Vis en liten forhåndsvisning av dataene på nettsiden
        st.write("### Forhåndsvisning av data (første 5 rader):")
        st.dataframe(df.head())

        # 5. Konverter DataFrame til Excel i minnet (Streamlit-måten)
        import io
        towrite = io.BytesIO()
        df.to_excel(towrite, index=False, engine='openpyxl')
        towrite.seek(0)

        # 6. Lag en stor, fin nedlastingsknapp
        st.download_button(
            label="📥 Last ned ferdig Excel-fil",
            data=towrite,
            file_name="rensefisk_rapport.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Det skjedde en feil under prosesseringen: {e}")
```

Jeg testet at appen funket, lagde en repository på GitHub, fulgte instruksjonene til Gemini for å få den på streamlit, og nå fungerer siden og er tilgjengelig på https://rensefisk-konverterer.streamlit.app