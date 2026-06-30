# Last inn nødvendige pakker
import pandas as pd
import json
import os
from tkinter import Tk, filedialog


# Last inn JSON-filen
def hent_og_flat_ut_json():
    # 1. Skjul det tomme hovedvinduet til Tkinter (vi vil bare ha fildialogen)
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)  # Gjør at vinduet legger seg foran VS Code

    # 2. Åpne maskinens filutforsker og la brukeren velge en .json-fil
    print("Åpner filutforsker...")
    filsti = filedialog.askopenfilename(
        title="Velg JSON-filen fra Mattilsynet",
        filetypes=[("JSON-filer", "*.json")]
    )

    # Hvis brukeren trykker "Avbryt" i filutforskeren, stoppes programmet
    if not filsti:
        print("Ingen fil ble valgt. Avbryter.")
        return

    # 3. Last inn filen brukeren valgte
    print(f"Laster inn: {os.path.basename(filsti)}")
    with open(filsti, 'r', encoding='utf-8') as f:
        data = json.load(f)

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

    # 5. Finn ut hvor brukeren vil lagre den ferdige Excel-filen
    lagringssti = filedialog.asksaveasfilename(
        title="Hvor vil du lagre Excel-filen?",
        defaultextension=".xlsx",
        filetypes=[("Excel-filer", "*.xlsx")],
        initialfile="rensefisk_rapport.xlsx"
    )

    if lagringssti:
        df.to_excel(lagringssti, index=False)
        print(f"Suksess! Excel-filen ble lagret her: {lagringssti}")
    else:
        print("Lagring avbrutt.")


if __name__ == '__main__':
    hent_og_flat_ut_json()