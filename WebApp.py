import streamlit as st
import pandas as pd
import json

# Sett opp sidetittel og ikon i nettleser-fanen
st.set_page_config(page_title="Mattilsynet JSON-Konverterer", page_icon="🐟")

st.title("🐟 Konverter fra JSON til Excel for Rensefiskrapportering fra Mattilsynet")
st.write("Dra og slipp JSON-filen din under for å flate ut dataene til et Excel-ark.")

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