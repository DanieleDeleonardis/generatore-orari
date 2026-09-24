import streamlit as st
import pandas as pd
import datetime
import random
from io import BytesIO

# --- INTERFACCIA UTENTE (APP) ---
st.set_page_config(page_title="Generatore SIA", page_icon="📅")
st.title("Generatore Orari SIA")

st.markdown("Modifica i valori qui sotto se necessario, oppure lascia quelli di default e premi Genera.")

# --- CAMPI DI INPUT CON PLACEHOLDER ---
with st.expander("Dati Utente e Commessa", expanded=True):
    cod_addetto = st.text_input("Codice Addetto", value="115")
    cod_commessa = st.text_input("Codice Commessa", value="TRASVCON01")

with st.expander("Attività Principale (Giornata intera o 1° metà)"):
    descrizione_1 = st.text_input("Descrizione", value="Creazione nuovo software - Test Funzionali")
    cod_attivita_1 = st.text_input("Codice", value="200923")

with st.expander("Attività Secondaria (2° metà per giornate spezzate)"):
    descrizione_2 = st.text_input("Descrizione", value="Supporto ai Clienti/Altri Settori - Altri Tipi di Supporto")
    cod_attivita_2 = st.text_input("Codice", value="201078")

# --- LOGICA DI GENERAZIONE ---
if st.button("Genera File Excel", type="primary"):
    
    valori_default = {
        "Cod_Addetto": int(cod_addetto),
        "Cod_Commessa": cod_commessa,
        "Cod_Tipologia": "",
        "Cod_SubTipologia": "",
        "Chiamata": "",
        "Issue": ""
    }

    oggi = datetime.date.today()
    start_date = oggi - datetime.timedelta(days=oggi.weekday())
    end_date = oggi

    giorni_candidati = []
    temp_date = start_date
    while temp_date <= end_date:
        if temp_date.weekday() in [0, 1, 3, 4]:
            giorni_candidati.append(temp_date)
        temp_date += datetime.timedelta(days=1)

    num_da_spezzare = min(2, len(giorni_candidati))
    giorni_spezzati = random.sample(giorni_candidati, num_da_spezzare)

    records = []
    current_date = start_date

    while current_date <= end_date:
        if current_date.weekday() < 5:  
            if current_date in giorni_spezzati:
                record1 = {**valori_default, "Data": current_date.strftime("%d/%m/%Y"), "Durata": "04:00", "Cod_Attività": int(cod_attivita_1), "Descrizione": descrizione_1}
                record2 = {**valori_default, "Data": current_date.strftime("%d/%m/%Y"), "Durata": "04:00", "Cod_Attività": int(cod_attivita_2), "Descrizione": descrizione_2}
                records.extend([record1, record2])
            else:
                durata = "07:30" if current_date.weekday() == 2 else "08:00"
                record = {**valori_default, "Data": current_date.strftime("%d/%m/%Y"), "Durata": durata, "Cod_Attività": int(cod_attivita_1), "Descrizione": descrizione_1}
                records.append(record)
        current_date += datetime.timedelta(days=1)

    df = pd.DataFrame(records)
    colonne_ordinate = [
        "Cod_Addetto", "Cod_Commessa", "Cod_Attività", "Data", "Durata", 
        "Cod_Tipologia", "Cod_SubTipologia", "Descrizione", "Chiamata", "Issue"
    ]
    df = df[colonne_ordinate]

    # Prepara il file per il download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Orari')
    excel_data = output.getvalue()

    str_inizio = start_date.strftime("%Y%m%d")
    str_fine = end_date.strftime("%Y%m%d")
    nome_file = f"SIA_ATTIVITA_{str_inizio}_al_{str_fine}.xlsx"

    st.success("File generato con successo!")
    
    # Pulsante per scaricare il file sul telefono o PC
    st.download_button(
        label="📥 Scarica File Excel",
        data=excel_data,
        file_name=nome_file,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )