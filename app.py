import streamlit as st
import pandas as pd
import datetime
import random
import xlwt
import json
from io import BytesIO
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# --- INTERFACCIA UTENTE (APP) ---
st.set_page_config(page_title="Generatore SIA", page_icon="📅")
st.title("Generatore Orari SIA (.xls)")

st.markdown("Seleziona il periodo desiderato, modifica i valori se necessario e premi Genera.")

# --- SCELTA DATE CON CALENDARIO ---
oggi = datetime.date.today()
default_start = oggi - datetime.timedelta(days=oggi.weekday())

col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Data di inizio", value=default_start)
with col2:
    end_date = st.date_input("Data di fine", value=oggi)

# --- CAMPI DI INPUT ---
with st.expander("Dati Utente e Commessa", expanded=False):
    cod_addetto = st.text_input("Codice Addetto", value="115")
    cod_commessa = st.text_input("Codice Commessa", value="TRASVCON01")

with st.expander("Attività Principale (Giornata intera o 1° metà)", expanded=False):
    descrizione_1 = st.text_input("Descrizione 1", value="Creazione nuovo software - Test Funzionali")
    cod_attivita_1 = st.text_input("Codice 1", value="200923")

with st.expander("Attività Secondaria (2° metà per giornate spezzate)", expanded=False):
    descrizione_2 = st.text_input("Descrizione 2", value="Supporto ai Clienti/Altri Settori - Altri Tipi di Supporto")
    cod_attivita_2 = st.text_input("Codice 2", value="201078")

# OPZIONE DRIVE
salva_su_drive = st.checkbox("☁️ Salva direttamente su Google Drive", value=True)

# --- LOGICA DI GENERAZIONE ---
if st.button("Genera File Excel", type="primary"):
    
    if start_date > end_date:
        st.error("Errore: La data di inizio non può essere successiva alla data di fine!")
    else:
        valori_default = {
            "Cod_Addetto": int(cod_addetto),
            "Cod_Commessa": cod_commessa,
            "Cod_Tipologia": "",
            "Cod_SubTipologia": "",
            "Chiamata": "",
            "Issue": ""
        }

        # --- LOGICA GIORNI SPEZZATI ---
        giorni_spezzati = []
        settimane = {}
        
        temp_date = start_date
        while temp_date <= end_date:
            if temp_date.weekday() < 5: 
                anno_iso, sett_iso, _ = temp_date.isocalendar()
                chiave_settimana = (anno_iso, sett_iso)
                
                if chiave_settimana not in settimane:
                    settimane[chiave_settimana] = []
                settimane[chiave_settimana].append(temp_date)
            temp_date += datetime.timedelta(days=1)

        for giorni_della_settimana in settimane.values():
            num_da_spezzare = min(2, len(giorni_della_settimana))
            giorni_spezzati.extend(random.sample(giorni_della_settimana, num_da_spezzare))

        # --- GENERAZIONE DEI RECORD ---
        records = []
        current_date = start_date

        while current_date <= end_date:
            if current_date.weekday() < 5:  
                if current_date in giorni_spezzati:
                    if current_date.weekday() == 2:
                        durata_1, durata_2 = "04:00", "04:00"
                    else:
                        durata_1, durata_2 = "04:00", "03:30"

                    record1 = {**valori_default, "Data": current_date.strftime("%d/%m/%Y"), "Durata": durata_1, "Cod_Attività": int(cod_attivita_1), "Descrizione": descrizione_1}
                    record2 = {**valori_default, "Data": current_date.strftime("%d/%m/%Y"), "Durata": durata_2, "Cod_Attività": int(cod_attivita_2), "Descrizione": descrizione_2}
                    records.extend([record1, record2])
                else:
                    durata = "08:00" if current_date.weekday() == 2 else "07:30"
                    record = {**valori_default, "Data": current_date.strftime("%d/%m/%Y"), "Durata": durata, "Cod_Attività": int(cod_attivita_1), "Descrizione": descrizione_1}
                    records.append(record)
            current_date += datetime.timedelta(days=1)

        df = pd.DataFrame(records)
        colonne_ordinate = [
            "Cod_Addetto", "Cod_Commessa", "Cod_Attività", "Data", "Durata", 
            "Cod_Tipologia", "Cod_SubTipologia", "Descrizione", "Chiamata", "Issue"
        ]
        df = df[colonne_ordinate]

        # --- CREAZIONE DEL FILE IN MEMORIA ---
        output = BytesIO()
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('Sheet1')
        
        for col_idx, column_name in enumerate(df.columns):
            sheet.write(0, col_idx, column_name)
        
        for row_idx, row in enumerate(df.values):
            for col_idx, value in enumerate(row):
                 if pd.isna(value):
                     sheet.write(row_idx + 1, col_idx, "")
                 else:
                     sheet.write(row_idx + 1, col_idx, value)
        
        workbook.save(output)
        xls_data = output.getvalue()

        str_inizio = start_date.strftime("%Y%m%d")
        str_fine = end_date.strftime("%Y%m%d")
        nome_file = f"SIA_ATTIVITA_{str_inizio}_al_{str_fine}.xls"

        st.success(f"File generato con successo! ({len(df)} righe totali create)")
        
        st.download_button(
            label="📥 Scarica File Excel (.xls) sul telefono",
            data=xls_data,
            file_name=nome_file,
            mime="application/vnd.ms-excel"
        )

        # --- CARICAMENTO SU GOOGLE DRIVE ---
        if salva_su_drive:
            try:
                # 1. Recupera le credenziali dai segreti di Streamlit
                creds_dict = json.loads(st.secrets["GCP_CREDENTIALS"])
                scopes = ['https://www.googleapis.com/auth/drive.file']
                creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
                service = build('drive', 'v3', credentials=creds)

                # 2. Riporta il cursore del file in memoria all'inizio prima di caricarlo
                output.seek(0)

                # 3. Imposta i metadata (Metti qui il tuo FOLDER_ID!!!)
                file_metadata = {
                    'name': nome_file,
                    'parents': ['1rMuHo27iEoB1P7eab0MTLtARpzlOxrwP']  # <--- SOSTITUISCI QUESTO!
                }
                
                # 4. Invia a Google Drive
                media = MediaIoBaseUpload(output, mimetype='application/vnd.ms-excel', resumable=True)
                file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                
                st.success("✅ File caricato con successo sul tuo Google Drive!")
            except Exception as e:
                st.error(f"Errore durante il caricamento su Google Drive: {e}")
