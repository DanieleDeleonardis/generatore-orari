import streamlit as st
import pandas as pd
import datetime
import random
import xlwt
from io import BytesIO

# --- INTERFACCIA UTENTE (APP) ---
st.set_page_config(page_title="Generatore SIA", page_icon="📅")
st.title("Generatore Orari SIA (.xls)")

st.markdown("Seleziona il periodo desiderato, modifica i valori se necessario e premi Genera.")

# --- SCELTA DATE CON CALENDARIO ---
oggi = datetime.date.today()
# Di default propone la settimana corrente (dal lunedì a oggi)
default_start = oggi - datetime.timedelta(days=oggi.weekday())

col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Data di inizio", value=default_start)
with col2:
    end_date = st.date_input("Data di fine", value=oggi)

# --- CAMPI DI INPUT CON PLACEHOLDER ---
# Ho impostato expanded=False per tenere chiuse queste tendine di default e rendere l'app più pulita
with st.expander("Dati Utente e Commessa", expanded=False):
    cod_addetto = st.text_input("Codice Addetto", value="115")
    cod_commessa = st.text_input("Codice Commessa", value="TRASVCON01")

with st.expander("Attività Principale (Giornata intera o 1° metà)", expanded=False):
    descrizione_1 = st.text_input("Descrizione 1", value="Creazione nuovo software - Test Funzionali")
    cod_attivita_1 = st.text_input("Codice 1", value="200923")

with st.expander("Attività Secondaria (2° metà per giornate spezzate)", expanded=False):
    descrizione_2 = st.text_input("Descrizione 2", value="Supporto ai Clienti/Altri Settori - Altri Tipi di Supporto")
    cod_attivita_2 = st.text_input("Codice 2", value="201078")

# --- LOGICA DI GENERAZIONE ---
if st.button("Genera File Excel", type="primary"):
    
    # Controllo di sicurezza sulle date
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

        # --- LOGICA GIORNI SPEZZATI (PER SETTIMANA) ---
        giorni_spezzati = []
        settimane = {}
        
        # Raggruppa i giorni candidati (Lun, Mar, Gio, Ven) per settimana
        temp_date = start_date
        while temp_date <= end_date:
            if temp_date.weekday() in [0, 1, 3, 4]:
                # Ottiene l'anno e il numero della settimana
                anno_iso, sett_iso, _ = temp_date.isocalendar()
                chiave_settimana = (anno_iso, sett_iso)
                
                if chiave_settimana not in settimane:
                    settimane[chiave_settimana] = []
                settimane[chiave_settimana].append(temp_date)
            temp_date += datetime.timedelta(days=1)

        # Per OGNI settimana trovata nel range, pesca 2 giorni a caso da spezzare
        for giorni_della_settimana in settimane.values():
            num_da_spezzare = min(2, len(giorni_della_settimana))
            giorni_spezzati.extend(random.sample(giorni_della_settimana, num_da_spezzare))

        # --- GENERAZIONE DEI RECORD ---
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

        # --- CREAZIONE DEL FILE .XLS IN MEMORIA CON XLWT ---
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
            label="📥 Scarica File Excel (.xls)",
            data=xls_data,
            file_name=nome_file,
            mime="application/vnd.ms-excel"
        )
