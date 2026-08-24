# =====================================================================================
# BLOCCO 1: STRUTTURA DI BASE, ICONA IMMAGINE REALE E RIMOZIONE MENU DI SISTEMA
# IMPOSTA IL LOGO PROPRIETARIO PNG COME ICONA SCHEDA E SALVATAGGIO HOME TELEFONO
# =====================================================================================
import os
import io
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time as dtime

# CONTROLLO PRESENZA LOGO: Usa il file PNG se esiste, altrimenti usa l'emoji di backup
icona_app = "logo.png" if os.path.exists("logo.png") else "📅"

st.set_page_config(
    page_title="Ferie Gestori", 
    page_icon=icona_app, 
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    /* 🚫 RIMOZIONE RIGIDA DI QUALSIASI TOOLBAR, ICONA O FOOTER RESIDUO IN BASSO E IN ALTO */
    #MainMenu, footer, header, .stDecoration, [data-testid="stHeader"], [data-testid="stFooter"] {
        visibility: hidden !important;
        display: none !important;
    }
    
    /* 🚫 ELIMINAZIONE CHIRURGICA DELLA BARRA STREAMLIT IN BASSO */
    .stStatusWidget, [data-testid="stStatusWidget"], [data-testid="viewerToolbar"], [data-testid="stStatusWidgetContainer"], .stActionButton, [data-testid="stActionButton"] {
        display: none !important;
        visibility: hidden !important;
        height: 0px !important;
        width: 0px !important;
        opacity: 0 !important;
    }
    
    /* Sfondo dell'app chiaro e luminoso per la massima leggibilità */
    .stApp { background-color: #f8fafc !important; color: #1e293b !important; font-family: 'Segoe UI', sans-serif; }
    
    /* Titolo principale scuro e nitido */
    h1 { color: #115e59 !important; font-size: 28px !important; text-align: center; font-weight: 800 !important; margin-bottom: 25px; }
    
    /* Tutte le scritte e i titoli dei campi neri/blu scuro */
    .stMarkdown h3, label, p, [data-testid="stWidgetLabel"] p, .stSelectbox label { color: #1e293b !important; font-weight: 800 !important; font-size: 16px !important; opacity: 1 !important; }
    
    /* Contenitore del modulo con sfondo bianco puro e bordo grigio scuro marcato */
    div[data-testid="stForm"] { background-color: #ffffff !important; border: 2px solid #94a3b8 !important; border-radius: 14px !important; padding: 25px !important; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); }
    
    /* Caselle di testo, Date e Menu a tendina bianchi con testo nero e bordo scuro definito */
    input, div[data-baseweb="select"], div[data-baseweb="input"], select { background-color: #ffffff !important; color: #0f172a !important; border: 2px solid #64748b !important; border-radius: 8px !important; font-weight: 700 !important; }
    input, div[data-baseweb="select"] *, select { color: #0f172a !important; }
    
    /* PULSANTE VERDE SMERALDO: Più leggero ma ad alto contrasto */
    .stButton>button { background: linear-gradient(135deg, #0f766e 0%, #115e59 100%) !important; color: #ffffff !important; font-weight: 800 !important; font-size: 17px !important; width: 100%; border-radius: 10px !important; height: 54px !important; border: none !important; box-shadow: 0 4px 14px rgba(17, 94, 89, 0.3); }
    .stButton>button:hover { background: linear-gradient(135deg, #14b8a6 0%, #0f766e 100%) !important; box-shadow: 0 6px 20px rgba(20, 184, 166, 0.4); }
    
    /* Badge del Tecnico Loggato chiaro con bordo verde smeraldo */
    .user-badge { background-color: #ffffff; padding: 14px; border-radius: 10px; border: 2px solid #115e59; margin-bottom: 30px; text-align: center; color: #115e59 !important; font-weight: 800; font-size: 16px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
""", unsafe_allow_html=True)



# =====================================================================================
# BLOCCO 2: COLLEGAMENTO E CARICAMENTO AUTOMATICO DEI FILE EXCEL (LOCALI, TECNICI E STORICO)
# CONFIGURA GLI INDIRIZZI EMAIL AZIENDALI DI RIFERIMENTO E CARICA LE TABELLE IN MEMORIA
# =====================================================================================
FILE_LOCALI = "elenco_locali.xlsx"
FILE_TECNICI = "elenco_tecnici.xlsx"
FILE_STORICO_PERMANENTE = "storico_ferie.xlsx"

EMAIL_MITTENTE_GMAIL = "wingamingsrl@gmail.com"
EMAIL_MANUELA_RICEVENTE = "manuela.arigoni@wingaming.it"

def carica_database_locale():
    if not os.path.exists(FILE_LOCALI):
        dati_locali = {"CODICE_LOCALE": ["LOC001"], "NOME_LOCALE": ["Punto Vendita Demo"], "CONCESSIONARIO": ["Snaitech"]}
        pd.DataFrame(dati_locali).to_excel(FILE_LOCALI, index=False)
    if not os.path.exists(FILE_TECNICI):
        dati_tecnici = {"NOME": ["Manuela"], "EMAIL": ["manuela.arigoni@wingaming.it"], "PASSWORD": ["WinManuela4"]}
        pd.DataFrame(dati_tecnici).to_excel(FILE_TECNICI, index=False)
    
    df_l = pd.read_excel(FILE_LOCALI).fillna("")
    df_t = pd.read_excel(FILE_TECNICI).fillna("")
    
    if os.path.exists(FILE_STORICO_PERMANENTE):
        df_s = pd.read_excel(FILE_STORICO_PERMANENTE).fillna("")
    else:
        df_s = pd.DataFrame(columns=["DATA_INSERIMENTO", "TECNICO", "LOCALE", "INIZIO_FERIE", "FINE_FERIE", "COPIA_PROMEMORIA"])
    return df_l, df_t, df_s

df_locali, df_tecnici, df_storico_file = carica_database_locale()

if "storico_cloud" not in st.session_state:
    st.session_state.storico_cloud = df_storico_file.to_dict('records')

# =====================================================================================
# BLOCCO 3: SCHERMATA DI ACCESSO PROTETTA (LOGIN AZIENDALE)
# VERIFICA LE CREDENZIALI NEL FILE EXCEL ED ESTRAE IL NOME PULITO DEL TECNICO CHE EFFETTUA L'ACCESSO
# =====================================================================================
if "autenticato" not in st.session_state:
    st.markdown("<h1>🛡️ ACCESSO AREA TECNICI</h1>", unsafe_allow_html=True)
    with st.container(border=True):
        st.write("🔒 Autenticazione Richiesta")
        input_email = st.text_input("Nome Utente (E-mail):").strip().lower()
        input_password = st.text_input("Password di Sicurezza:", type="password").strip()
        if st.button("EFFETTUA IL LOGIN"):
            utente_valido = df_tecnici[(df_tecnici["EMAIL"].astype(str).str.strip().str.lower() == input_email) & (df_tecnici["PASSWORD"].astype(str).str.strip() == input_password)]
            if not utente_valido.empty:
                st.session_state.autenticato = True
                st.session_state.user_email = input_email
                st.session_state.user_nome = str(utente_valido["NOME"].values[0]).strip()
                st.rerun()
            else:
                st.error("❌ Credenziali errate. Riprova.")
    st.stop()

esecutore_nome = st.session_state.user_nome
esecutore_email = st.session_state.user_email

st.markdown("<h1>🛡️ SATELLITE FERIE GESTORI</h1>", unsafe_allow_html=True)
st.markdown(f"<div class='user-badge'>👤 {esecutore_nome} ({esecutore_email})</div>", unsafe_allow_html=True)

# =====================================================================================
# BLOCCO 4: MOTORE DI SPEDIZIONE EMAIL DIRETTO SU CASSAFORTE GOOGLE GMAIL
# COMPONE IL TESTO INSERENDO UN ELENCO PUNTATO PERFETTAMENTE ALLINEATO PER I CONCESSIONARI MULTIPLI
# =====================================================================================
def invia_mail_diretta_smtp(lista_m, locale, concessionario_testo, chiusura, riapertura, esecutore):
    try:
        pass_gmail = str(st.secrets["gmail"]["password_applicativa"]).strip()
        
        msg = MIMEMultipart()
        msg['From'] = EMAIL_MITTENTE_GMAIL
        msg['To'] = ", ".join(lista_m)
        msg['Subject'] = f"🛡️ Registrazione Chiusura Ferie - {locale}"
        
        linee_concessionari = ""
        elenco_conc = [c.strip() for c in concessionario_testo.split(",") if c.strip()]
        
        if len(elenco_conc) > 1:
            linee_concessionari = "\n" + "\n".join([f"                     • {c}" for c in elenco_conc])
        else:
            linee_concessionari = f" {concessionario_testo}"
            
        corpo = f"""Nuova chiusura ferie registrata nel sistema WinGaming.

Dettagli dell'inserimento:
--------------------------------------------------
👤 Tecnico Esecutore: {esecutore}
📍 Locale Coinvolto:  {locale}
🏢 Concessionario/i:{linee_concessionari}
📅 Inizio Chiusura:   {chiusura}
🚚 Data Riapertura:   {riapertura}
--------------------------------------------------

WINGAMING SRL"""
        
        msg.attach(MIMEText(corpo, 'plain'))
        
        server = smtplib.SMTP_SSL('64.233.184.108', 465, timeout=10)
        server.login(EMAIL_MITTENTE_GMAIL, pass_gmail)
        server.sendmail(EMAIL_MITTENTE_GMAIL, lista_m, msg.as_string())
        server.quit()
        return True, "OK"
    except Exception as e:
        return False, str(e)

# =====================================================================================
# BLOCCO 5: MODULO DI COMPILAZIONE (FORM CENTRALE) CON MENU A TENDINA COMPATTATO
# RAGGRUPPA I LOCALI CON LO STESSO CODICE ED ELENCA I CONCESSIONARI ASSOCIATI TRA PARENTESI
# =====================================================================================
if "form_id" not in st.session_state:
    st.session_state.form_id = 0

with st.form(key=f"modulo_ferie_{st.session_state.form_id}"):
    st.markdown("### 📝 Registra Chiusura Ferie")
    
    elenco_c = [f"{r['NOME']} ({r['EMAIL']})" for _, r in df_tecnici.iterrows() if str(r['EMAIL']).lower().strip() != esecutore_email.lower()]
    co_destinatario = st.selectbox("Invia copia promemoria a:", ["Nessun collega"] + elenco_c)
    
    st.markdown("---")
    
    locali_raggruppati = {}
    mappa_concessionari = {}
    
    for _, r in df_locali.iterrows():
        cod_loc = str(r['CODICE_LOCALE']).strip()
        nome_loc = str(r['NOME_LOCALE']).strip()
        conc_loc = str(r['CONCESSIONARIO']).strip()
        
        chiave_chiave = f"{cod_loc} - {nome_loc}"
        
        # CORREZIONE ERRORE: Rimosse le "s" errate per allinearsi alla variabile corretta
        if chiave_chiave not in locali_raggruppati:
            locali_raggruppati[chiave_chiave] = []
        if conc_loc and conc_loc not in locali_raggruppati[chiave_chiave]:
            locali_raggruppati[chiave_chiave].append(conc_loc)

    lista_pvd = ["- Selezionare il Locale -"]
    for etichetta, lista_conc in locali_raggruppati.items():
        concessionari_uniti = ", ".join(lista_conc)
        mappa_concessionari[etichetta] = concessionari_uniti
        lista_pvd.append(f"{etichetta} ({concessionari_uniti})")
        
    scelta_pvd = st.selectbox("Seleziona o cerca locale:", lista_pvd, index=0)
    
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1: data_chiusura = st.date_input("Giorno Chiusura:", datetime.now(), format="DD-MM-YYYY")
    with col2: ora_chiusura = st.time_input("Ora Chiusura:", dtime(6, 0))
    
    st.markdown("---")
    col3, col4 = st.columns(2)
    with col3: data_riapertura = st.date_input("Giorno Riapertura:", datetime.now() + timedelta(days=14), format="DD-MM-YYYY")
    with col4: ora_riapertura = st.time_input("Ora Riapertura:", dtime(12, 0))
    
    forza_sovrascrittura = st.checkbox("⚠️ Spunta questa casella per confermare la modifica/sovrascrittura del periodo passato")
    
    submit_button = st.form_submit_button("🚀 INVIA E REGISTRA CHIUSURA")

# =====================================================================================
# BLOCCO 6: VALIDAZIONE, NOTIFICA EMAIL E INIETTORE AUTOMATICO DIRETTO IN SANSONE
# VERSIONE FINALE ALLINEATA SUL LINK DIRETTO GESTIONALE.GAMESLODI.IT/LOCALI/EDIT
# =====================================================================================
if submit_button:
    if scelta_pvd == "- Selezionare il Locale -":
        st.error("Errore: Seleziona un locale valido.")
    elif datetime.combine(data_riapertura, ora_riapertura) <= datetime.combine(data_chiusura, ora_chiusura):
        st.error("Errore: La data di riapertura deve essere successiva alla chiusura.")
    else:
        new_inizio = data_chiusura
        new_fine = data_riapertura
        
        sovrapposizione_rilevata = False
        riga_conflitto_idx = None
        dettagli_conflitto = ""
        
        for idx, row in enumerate(st.session_state.storico_cloud):
            if str(row["LOCALE"]).strip() == str(scelta_pvd).strip():
                try:
                    old_inizio = datetime.strptime(row["INIZIO_FERIE"], "%d-%m-%Y").date()
                    old_fine = datetime.strptime(row["FINE_FERIE"], "%d-%m-%Y").date()
                    if (new_inizio <= old_fine) and (new_fine >= old_inizio):
                        sovrapposizione_rilevata = True
                        riga_conflitto_idx = idx
                        dettagli_conflitto = f"Dal {row['INIZIO_FERIE']} al {row['FINE_FERIE']} (Inserito da: {row['TECNICO']})"
                        break
                except Exception: continue

        if sovrapposizione_rilevata and not forza_sovrascrittura:
            st.error(f"⚠️ ATTENZIONE: Questo locale risulta già chiuso nel periodo richiesto!\n\n📌 **Periodo registrato:** {dettagli_conflitto}.\n\nSe desideri modificare o aggiornare questo periodo con le nuove date, spunta la casella di conferma in fondo al modulo e premi di nuovo il pulsante di invio.")
        else:
            str_c = f"{data_chiusura.strftime('%d-%m-%Y')} {ora_chiusura.strftime('%H:%M')}"
            str_r = f"{data_riapertura.strftime('%d-%m-%Y')} {ora_riapertura.strftime('%H:%M')}"
            nuova = {"DATA_INSERIMENTO": datetime.now().strftime("%d-%m-%Y %H:%M:%S"), "TECNICO": esecutore_nome, "LOCALE": scelta_pvd, "INIZIO_FERIE": data_chiusura.strftime('%d-%m-%Y'), "FINE_FERIE": data_riapertura.strftime('%d-%m-%Y'), "COPIA_PROMEMORIA": co_destinatario}
            
            testo_selezione = str(scelta_pvd)
            # CORREZIONE TASSATIVA: Inserito l'indice [0] prima dello strip per evitare l'AttributeError
            chiave_pulita = testo_selezione.split(" (")[0].strip() if " (" in testo_selezione else testo_selezione.strip()
            concessionario_estratto = mappa_concessionari.get(chiave_pulita, "")
            
            lista_m = [EMAIL_MANUELA_RICEVENTE, esecutore_email]
            if co_destinatario != "Nessun collega":
                mail_pulita = co_destinatario.split(" (")[-1].replace(")", "").strip()
                lista_m.append(mail_pulita)
                
            with st.spinner("Salvataggio e sincronizzazione database..."):
                invio_ok, risposta_server = invia_mail_diretta_smtp(lista_m, chiave_pulita, concessionario_estratto, str_c, str_r, esecutore_nome)
            
            if invio_ok:
                if sovrapposizione_rilevata and riga_conflitto_idx is not None:
                    st.session_state.storico_cloud.pop(riga_conflitto_idx)
                    
                st.session_state.storico_cloud.append(nuova)
                
                try:
                    pd.DataFrame(st.session_state.storico_cloud).to_excel(FILE_STORICO_PERMANENTE, index=False)
                except Exception: pass
                
                # 🚀 --- MOTORE INIETTORE SU LINK CORRETTO GESTIONALE ---
                status_sansone = ""
                try:
                    import requests
                    import time
                    
                    s_user = str(st.secrets["sansone"]["email"]).strip()
                    s_pass = str(st.secrets["sansone"]["password"]).strip()
                    
                    sessione = requests.Session()
                    
                    # Login puntato rigidamente al sottodominio gestionale
                    url_login = "https://gameslodi.it"
                    res_login = sessione.post(url_login, data={"email": s_user, "password": s_pass}, timeout=7)
                    
                    ts = int(time.time())
                    
                    payload_sansone = {
                        "localeId": "10900",  
                        "localeCodice": str(chiave_pulita),
                        f"ferie[{ts}][data_inizio]": str(nuova["INIZIO_FERIE"].replace("-", "/")),
                        f"ferie[{ts}][data_fine]": str(nuova["FINE_FERIE"].replace("-", "/")),
                        f"ferie[{ts}][note]": "Inserimento automatico App"
                    }
                    
                    # Invio alla rotta di edit corretta ricavata da F12
                    url_salva = "https://gameslodi.it"
                    res_salva = sessione.post(url_salva, data=payload_sansone, timeout=7)
                    
                    if res_salva.status_code == 200:
                        status_sansone = "✅ Gestionale Sansone: Aggiornato in automatico!"
                    else:
                        status_sansone = f"⚠️ Gestionale Sansone: Errore di sincronizzazione (Codice {res_salva.status_code})"
                except Exception as e:
                    status_sansone = f"❌ Sincronizzazione fallita: {str(e)}"
                
                st.success("✅ OPERAZIONE COMPLETATA!\n\n📧 Registro aggiornato e notifica inviata.")
                if esecutore_email.lower() == EMAIL_MANUELA_RICEVENTE.lower():
                    st.info(status_sansone)
                    
                st.session_state.form_id += 1
                time.sleep(4)
                st.rerun()
            else:
                st.error(f"❌ Errore Google SMTP: {risposta_server}. Verifica la password applicativa nei Secrets.")

st.markdown("---")
st.markdown("### 📅 Promemoria Giri Logistici (Preavviso 3 Giorni)")
oggi = datetime.now().date()
alert_c, alert_r = [], []
for row in st.session_state.storico_cloud:
    try:
        d_i = datetime.strptime(row["INIZIO_FERIE"], "%d-%m-%Y").date()
        d_f = datetime.strptime(row["FINE_FERIE"], "%d-%m-%Y").date()
        if d_i - oggi == timedelta(days=3): alert_c.append(f"⚠️ **{row['LOCALE']}** chiude tra 3 giorni")
        if d_f - oggi == timedelta(days=3): alert_r.append(f"🚚 **{row['LOCALE']}** riapre tra 3 giorni")
    except Exception: continue
for a in alert_c: st.error(a)
for r in alert_r: st.warning(r)
if not alert_c and not alert_r: st.write("✅ Nessun adempimento logistico per i primi 3 giorni di scadenza.")

if st.sidebar.button("🚪 Disconnetti Account"):
    del st.session_state.autenticato
    st.rerun()

if esecutore_email.lower() == EMAIL_MANUELA_RICEVENTE.lower():
    st.markdown("<br>### 📊 Registro Storico Chiusure Centralizzato", unsafe_allow_html=True)
    if st.session_state.storico_cloud:
        df_vis = pd.DataFrame(st.session_state.storico_cloud)
        st.dataframe(df_vis, hide_index=True)
        with io.BytesIO() as buffer:
            df_vis.to_excel(buffer, index=False)
            st.download_button(label="📥 Scarica File Excel Aggiornato", data=buffer.getvalue(), file_name="storico_ferie.xlsx", mime="application/vnd.ms-excel")

