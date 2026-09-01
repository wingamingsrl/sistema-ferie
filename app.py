# =====================================================================================
# BLOCCO 1: STRUTTURA DI BASE, CONFIGURAZIONE ICONA E PULIZIA INTERFACCIA UTENTE
# CONFIGURA L'ICONA PROPRIETARIA PNG ANCHE PER IL SALVATAGGIO SULLA HOME DI IOS E ANDROID
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
import requests
import base64

st.set_page_config(
    page_title="Ferie Gestori", 
    page_icon="logo.png" if os.path.exists("logo.png") else "📅", 
    layout="centered",
    initial_sidebar_state="collapsed"
)

# MARCATORE GRAFICO SMARTPHONE: Costringe iOS e Android a usare il logo.png per la Home
st.markdown("""
    <link rel="apple-touch-icon" sizes="180x190" href="logo.png">
    <link rel="icon" type="image/png" sizes="192x192" href="logo.png">
    <link rel="shortcut icon" href="logo.png">
""", unsafe_allow_html=True)

st.markdown("""
    <style>
    #MainMenu, footer, header, .stDecoration, [data-testid="stHeader"], [data-testid="stFooter"] {
        visibility: hidden !important; display: none !important;
    }
    .stStatusWidget, [data-testid="stStatusWidget"], [data-testid="viewerToolbar"], [data-testid="stStatusWidgetContainer"], .stActionButton, [data-testid="stActionButton"] {
        display: none !important; visibility: hidden !important; height: 0px !important; width: 0px !important; opacity: 0 !important;
    }
    .stApp { background-color: #f8fafc !important; color: #1e293b !important; font-family: 'Segoe UI', sans-serif; }
    h1 { color: #115e59 !important; font-size: 28px !important; text-align: center; font-weight: 800 !important; margin-bottom: 25px; }
    .stMarkdown h3, label, p, [data-testid="stWidgetLabel"] p, .stSelectbox label { color: #1e293b !important; font-weight: 800 !important; font-size: 16px !important; opacity: 1 !important; }
    div[data-testid="stForm"] { background-color: #ffffff !important; border: 2px solid #94a3b8 !important; border-radius: 14px !important; padding: 25px !important; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); }
    input, div[data-baseweb="select"], div[data-baseweb="input"], select { background-color: #ffffff !important; color: #0f172a !important; border: 2px solid #64748b !important; border-radius: 8px !important; font-weight: 700 !important; }
    input, div[data-baseweb="select"] *, select { color: #0f172a !important; }
    .stButton>button { background: linear-gradient(135deg, #0f766e 0%, #115e59 100%) !important; color: #ffffff !important; font-weight: 800 !important; font-size: 17px !important; width: 100%; border-radius: 10px !important; height: 54px !important; border: none !important; box-shadow: 0 4px 14px rgba(17, 94, 89, 0.3); }
    .stButton>button:hover { background: linear-gradient(135deg, #14b8a6 0%, #0f766e 100%) !important; box-shadow: 0 6px 20px rgba(20, 184, 166, 0.4); }
    .user-badge { background-color: #ffffff; padding: 14px; border-radius: 10px; border: 2px solid #115e59; margin-bottom: 30px; text-align: center; color: #115e59 !important; font-weight: 800; font-size: 16px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
""", unsafe_allow_html=True)

# =====================================================================================
# BLOCCO 2: COLLEGAMENTO FILE EXCEL PERMANENTI E ALLINEAMENTO MEMORIA CLOUD
# VERSIONE DI PRODUZIONE SIGILLATA — REINTEGRAZIONE SENZA CACHE FORZATA ALL'ULTIMO SECONDO
# =====================================================================================
FILE_LOCALI = "elenco_locali.xlsx"
FILE_TECNICI = "elenco_tecnici.xlsx"
FILE_STORICO_PERMANENTE = "storico_ferie.xlsx"

EMAIL_MITTENTE_GMAIL = "wingamingsrl@gmail.com"
EMAIL_MANUELA_RICEVENTE = "manuela.arigoni@wingaming.it"

def scarica_file_da_github_se_esiste(nome_file):
    try:
        t_git = str(st.secrets["github"]["token_accesso"]).strip()
        
        # 🛡️ DISTRUZIONE CACHE CLOUD: Generiamo un parametro temporale unico al millisecondo (timestamp)
        # Questo costringe il server di Streamlit a ignorare i dati vecchi e a prendere il file reale da GitHub
        timestamp_live = int(time.time())
        url_git = f"https://github.com{nome_file}?t={timestamp_live}"
        
        h = {
            "Authorization": f"Bearer {t_git}", 
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": f"WinGaming-Cloud-Engine-{timestamp_live}"
        }
        
        # Effettuiamo la chiamata forzata distruggendo la vecchia memoria
        risposta = requests.get(url_git, headers=h, timeout=7)
        if risposta.status_code == 200:
            b64_content = risposta.json().get("content", "")
            # Decodifica l'Excel fresco di secondo direttamente in RAM
            return pd.read_excel(io.BytesIO(base64.b64decode(b64_content)))
    except Exception:
        pass
    return None

def carica_database_locale():
    df_l = pd.read_excel(FILE_LOCALI).fillna("") if os.path.exists(FILE_LOCALI) else pd.DataFrame(columns=["CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO"])
    df_t = pd.read_excel(FILE_TECNICI).fillna("") if os.path.exists(FILE_TECNICI) else pd.DataFrame(columns=["NOME", "EMAIL", "PASSWORD", "RUOLO"])
    
    df_s = scarica_file_da_github_se_esiste(FILE_STORICO_PERMANENTE)
    if df_s is None:
        if os.path.exists(FILE_STORICO_PERMANENTE):
            df_s = pd.read_excel(FILE_STORICO_PERMANENTE).fillna("")
        else:
            df_s = pd.DataFrame(columns=["DATA_INSERIMENTO", "TECNICO", "LOCALE", "INIZIO_FERIE", "FINE_FERIE", "COPIA_PROMEMORIA"])
    return df_l, df_t, df_s.fillna("")

# Forza lo scaricamento pulito ad ogni singolo rinfresco della pagina web
df_locali, df_tecnici, df_storico_file = carica_database_locale()

# 🛡️ RE-INIZIALIZZAZIONE DELLA SESSIONE: Sovrascrive la memoria vecchia con il file reale di GitHub
st.session_state.storico_cloud = df_storico_file.to_dict('records')

def push_excel_su_github(df_da_salvare):
    try:
        t_git = str(st.secrets["github"]["token_accesso"]).strip()
        url_git = f"https://github.com{FILE_STORICO_PERMANENTE}"
        
        output_binario = io.BytesIO()
        df_da_salvare.to_excel(output_binario, index=False)
        dati_base64 = base64.b64encode(output_binario.getvalue()).decode('utf-8')
        
        headers_git = {
            "Authorization": f"Bearer {t_git}", 
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "WinGaming-Cloud-App"
        }
        
        # Recupera lo SHA passando esplicitamente il parametro del ramo principale per sbloccare la sovrascrittura
        res_get = requests.get(url_git, headers=headers_git, params={"ref": "main"}, timeout=5)
        sha_file = res_get.json().get("sha", "") if res_get.status_code == 200 else ""
        
        payload_git = {
            "message": "🤖 [App] Sincronizzazione permanente ed allineamento database ferie", 
            "content": dati_base64,
            "branch": "main"
        }
        
        if sha_file: 
            payload_git["sha"] = sha_file
            
        risposta_put = requests.put(url_git, json=payload_git, headers=headers_git, timeout=5)
        
        if risposta_put.status_code == 200 or risposta_put.status_code == 201:
            return True
        else:
            return False
    except Exception:
        return False

df_locali, df_tecnici, df_storico_file = carica_database_locale()

# 🛡️ CORREZIONE MEMORIA: Forza l'aggiornamento della sessione a ogni avvio della pagina
# Questo distrugge la persistenza della RAM del server e la riallinea a GitHub
st.session_state.storico_cloud = df_storico_file.to_dict('records')



# =====================================================================================
# BLOCCO 3: AUTENTICAZIONE E GESTIONE CREDENZIALI DINAMICHE DA EXCEL (RUOLI)
# IDENTIFICA IL TECNICO LOGGATO E MEMORIZZA SE È UN AMMINISTRATORE O UN TECNICO SEMPLICE
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
                
                # 🛡️ SOLUZIONE DEFINITIVA LOOP: Inserito l'indice [0] su tutti i valori estratti dal DataFrame
                st.session_state.user_nome = str(utente_valido["NOME"].values[0]).strip()
                
                ruolo_estratto = str(utente_valido["RUOLO"].values[0]).strip().lower() if "RUOLO" in utente_valido.columns else "tecnico"
                st.session_state.user_ruolo = ruolo_estratto
                st.rerun()
            else:
                st.error("❌ Credenziali errate. Riprova.")
    st.stop()

esecutore_nome = st.session_state.user_nome
esecutore_email = st.session_state.user_email
esecutore_ruolo = st.session_state.get("user_ruolo", "tecnico")

st.markdown("<h1>🛡️ SATELLITE FERIE GESTORI</h1>", unsafe_allow_html=True)
st.markdown(f"<div class='user-badge'>👤 {esecutore_nome} ({esecutore_email}) — Ruolo: {esecutore_ruolo.upper()}</div>", unsafe_allow_html=True)



# =====================================================================================
# BLOCCO 4: MOTORE NOTIFICA EMAIL SMTP GOOGLE CON ELENCO CONCESSIONARI INCOLONNATO
# CONNETTE IL CANALE CIFRETO VERSO GMAIL GESTENDO I DESTINATARI MULTIPLI E CO-TECNICI
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
# BLOCCO 6: VERIFICA SOVRAPPOSIZIONI, CARICAMENTO SU GITHUB E AREA AMMINISTRATORE DINAMICA
# VERSIONE DI PRODUZIONE SIGILLATA — RISOLTI I CONFLITTI DI MEMORIA E PERCORSI ID
# =====================================================================================
if submit_button:
    if scelta_pvd == "- Selezionare il Locale -":
        st.error("Errore: Seleziona un locale valido.")
    elif datetime.combine(data_riapertura, ora_riapertura) <= datetime.combine(data_chiusura, ora_chiusura):
        st.error("Errore: La data di riapertura deve essere successiva alla chiusura.")
    else:
        new_inizio, new_fine = data_chiusura, data_riapertura
        sovrapposizione_rilevata, riga_conflitto_idx, dettagli_conflitto = False, None, ""
        
        for idx, row in enumerate(st.session_state.storico_cloud):
            if str(row["LOCALE"]).strip() == str(scelta_pvd).strip():
                try:
                    data_inizio_estratta = str(row["INIZIO_FERIE"]).split(" ")[0]
                    data_fine_estratta = str(row["FINE_FERIE"]).split(" ")[0]
                    
                    old_inizio = datetime.strptime(data_inizio_estratta, "%d-%m-%Y").date()
                    old_fine = datetime.strptime(data_fine_estratta, "%d-%m-%Y").date()
                    if (new_inizio <= old_fine) and (new_fine >= old_inizio):
                        sovrapposizione_rilevata, riga_conflitto_idx = True, idx
                        dettagli_conflitto = f"Dal {row['INIZIO_FERIE']} al {row['FINE_FERIE']} (Inserito da: {row['TECNICO']})"
                        break
                except Exception: continue

        if sovrapposizione_rilevata and not forza_sovrascrittura:
            st.error(f"⚠️ ATTENZIONE: Questo locale risulta già chiuso nel periodo richiesto!\n\n📌 **Periodo registrato:** {dettagli_conflitto}.\n\nSe desideri modificare o aggiornare questo periodo con le nuove date, spunta la casella di conferma in fondo al modulo e primi di nuovo il pulsante di invio.")
        else:
            str_c = f"{data_chiusura.strftime('%d-%m-%Y')} {ora_chiusura.strftime('%H:%M')}"
            str_r = f"{data_riapertura.strftime('%d-%m-%Y')} {ora_riapertura.strftime('%H:%M')}"
            
            nuova = {
                "DATA_INSERIMENTO": datetime.now().strftime("%d-%m-%Y %H:%M:%S"), 
                "TECNICO": esecutore_nome, 
                "LOCALE": scelta_pvd, 
                "INIZIO_FERIE": str_c,   
                "FINE_FERIE": str_r,     
                "COPIA_PROMEMORIA": co_destinatario
            }
            
            concessionario_estratto = mappa_concessionari.get(scelta_pvd, "")
            chiave_pulita = scelta_pvd.split(" (")[0].strip() if " (" in scelta_pvd else scelta_pvd.strip()
            
            lista_m = [EMAIL_MANUELA_RICEVENTE, esecutore_email]
            if co_destinatario != "Nessun collega":
                mail_pulita = co_destinatario.split(" (")[-1].replace(")", "").strip() if " (" in co_destinatario else co_destinatario.strip()
                lista_m.append(mail_pulita)
                
            with st.spinner("Salvataggio e invio notifica..."):
                invio_ok, risposta_server = invia_mail_diretta_smtp(lista_m, chiave_pulita, concessionario_estratto, str_c, str_r, esecutore_nome)
            
            if invio_ok:
                if sovrapposizione_rilevata and riga_conflitto_idx is not None:
                    st.session_state.storico_cloud.pop(riga_conflitto_idx)
                st.session_state.storico_cloud.append(nuova)
                
                df_salva = pd.DataFrame(st.session_state.storico_cloud)
                df_salva.to_excel(FILE_STORICO_PERMANENTE, index=False)
                push_excel_su_github(df_salva)
                
                st.success("✅ OPERAZIONE COMPLETATA!\n\n📧 Registro allineato su GitHub e e-mail inviata.")
                st.session_state.form_id += 1
                time.sleep(2)
                st.rerun()
            else:
                st.error(f"❌ Errore Google SMTP: {risposta_server}. Spedizione e-mail fallita.")

st.markdown("---")
st.markdown("### 📅 Promemoria Giri Logistici (Preavviso 3 Giorni)")
oggi = datetime.now().date()
alert_c, alert_r = [], []
for row in st.session_state.storico_cloud:
    try:
        data_inizio_estratta = str(row["INIZIO_FERIE"]).split(" ")[0]
        data_fine_estratta = str(row["FINE_FERIE"]).split(" ")[0]
        
        d_i = datetime.strptime(data_inizio_estratta, "%d-%m-%Y").date()
        d_f = datetime.strptime(data_fine_estratta, "%d-%m-%Y").date()
        if d_i - oggi == timedelta(days=3): alert_c.append(f"⚠️ **{row['LOCALE']}** chiude tra 3 giorni")
        if d_f - oggi == timedelta(days=3): alert_r.append(f"🚚 **{row['LOCALE']}** riapre tra 3 giorni")
    except Exception: continue
for a in alert_c: st.error(a)
for r in alert_r: st.warning(r)
if not alert_c and not alert_r: st.write("✅ Nessun adempimento logistico per i primi 3 giorni di scadenza.")

if st.sidebar.button("🚪 Disconnetti Account"):
    del st.session_state.autenticato
    st.rerun()

# =====================================================================================
# PANNELLO AMMINISTRATORE CENTRALIZZATO 
# =====================================================================================
if esecutore_ruolo == "admin":
    st.markdown("<br>### 📊 Registro Storico Chiusure Centralizzato", unsafe_allow_html=True)
    if st.session_state.storico_cloud:
        df_vis = pd.DataFrame(st.session_state.storico_cloud)
        st.dataframe(df_vis, hide_index=True)
        
        with io.BytesIO() as buffer:
            df_vis.to_excel(buffer, index=False)
            st.download_button(label="📥 Scarica Registro Excel Storico", data=buffer.getvalue(), file_name="storico_ferie.xlsx", mime="application/vnd.ms-excel")
            
        st.markdown("---")
        st.markdown("### 🏢 Locali SNAITECH da inserire a sistema")
        
        righe_snaitech = [row for row in st.session_state.storico_cloud if "snaitech" in str(row["LOCALE"]).lower() or "snai" in str(row["LOCALE"]).lower()]
        
        if righe_snaitech:
            df_snai = pd.DataFrame(righe_snaitech)
            st.dataframe(df_snai[["LOCALE", "INIZIO_FERIE", "FINE_FERIE", "TECNICO"]], hide_index=True)
        else:
            st.write("✅ Nessuna chiusura attiva per locali Snaitech.")
            
        st.markdown("---")
        st.markdown("### 🗑️ Cancella un Periodo Registrato (Se il cliente cambia idea)")
        
        opzioni_cancellazione = ["- Seleziona la riga da eliminare -"]
        for idx, row in enumerate(st.session_state.storico_cloud):
            opzioni_cancellazione.append(f"ID {idx} | {row['LOCALE']} (Dal {row['INIZIO_FERIE']} al {row['FINE_FERIE']})")
            
        # 🛡️ LA VARIABILE ALLINEATA: `selezione_delete`
        selezione_delete = st.selectbox("Scegli la chiusura da eliminare dal database:", opzioni_cancellazione)
        
        if selezione_delete != "- Seleziona la riga da eliminare -":
            # 🛡️ FIX CHIRURGICO FLUIDO: Isola in modo matematico l'ID riga evitando crash di split
            stringa_id = selezione_delete.split("ID ")[1]
            idx_da_eliminare = int(stringa_id.split(" |")[0])
            
            if st.button("❌ ELIMINA DEFINITIVAMENTE QUESTA CHIUSURA"):
                with st.spinner("Rimozione e riallineamento database cloud..."):
                    # Rimuove l'elemento dalla sessione cloud di origine dei dati
                    st.session_state.storico_cloud.pop(idx_da_eliminare)
                    
                    # Genera la tabella aggiornata e invia a GitHub
                    df_nuovo_salva = pd.DataFrame(st.session_state.storico_cloud)
                    df_nuovo_salva.to_excel(FILE_STORICO_PERMANENTE, index=False)
                    push_excel_su_github(df_nuovo_salva)
                    
                st.success("🗑️ Chiusura eliminata con successo! Il database Excel su GitHub è stato aggiornato.")
                time.sleep(2)
                st.rerun()
    else:
        st.write("Nessuna chiusura presente nel registro.")
