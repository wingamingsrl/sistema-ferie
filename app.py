import os
import io
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time as dtime

st.set_page_config(page_title="Ferie Gestori", page_icon="🛡️", layout="centered")

st.markdown("""
    <style>
    .stApp { background-color: #0f0f14; color: #e2e8f0; font-family: 'Segoe UI', sans-serif; }
    h1 { color: #3b82f6; text-shadow: 0px 0px 10px rgba(59, 130, 246, 0.5); font-size: 24px !important; text-align: center; font-weight: 800 !important; }
    div[data-testid="stForm"] { background-color: #1e1e26 !important; border: 1px solid #2d2d3d !important; border-radius: 12px !important; padding: 20px !important; }
    div[data-baseweb="select"], input { background-color: #252538 !important; color: #ffffff !important; border-radius: 6px !important; }
    .stButton>button { background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%) !important; color: #ffffff !important; font-weight: bold !important; width: 100%; border-radius: 10px !important; height: 50px !important; border: none !important; }
    .user-badge { background-color: #1e1e26; padding: 12px; border-radius: 8px; border: 1px solid #3b82f6; margin-bottom: 25px; text-align: center; color: #e2e8f0; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

FILE_LOCALI = "elenco_locali.xlsx"
FILE_TECNICI = "elenco_tecnici.xlsx"
FILE_STORICO_PERMANENTE = "registro_ferie_salvato.xlsx"

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

# --- 🔐 SCHERMATA DI LOGIN AZIENDALE ---
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
                st.session_state.user_nome = str(utente_valido.loc[:, "NOME"].values).strip()
                st.rerun()
            else:
                st.error("❌ Credenziali errate. Riprova.")
    st.stop()

esecutore_nome = st.session_state.user_nome
esecutore_email = st.session_state.user_email

st.markdown("<h1>🛡️ SATELLITE FERIE GESTORI</h1>", unsafe_allow_html=True)
st.markdown(f"<div class='user-badge'>👤 {esecutore_nome} ({esecutore_email})</div>", unsafe_allow_html=True)

def invia_mail_diretta_smtp(lista_m, locale, chiusura, riapertura, esecutore):
    try:
        pass_gmail = str(st.secrets["gmail"]["password_applicativa"]).strip()
        
        msg = MIMEMultipart()
        msg['From'] = EMAIL_MITTENTE_GMAIL
        msg['To'] = ", ".join(lista_m)
        msg['Subject'] = f"🛡️ Registrazione Chiusura Ferie - {locale}"
        
        corpo = f"Nuova chiusura ferie registrata.\n\nLocale: {locale}\nTecnico: {esecutore}\nInizio: {chiusura}\nRiapertura: {riapertura}"
        msg.attach(MIMEText(corpo, 'plain'))
        
        # 🚀 ACCESSO AD IP DIRETTO DI GOOGLE (AGGIRA IL BLOCCO DI RESOLUTION INTERNO)
        server = smtplib.SMTP_SSL('64.233.184.108', 465, timeout=10)
        server.login(EMAIL_MITTENTE_GMAIL, pass_gmail)
        server.sendmail(EMAIL_MITTENTE_GMAIL, lista_m, msg.as_string())
        server.quit()
        return True, "OK"
    except Exception as e:
        return False, str(e)

if "form_id" not in st.session_state:
    st.session_state.form_id = 0

with st.form(key=f"modulo_ferie_{st.session_state.form_id}"):
    st.markdown("### 📝 Registra Chiusura Ferie")
    elenco_c = [f"{r['NOME']} ({r['EMAIL']})" for _, r in df_tecnici.iterrows() if str(r['EMAIL']).lower().strip() != esecutore_email.lower()]
    co_destinatario = st.selectbox("Invia copia promemoria a:", ["Nessun collega"] + elenco_c)
    st.markdown("---")
    filtro_testo = st.text_input("🔍 Cerca Locale:").strip().lower()
    lista_pvd = ["- Selezionare il Locale -"]
    for _, r in df_locali.iterrows():
        if not filtro_testo or filtro_testo in str(r["NOME_LOCALE"]).lower() or filtro_testo in str(r["CODICE_LOCALE"]).lower():
            lista_pvd.append(f"{r['CODICE_LOCALE']} - {r['NOME_LOCALE']} ({r['CONCESSIONARIO']})")
    scelta_pvd = st.selectbox("Seleziona locale:", lista_pvd, index=0)
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1: data_chiusura = st.date_input("Giorno Chiusura:", datetime.now(), format="DD-MM-YYYY")
    with col2: ora_chiusura = st.time_input("Ora Chiusura:", dtime(6, 0))
    st.markdown("---")
    col3, col4 = st.columns(2)
    with col3: data_riapertura = st.date_input("Giorno Riapertura:", datetime.now() + timedelta(days=14), format="DD-MM-YYYY")
    with col4: ora_riapertura = st.time_input("Ora Riapertura:", dtime(12, 0))
    submit_button = st.form_submit_button("🚀 INVIA E REGISTRA CHIUSURA")

if submit_button:
    if scelta_pvd == "- Selezionare il Locale -":
        st.error("Errore: Seleziona un locale valido.")
    elif datetime.combine(data_riapertura, ora_riapertura) <= datetime.combine(data_chiusura, ora_chiusura):
        st.error("Errore: La data di riapertura deve essere successiva alla chiusura.")
    else:
        str_c = f"{data_chiusura.strftime('%d-%m-%Y')} {ora_chiusura.strftime('%H:%M')}"
        str_r = f"{data_riapertura.strftime('%d-%m-%Y')} {ora_riapertura.strftime('%H:%M')}"
        nuova = {"DATA_INSERIMENTO": datetime.now().strftime("%d-%m-%Y %H:%M:%S"), "TECNICO": esecutore_nome, "LOCALE": scelta_pvd, "INIZIO_FERIE": data_chiusura.strftime('%d-%m-%Y'), "FINE_FERIE": data_riapertura.strftime('%d-%m-%Y'), "COPIA_PROMEMORIA": co_destinatario}
        
        lista_m = [EMAIL_MANUELA_RICEVENTE, esecutore_email]
        if co_destinatario != "Nessun collega":
            mail_pulita = co_destinatario.split(" (")[-1].replace(")", "").strip()
            lista_m.append(mail_pulita)
            
        with st.spinner("Salvataggio in corso..."):
            invio_ok, risposta_server = invia_mail_diretta_smtp(lista_m, scelta_pvd, str_c, str_r, esecutore_nome)
        
        if invio_ok:
            st.session_state.storico_cloud.append(nuova)
            pd.DataFrame(st.session_state.storico_cloud).to_excel(FILE_STORICO_PERMANENTE, index=False)
            st.success(f"✅ OPERAZIONE COMPLETATA!\n\n📧 **Notifica inviata con successo a:** {', '.join(lista_m)}")
            st.session_state.form_id += 1
            time.sleep(5)
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
            st.download_button(label="📥 Scarica Registro Storico in Excel", data=buffer.getvalue(), file_name="registro_chiusure_wingaming.xlsx", mime="application/vnd.ms-excel")
    else:
        st.write("Nessuna chiusura presente nel registro.")

