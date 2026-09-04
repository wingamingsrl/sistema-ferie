# =====================================================================================
# SW GENERALE DI INSERIMENTO FERIE GESTORI WIN GAMING — PRODUZIONE INTEGRALE FINALE
# BLOCCO 1: CARICAMENTO MODULI E CONFIGURAZIONE STILE GRAFICO APPLICAZIONE CLOUD
# =====================================================================================
import os
import io
import time
import base64
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime, timedelta, time as dtime

icona_app = "logo.png" if os.path.exists("logo.png") else "📅"

st.set_page_config(
    page_title="Ferie Gestori", 
    page_icon=icona_app, 
    layout="centered",
    initial_sidebar_state="collapsed"
)

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
    .stButton>button { background: linear-gradient(135deg, #0f766e 0%, #115e59 100%) !important; color: #ffffff !important; font-weight: 800 !important; font-size: 17px !important; width: 100%; border-radius: 10px !important; height: 54px !important; border: none !important; box-shadow: 0 4px 14px rgba(17, 94, 89, 0.3); }
    .user-badge { background-color: #ffffff; padding: 14px; border-radius: 10px; border: 2px solid #115e59; margin-bottom: 30px; text-align: center; color: #115e59 !important; font-weight: 800; font-size: 16px; }
    </style>
""", unsafe_allow_html=True)

# =====================================================================================
# BLOCCO 2: COLLEGAMENTO FILE EXCEL PERMANENTI E PARSAMENTO REALE CONTRACCOSMICO
# VERSIONE DI PRODUZIONE 100% EXCEL NATIVO — BLINDATURA CANCELLAZIONI E MODIFICHE REALI
# =====================================================================================
FILE_LOCALI = "elenco_locali.xlsx"
FILE_TECNICI = "elenco_tecnici.xlsx"
FILE_STORICO_PERMANENTE = "storico_ferie.xlsx"

EMAIL_MITTENTE_GMAIL = "wingamingsrl@gmail.com"
EMAIL_MANUELA_RICEVENTE = "manuela.arigoni@wingaming.it"

COLONNE_REALI_UFFICIO = ["DATA_INSERIMENTO", "TECNICO_INSERIMENTO", "CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO", "INIZIO_FERIE", "FINE_FERIE", "PROMEMORIA_IN_COPIA", "STATO_INVIO"]

def scarica_file_da_github_se_esiste(nome_file):
    try:
        t_git = str(st.secrets["github"]["token_accesso"]).strip()
        # Genera un marcatore di millisecondi per costringere GitHub a ignorare la cache vecchia
        c_time = str(int(time.time() * 1000))
        
        indirizzo_base = "https://github.com"
        url_git = indirizzo_base + "/" + str(nome_file) + "?_nonce=" + c_time
        
        h = {
            "Authorization": f"token {t_git}", 
            "Accept": "application/vnd.github.v3.raw",
            "User-Agent": "WinGaming-Cloud-App"
        }
        r = requests.get(url_git, headers=h, timeout=5)
        if r.status_code == 200:
            return pd.read_excel(io.BytesIO(r.content))
    except Exception:
        pass
    return None

def carica_database_locale():
    df_l = pd.read_excel(FILE_LOCALI).fillna("") if os.path.exists(FILE_LOCALI) else pd.DataFrame(columns=["CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO"])
    df_t = pd.read_excel(FILE_TECNICI).fillna("") if os.path.exists(FILE_TECNICI) else pd.DataFrame(columns=["NOME", "EMAIL", "PASSWORD"])
    
    # Se abbiamo appena fatto una cancellazione o modifica, impedisce a GitHub di sovrascrivere la RAM con i dati vecchi di rete
    if st.session_state.get("congelamento_sincro_attivo", False):
        if "storico_cloud" in st.session_state:
            df_s = pd.DataFrame(st.session_state.storico_cloud)
        else:
            df_s = pd.DataFrame(columns=COLONNE_REALI_UFFICIO)
    else:
        df_s = scarica_file_da_github_se_esiste(FILE_STORICO_PERMANENTE)
        if df_s is None or df_s.empty:
            if os.path.exists(FILE_STORICO_PERMANENTE):
                df_s = pd.read_excel(FILE_STORICO_PERMANENTE).fillna("")
            else:
                df_s = pd.DataFrame(columns=COLONNE_REALI_UFFICIO)
            
    df_s = df_s.reindex(columns=COLONNE_REALI_UFFICIO).fillna("")
    
    # Pulizia automatica rigida delle ferie già trascorse rispetto ad oggi
    righe_valide = []
    oggi_ora = datetime.now()
    file_modificato_pulizia = False
    
    for _, row in df_s.iterrows():
        testo_fine = str(row.get("FINE_FERIE", "")).strip()
        if testo_fine:
            try:
                data_fine_valida = datetime.strptime(testo_fine, "%d-%m-%Y %H:%M")
                if data_fine_valida < oggi_ora:
                    file_modificato_pulizia = True
                    continue
            except Exception:
                try:
                    data_fine_valida = datetime.strptime(testo_fine.split(" "), "%d-%m-%Y")
                    if data_fine_valida.date() < oggi_ora.date():
                        file_modificato_pulizia = True
                        continue
                except Exception: pass
        righe_valide.append(row)
        
    if righe_valide:
        df_s = pd.DataFrame(righe_valide)
    else:
        df_s = pd.DataFrame(columns=COLONNE_REALI_UFFICIO)
        
    df_s = df_s.reindex(columns=COLONNE_REALI_UFFICIO).fillna("")
    
    if file_modificato_pulizia and not st.session_state.get("congelamento_sincro_attivo", False):
        push_excel_su_github(df_s)
        
    return df_l, df_t, df_s

df_locali, df_tecnici, df_storico_file = carica_database_locale()

if "storico_cloud" not in st.session_state:
    st.session_state.storico_cloud = df_storico_file.to_dict('records')

def push_excel_su_github(df_da_salvare):
    try:
        t_git = str(st.secrets["github"]["token_accesso"]).strip()
        indirizzo_base = "https://github.com"
        url_git = indirizzo_base + "/" + str(FILE_STORICO_PERMANENTE)
        
        df_pulito_salva = df_da_salvare.reindex(columns=COLONNE_REALI_UFFICIO).astype(str).fillna("")
        
        output_binario = io.BytesIO()
        with pd.ExcelWriter(output_binario, engine='openpyxl') as writer:
            df_pulito_salva.to_excel(writer, index=False)
        dati_base64 = base64.b64encode(output_binario.getvalue()).decode('utf-8')
        
        headers_git = {
            "Authorization": f"token {t_git}", 
            "Accept": "application/vnd.github+json",
            "User-Agent": "WinGaming-Cloud-App"
        }
        
        res_get = requests.get(url_git, headers=headers_git, timeout=5)
        sha_file = res_get.json().get("sha", "") if res_get.status_code == 200 else ""
        
        payload_git = {"message": "🤖 [App] Sincronizzazione database ferie", "content": str(dati_base64), "branch": "main"}
        if sha_file: 
            payload_git["sha"] = str(sha_file)
            
        risposta_put = requests.put(url_git, json=payload_git, headers=headers_git, timeout=5)
        
        if risposta_put.status_code == 422:
            if "sha" in payload_git: 
                del payload_git["sha"]
            risposta_put = requests.put(url_git, json=payload_git, headers=headers_git, timeout=5)
            
        if risposta_put.status_code == 200 or risposta_put.status_code == 201:
            st.toast("✅ File Excel allineato su GitHub!", icon="💾")
            return True
        return False
    except Exception:
        return False


# =====================================================================================
# BLOCCO 3: ACCESSO UTENTI CON MEMORIZZAZIONE SESSIONE FISSA (VALIDITÀ 2 ORE)
# =====================================================================================
if "autenticato" not in st.session_state:
    st.session_state.autenticato = False

if "token_sessione" in st.query_params:
    token_salvato = str(st.query_params["token_sessione"]).strip()
    if "_" in token_salvato:
        try:
            email_t = token_salvato.split("_")[0].strip().lower()
            st.session_state.autenticato = True
            st.session_state.user_email = email_t
            ut = df_tecnici[df_tecnici["EMAIL"].astype(str).str.lower().str.strip() == email_t]
            if not ut.empty:
                st.session_state.user_nome = str(ut["NOME"].values[0]).replace("[","").replace("]","").replace("'","").strip()
        except Exception:
            pass

if not st.session_state.autenticato:
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
                nome_grezzo = str(utente_valido["NOME"].values[0]).strip()
                st.session_state.user_nome = nome_grezzo.replace("[", "").replace("]", "").replace("'", "").replace('"', "").strip()
                st.query_params["token_sessione"] = f"{input_email}_{int(time.time())}"
                st.rerun()
            else:
                st.error("❌ Credenziali errate. Riprova.")
    st.stop()

esecutore_nome = st.session_state.user_nome
esecutore_email = st.session_state.user_email

st.markdown("<h1>🛡️ SATELLITE FERIE GESTORI</h1>", unsafe_allow_html=True)
st.markdown(f"<div class='user-badge'>👤 {esecutore_nome} ({esecutore_email})</div>", unsafe_allow_html=True)



# =====================================================================================
# BLOCCO 4: MOTORE NOTIFICA EMAIL SMTP GOOGLE CON CONVERSIONE ROTTA IP RIGIDA
# AGGIRA MANUALMENTE I BLACKOUT DELLE RETI PROTETTE DEI SERVER CLOUD DI STREAMLIT
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
            
        corpo = f"Nuova chiusura ferie registrata nel sistema WinGaming.\n\nDettagli dell'inserimento:\n--------------------------------------------------\n👤 Tecnico Esecutore: {esecutore}\n📍 Locale Coinvolto:  {locale}\n🏢 Concessionario/i:{linee_concessionari}\n📅 Inizio Chiusura:   {chiusura}\n🚚 Data Riapertura:   {riapertura}\n--------------------------------------------------\n\nWINGAMING SRL"
        msg.attach(MIMEText(corpo, 'plain'))
        
        server = smtplib.SMTP_SSL('64.233.184.108', 465, timeout=10)
        server.login(EMAIL_MITTENTE_GMAIL, pass_gmail)
        server.sendmail(EMAIL_MITTENTE_GMAIL, lista_m, msg.as_string())
        server.quit()
        return True, "OK"
    except Exception as e:
        return False, str(e)

# =====================================================================================
# BLOCCO 5: MASCHERA DI INSERIMENTO DATI (FORM ACQUISIZIONE SMARTPHONE)
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
# BLOCCO 6 - PARTE A: ELABORAZIONE INSERIMENTI E MOTORE EMAIL DINAMICO MODIFICHE
# =====================================================================================
if submit_button:
    if scelta_pvd == "- Selezionare il Locale -":
        st.error("Errore: Seleziona un locale valido.")
    elif datetime.combine(data_riapertura, ora_riapertura) <= datetime.combine(data_chiusura, ora_chiusura):
        st.error("Errore: La data di riapertura deve essere successiva alla chiusura.")
    else:
        str_c, str_r = f"{data_chiusura.strftime('%d-%m-%Y')} {ora_chiusura.strftime('%H:%M')}", f"{data_riapertura.strftime('%d-%m-%Y')} {ora_riapertura.strftime('%H:%M')}"
        testo_pvd = str(scelta_pvd)
        
        sovrapposizione_rilevata, riga_conflitto_idx, dettagli_conflitto = False, None, ""
        data_inizio_nuova = datetime.combine(data_chiusura, ora_chiusura)
        data_fine_nuova = datetime.combine(data_riapertura, ora_riapertura)
        
        for idx, row in enumerate(st.session_state.storico_cloud):
            if str(row.get("NOME_LOCALE", "")).strip() in testo_pvd or testo_pvd.strip() in str(row.get("NOME_LOCALE", "")):
                try:
                    old_i = datetime.strptime(str(row.get("INIZIO_FERIE", "")).strip(), "%d-%m-%Y %H:%M")
                    old_f = datetime.strptime(str(row.get("FINE_FERIE", "")).strip(), "%d-%m-%Y %H:%M")
                    if (data_inizio_nuova <= old_f) and (data_fine_nuova >= old_i):
                        sovrapposizione_rilevata, riga_conflitto_idx = True, idx
                        dettagli_conflitto = f"Dal {row.get('INIZIO_FERIE','')} al {row.get('FINE_FERIE','')}"
                        break
                except Exception:
                    try:
                        old_i = datetime.strptime(str(row.get("INIZIO_FERIE", "")).split(" ").strip(), "%d-%m-%Y")
                        old_f = datetime.strptime(str(row.get("FINE_FERIE", "")).split(" ").strip(), "%d-%m-%Y")
                        if (data_chiusura <= old_f.date()) and (data_riapertura >= old_i.date()):
                            sovrapposizione_rilevata, riga_conflitto_idx = True, idx
                            dettagli_conflitto = f"Dal {row.get('INIZIO_FERIE','')} al {row.get('FINE_FERIE','')}"
                            break
                    except Exception: continue

        if './' in str(scelta_pvd) or '/' in str(scelta_pvd):
            st.error("Rilevato elemento non conforme nella stringa di testo.")
        elif sovrapposizione_rilevata and not forza_sovrascrittura:
            st.error(f"⚠️ ATTENZIONE: Questo locale risulta già inserito nel periodo richiesto!\n\n📌 **Periodo registrato:** {dettagli_conflitto}.\n\nSe si tratta di una modifica spunta la casella in fondo e reinvia.")
        else:
            codice_estratto = ""
            nome_puro_locale = ""
            concessionario_estratto = ""
            
            if " - " in testo_pvd:
                parti_t = testo_pvd.split(" - ")
                codice_estratto = str(parti_t[0]).strip()
                resto_s = str(parti_t[1]).strip() if len(parti_t) > 1 else testo_pvd
            else:
                resto_s = testo_pvd.strip()
                
            if " (" in resto_s:
                parti_p = resto_s.split(" (")
                nome_puro_locale = str(parti_p[0]).strip()
                concessionario_estratto = str(parti_p[1]).replace(")", "").strip() if len(parti_p) > 1 else ""
            else:
                nome_puro_locale = resto_s
                concessionario_estratto = mappa_concessionari.get(testo_pvd, "")

            nuova = {
                "DATA_INSERIMENTO": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
                "TECNICO_INSERIMENTO": str(esecutore_nome),
                "CODICE_LOCALE": str(codice_estratto),
                "NOME_LOCALE": str(nome_puro_locale),
                "CONCESSIONARIO": str(concessionario_estratto),
                "INIZIO_FERIE": str(str_c),
                "FINE_FERIE": str(str_r),
                "PROMEMORIA_IN_COPIA": str(co_destinatario),
                "STATO_INVIO": "In attesa"
            }
            
            lista_m = [EMAIL_MANUELA_RICEVENTE, esecutore_email]
            if co_destinatario != "Nessun collega" and " (" in str(co_destinatario):
                try: lista_m.append(co_destinatario.split(" (")[-1].replace(")", "").strip())
                except Exception: pass
            
            titolo_azione = "Modifica Chiusura" if (sovrapposizione_rilevata and forza_sovrascrittura) else "Registrazione Chiusura"
            
            with st.spinner("Salvataggio e invio notifica..."):
                try:
                    pass_gmail = str(st.secrets["gmail"]["password_applicativa"]).strip()
                    msg = MIMEMultipart()
                    msg['From'] = EMAIL_MITTENTE_GMAIL
                    msg['To'] = ", ".join(lista_m)
                    msg['Subject'] = f"🛡️ {titolo_azione} Ferie - {nome_puro_locale}"
                    
                    linee_concessionari = f" {concessionario_estratto}"
                    corpo = f"Rilevato aggiornamento chiusura ferie nel sistema WinGaming.\n\nDettagli della pratica:\n--------------------------------------------------\n🔔 Stato Operazione:  {titolo_azione.upper()}\n👤 Tecnico Esecutore: {esecutore_nome}\n📍 Locale Coinvolto:  {nome_puro_locale}\n🏢 Concessionario/i:{linee_concessionari}\n📅 Inizio Chiusura:   {str_c}\n🚚 Data Riapertura:   {str_r}\n--------------------------------------------------\n\nWINGAMING SRL"
                    msg.attach(MIMEText(corpo, 'plain'))
                    
                    server = smtplib.SMTP_SSL('64.233.184.108', 465, timeout=10)
                    server.login(EMAIL_MITTENTE_GMAIL, pass_gmail)
                    server.sendmail(EMAIL_MITTENTE_GMAIL, lista_m, msg.as_string())
                    server.quit()
                    invio_ok, risposta_server = True, "OK"
                except Exception as e_mail:
                    invio_ok, risposta_server = False, str(e_mail)
            
            if invio_ok:
                nuova["STATO_INVIO"] = "Inviato OK"
                if sovrapposizione_rilevata and riga_conflitto_idx is not None:
                    st.session_state.storico_cloud.pop(riga_conflitto_idx)
                
                st.session_state.storico_cloud.append(nuova)
                df_salva = pd.DataFrame(st.session_state.storico_cloud)
                df_salva.to_excel(FILE_STORICO_PERMANENTE, index=False)
                push_excel_su_github(df_salva)
                
                st.success("✅ OPERAZIONE COMPLETATA!\n\n📧 Registro allineato su GitHub e e-mail inviata.")
                st.session_state.form_id += 1
                time.sleep(0.5)
                st.rerun()
            else:
                st.error(f"❌ Errore Google SMTP: {risposta_server}. Spedizione e-mail fallita.")
# =====================================================================================
# BLOCCO 6 - PARTE B: PROMEMORIA LOGISTICI 3 GG E PLANCIA DI VISUALIZZAZIONE ADMIN
# =====================================================================================
st.markdown("---")
st.markdown("### 📅 Promemoria Giri Logistici (Preavviso 3 Giorni)")
oggi = datetime.now().date()
alert_c, alert_r = [], []
for row in st.session_state.storico_cloud:
    try:
        d_i = datetime.strptime(str(row.get("INIZIO_FERIE", "")).strip().split(" ")[0], "%d-%m-%Y").date()
        d_f = datetime.strptime(str(row.get("FINE_FERIE", "")).strip().split(" ")[0], "%d-%m-%Y").date()
        if d_i - oggi == timedelta(days=3): alert_c.append(f"⚠️ **{row.get('NOME_LOCALE', 'Locale')}** chiude tra 3 giorni")
        if d_f - oggi == timedelta(days=3): alert_r.append(f"🚚 **{row.get('NOME_LOCALE', 'Locale')}** riapre tra 3 giorni")
    except Exception: continue
for a in alert_c: st.error(a)
for r in alert_r: st.warning(r)

if esecutore_email.lower() == EMAIL_MANUELA_RICEVENTE.lower():
    st.markdown("<br>### 📊 Registro Storico Chiusure Centralizzato", unsafe_allow_html=True)
    colonne_reali_ufficio = ["DATA_INSERIMENTO", "TECNICO_INSERIMENTO", "CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO", "INIZIO_FERIE", "FINE_FERIE", "PROMEMORIA_IN_COPIA", "STATO_INVIO"]
    
    if st.session_state.storico_cloud:
        df_vis = pd.DataFrame(st.session_state.storico_cloud)
        df_vis = df_vis.reindex(columns=colonne_reali_ufficio).fillna("")
        st.dataframe(df_vis, hide_index=True)
        
        with io.BytesIO() as buffer:
            df_vis.to_excel(buffer, index=False)
            st.download_button(label="📥 Scarica Registro Excel Storico", data=buffer.getvalue(), file_name="storico_ferie.xlsx", mime="application/vnd.ms-excel")
    else:
        st.info("📭 Nessuna chiusura presente in memoria. Trascina il file Excel storico in fondo per ripopolare la plancia.")
        
    st.markdown("---")
    st.markdown("### 🏢 Locali SNAITECH da inserire a sistema")
    righe_snaitech = [row for row in st.session_state.storico_cloud if "snai" in (str(row.get("CONCESSIONARIO", "")) + " " + str(row.get("NOME_LOCALE", ""))).lower()] if st.session_state.storico_cloud else []
    if righe_snaitech:
        df_snai = pd.DataFrame(righe_snaitech).reindex(columns=colonne_reali_ufficio).fillna("")
        st.dataframe(df_snai[["CODICE_LOCALE", "NOME_LOCALE", "INIZIO_FERIE", "FINE_FERIE", "TECNICO_INSERIMENTO"]], hide_index=True)
    else:
        st.write("✅ Nessuna chiusura attiva per locali Snaitech.")

# =====================================================================================
# BLOCCO 6 - PARTE C: TASTO CANCELLAZIONE EXCEL REALE, CARICATORE UFFICIO E LOGOUT
# =====================================================================================
    st.markdown("---")
    st.markdown("### 🗑️ Cancella un Periodo Registrato")
    opzioni_cancellazione = ["- Seleziona la riga da eliminare -"]
    if st.session_state.storico_cloud:
        for idx, row in enumerate(st.session_state.storico_cloud):
            opzioni_cancellazione.append(f"ID {idx} | {row.get('CODICE_LOCALE', '')} - {row.get('NOME_LOCALE', '')} (Dal {row.get('INIZIO_FERIE', '')})")
            
    selezione_delete = st.selectbox("Scegli la chiusura da eliminare dal database:", opzioni_cancellazione, disabled=not st.session_state.storico_cloud)
    if selezione_delete != "- Seleziona la riga da eliminare -" and st.session_state.storico_cloud:
        try:
            parti_s = selezione_delete.split("ID ")
            if len(parti_s) > 1:
                sub_stringa = parti_s[1]
                idx_iscolato_str = sub_stringa.split(" |")[0]
                idx_da_eliminare = int(idx_iscolato_str)
                
                if st.button("❌ ELIMINA DEFINITIVAMENTE QUESTA CHIUSURA"):
                    st.session_state.congelamento_sincro_attivo = True  # 🛡️ COSTRUTTORE DI PROTEZIONE RAM CANCELLAZIONI
                    st.session_state.storico_cloud.pop(idx_da_eliminare)
                    df_nuovo_salva = pd.DataFrame(st.session_state.storico_cloud)
                    df_nuovo_salva.to_excel(FILE_STORICO_PERMANENTE, index=False)
                    push_excel_su_github(df_nuovo_salva)
                    st.success("🗑️ Chiusura rimossa con successo!")
                    time.sleep(0.5)
                    st.rerun()
        except Exception: pass
        
    st.markdown("---")
    st.markdown("### 📤 Ricarica Registro Excel Aggiornato dall'Ufficio")
    file_caricato = st.file_uploader("Trascina il file storico_ferie.xlsx modificato per caricare i dati nel portale:", type=["xlsx"])
    if file_caricato is not None:
        try:
            df_caricato = pd.read_excel(file_caricato).fillna("")
            if "CODICE_LOCALE" in df_caricato.columns:
                if st.button("🔄 CONFERMA E SOVRASCRIVI DATABASE CON QUESTO FILE"):
                    st.session_state.storico_cloud = df_caricato.to_dict('records')
                    df_caricato.to_excel(FILE_STORICO_PERMANENTE, index=False)
                    push_excel_su_github(df_caricato)
                    st.success("✅ Database popolato e sincronizzato con successo su GitHub!")
                    time.sleep(1.5)
                    st.rerun()
            else:
                st.error("❌ Struttura file non valida. Controlla che i nomi delle colonne siano in orizzontale.")
        except Exception as e_load: st.error(f"❌ Errore lettura: {str(e_load)}")

# PULSANTE LOGOUT PRINCIPALE STRUTTURALE MARGINE ZERO
st.markdown("<br>", unsafe_allow_html=True)
col_out1, col_out2, col_out3 = st.columns([1, 2, 1])
with col_out2:
    if st.button("🚪 DISCONNETTI ACCOUNT / LOGOUT"):
        st.query_params.clear()
        st.session_state.autenticato = False
        st.success("Uscita effettuata con successo!")
        time.sleep(0.5)
        st.rerun()

