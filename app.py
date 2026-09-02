# =====================================================================================
# SW GENERALE DI INSERIMENTO FERIE GESTORI WIN GAMING — PRODUZIONE INTEGRALE FINALE
# BLOCCO 1: STRUTTURA DI BASE, LIBRERIE E PERSONALIZZAZIONE INTERFACCIA UTENTE
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
from datetime import datetime, timedelta, time as dtime

st.set_page_config(
    page_title="Ferie Gestori", 
    page_icon="📅", 
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    #MainMenu, footer, header, .stDecoration, [data-testid="stHeader"], [data-testid="stFooter"] {
        visibility: hidden !important; display: none !important;
    }
    .stApp { background-color: #f8fafc !important; color: #1e293b !important; font-family: 'Segoe UI', sans-serif; }
    h1 { color: #115e59 !important; font-size: 28px !important; text-align: center; font-weight: 800 !important; margin-bottom: 25px; }
    .stMarkdown h3, label, p, [data-testid="stWidgetLabel"] p, .stSelectbox label { color: #1e293b !important; font-weight: 800 !important; font-size: 16px !important; opacity: 1 !important; }
    div[data-testid="stForm"] { background-color: #ffffff !important; border: 2px solid #94a3b8 !important; border-radius: 14px !important; padding: 25px !important; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); }
    .stButton>button { background: linear-gradient(135deg, #0f766e 0%, #115e59 100%) !important; color: #ffffff !important; font-weight: 800 !important; font-size: 17px !important; width: 100%; border-radius: 10px !important; height: 54px !important; border: none !important; box-shadow: 0 4px 14px rgba(17, 94, 89, 0.3); }
    .user-badge { background-color: #ffffff; padding: 14px; border-radius: 10px; border: 2px solid #115e59; margin-bottom: 30px; text-align: center; color: #115e59 !important; font-weight: 800; font-size: 16px; }
    </style>
""", unsafe_allow_html=True)
# =====================================================================================
# BLOCCO 2: COLLEGAMENTO FILE EXCEL PERMANENTI E FUNZIONI DI SCRITTURA SU GITHUB
# VERSIONE DI PRODUZIONE STRUTTURATA 100% PER EXCEL (.XLSX) — EXCEL WRITER CLOSED FIXED
# =====================================================================================
FILE_LOCALI = "elenco_locali.xlsx"
FILE_TECNICI = "elenco_tecnici.xlsx"
FILE_STORICO_PERMANENTE = "storico_ferie.xlsx"

EMAIL_MITTENTE_GMAIL = "wingamingsrl@gmail.com"
EMAIL_MANUELA_RICEVENTE = "manuela.arigoni@wingaming.it"

def scarica_file_da_github_se_esiste(nome_file):
    try:
        t_git = str(st.secrets["github"]["token_accesso"]).strip()
        parte1 = "https://github.com"
        parte2 = "repos/wingamingsrl/sistema-ferie/contents"
        url_git = parte1 + "/" + parte2 + "/" + nome_file + "?t=" + str(int(time.time()))
        
        h = {
            "Authorization": "Bearer " + t_git, 
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "WinGaming-Cloud-App"
        }
        r = requests.get(url_git, headers=h, timeout=5)
        if r.status_code == 200:
            b64_content = r.json().get("content", "")
            return pd.read_excel(io.BytesIO(base64.b64decode(b64_content)))
    except Exception:
        pass
    return None

def carica_database_locale():
    df_l = pd.read_excel(FILE_LOCALI).fillna("") if os.path.exists(FILE_LOCALI) else pd.DataFrame(columns=["CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO"])
    df_t = pd.read_excel(FILE_TECNICI).fillna("") if os.path.exists(FILE_TECNICI) else pd.DataFrame(columns=["NOME", "EMAIL", "PASSWORD", "RUOLO"])
    
    df_s = scarica_file_da_github_se_esiste(FILE_STORICO_PERMANENTE)
    if df_s is None or df_s.empty:
        df_s = pd.DataFrame(columns=["DATA_INSERIMENTO", "TECNICO_INSERIMENTO", "CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO", "INIZIO_FERIE", "FINE_FERIE", "PROMEMORIA_IN_COPIA", "STATO_INVIO"])
    return df_l, df_t, df_s.fillna("")

df_locali, df_tecnici, df_storico_file = carica_database_locale()
st.session_state.storico_cloud = df_storico_file.to_dict('records')

def push_excel_su_github(df_da_salvare):
    try:
        t_git = str(st.secrets["github"]["token_accesso"]).strip()
        parte1 = "https://github.com"
        parte2 = "repos/wingamingsrl/sistema-ferie/contents"
        url_git = parte1 + "/" + parte2 + "/" + FILE_STORICO_PERMANENTE
        
        if df_da_salvare.empty:
            df_da_salvare = pd.DataFrame(columns=["DATA_INSERIMENTO", "TECNICO_INSERIMENTO", "CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO", "INIZIO_FERIE", "FINE_FERIE", "PROMEMORIA_IN_COPIA", "STATO_INVIO"])
        
        output_binario = io.BytesIO()
        with pd.ExcelWriter(output_binario, engine='openpyxl') as writer:
            df_da_salvare.to_excel(writer, index=False)
        dati_base64 = base64.b64encode(output_binario.getvalue()).decode('utf-8')
        
        headers_git = {
            "Authorization": "Bearer " + t_git, 
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "WinGaming-Cloud-App"
        }
        
        res_get = requests.get(url_git, headers=headers_git, params={"ref": "main"}, timeout=5)
        sha_file = res_get.json().get("sha", "") if res_get.status_code == 200 else ""
        
        payload_git = {
            "message": "🤖 [App] Scrittura database Excel", 
            "content": dati_base64,
            "branch": "main"
        }
        
        if sha_file: 
            payload_git["sha"] = sha_file
                        
        risposta_put = requests.put(url_git, json=payload_git, headers=headers_git, timeout=5)
        
        # 🛡️ DISINNESCORO TOTALE ERRORE 422: Se GitHub fa ostruzionismo e rifiuta la sovrascrittura,
        # applichiamo la forza bruta: eliminiamo il file vecchio e lo ricreiamo da zero in un millisecondo!
        if risposta_put.status_code == 422 and sha_file:
            payload_delete = {
                "message": "🧹 [Reset] Rimozione file per sblocco errore 422",
                "sha": sha_file,
                "branch": "main"
            }
            # 1. Rimuove il file bloccato
            requests.delete(url_git, json=payload_delete, headers=headers_git, timeout=5)
            # 2. Ricrea istantaneamente il file Excel sano, pulito e aggiornato
            if "sha" in payload_git: del payload_git["sha"]
            risposta_put = requests.put(url_git, json=payload_git, headers=headers_git, timeout=5)
        
        if risposta_put.status_code == 200 or risposta_put.status_code == 201:
            st.toast("✅ Excel salvato su GitHub!", icon="💾")
            return True
        else:
            st.error(f"❌ Rifiuto Scrittura GitHub. Stato: {risposta_put.status_code}")
            return False
    except Exception as e_err:
        st.error(f"💥 Errore Interno: {str(e_err)}")
        return False



# =====================================================================================
# BLOCCO 3: AUTENTICAZIONE E GESTIONE CREDENZIALI DINAMICHE DA EXCEL (RUOLI)
# VERSIONE DI PRODUZIONE CORRETTA — ESTRAZIONE VALORI PULITI DA PANDAS SERIES
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
                # 🛡️ FIX CHIRURGICO: Prende il primo valore pulito per eliminare le parentesi e la scritta dtype/length
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
# BLOCCO 4: NOTIFICA EMAIL SMTP GOOGLE CON ELENCO CONCESSIONARI INCOLONNATO
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
# SPEZZONE CORRETTO BLOCCO 6: ESTRAZIONE DATI GEOMETRICA SENZA STRIP SU LISTE
# =====================================================================================
        else:
            str_c = f"{data_chiusura.strftime('%d-%m-%Y')} {ora_chiusura.strftime('%H:%M')}"
            str_r = f"{data_riapertura.strftime('%d-%m-%Y')} {ora_riapertura.strftime('%H:%M')}"
            
            testo_pvd = str(scelta_pvd)
            
            # 🛡️ ESTRAZIONE SICURA: Separiamo prima il testo e poi puliamo i singoli elementi
            if " - " in testo_pvd:
                parti_pvd = testo_pvd.split(" - ")
                codice_estratto = str(parti_pvd[0]).strip()
                resto_nome = str(parti_pvd[1])
            else:
                codice_estratto = testo_pvd.strip()
                resto_nome = testo_pvd
                
            if " (" in resto_nome:
                parti_nome = resto_nome.split(" (")
                nome_puro_locale = str(parti_nome[0]).strip()
            else:
                nome_puro_locale = resto_nome.strip()
            
            concessionario_estratto = mappa_concessionari.get(scelta_pvd, "")
            
            # Coerenza assoluta a 9 colonne con celle di solo testo
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
            
            chiave_pulita = nome_puro_locale
            lista_m = [EMAIL_MANUELA_RICEVENTE, esecutore_email]
            if co_destinatario != "Nessun collega":
                mail_pulita = co_destinatario.split(" (")[-1].replace(")", "").strip() if " (" in co_destinatario else co_destinatario.strip()
                lista_m.append(mail_pulita)
                
            with st.spinner("Salvataggio e invio notifica..."):
                invio_ok, risposta_server = invia_mail_diretta_smtp(lista_m, chiave_pulita, concessionario_estratto, str_c, str_r, esecutore_nome)
            
            if invio_ok:
                nuova["STATO_INVIO"] = "Inviato OK"
                if sovrapposizione_rilevata and riga_conflitto_idx is not None:
                    st.session_state.storico_cloud.pop(riga_conflitto_idx)
                st.session_state.storico_cloud.append(nuova)
                
                df_salva = pd.DataFrame(st.session_state.storico_cloud)
                
                # Generazione dell'Excel pulito localmente
                with pd.ExcelWriter(FILE_STORICO_PERMANENTE, engine='openpyxl') as writer:
                    df_salva.to_excel(writer, index=False)
                    
                # Chiamata alla funzione del Blocco 2
                push_excel_su_github(df_salva)
                
                st.success("✅ OPERAZIONE COMPLETATA!\n\n📧 Registro allineato su GitHub e e-mail inviata.")
                st.session_state.form_id += 1
                time.sleep(2)
                st.rerun()
            else:
                st.error(f"❌ Errore Google SMTP: {risposta_server}. Spedizione e-mail fallita.")

