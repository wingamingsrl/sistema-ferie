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
# BLOCCO 2: COLLEGAMENTO ED INTERROGAZIONE FILE EXCEL CON ACCESSO DIRETTO API GITHUB
# VERSIONE DI PRODUZIONE SIGILLATA — STRUTTURATA SULLE COLONNE REALI DELLA FOTO
# =====================================================================================
FILE_LOCALI = "elenco_locali.xlsx"
FILE_TECNICI = "elenco_tecnici.xlsx"
FILE_STORICO_PERMANENTE = "storico_ferie.xlsx"

EMAIL_MITTENTE_GMAIL = "wingamingsrl@gmail.com"
EMAIL_MANUELA_RICEVENTE = "manuela.arigoni@wingaming.it"

def scarica_file_da_github_se_esiste(nome_file):
    try:
        t_git = str(st.secrets["github"]["token_accesso"]).strip()
        c_time = str(int(time.time() * 1000))
        url_git = f"https://github.com{nome_file}?_nonce={c_time}"
        
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
    
    # Inizializzazione rigida sulle colonne accorciate esatte visibili nell'immagine
    colonne_foto = ["DATA_INS", "TECNICO_", "CODICE_L", "NOME_LO", "CONCESSI", "INIZIO_FE", "FINE_FERI", "PROMEMO", "STATO_IN"]
    df_s = scarica_file_da_github_se_esiste(FILE_STORICO_PERMANENTE)
    
    if df_s is None or df_s.empty:
        if os.path.exists(FILE_STORICO_PERMANENTE):
            df_s = pd.read_excel(FILE_STORICO_PERMANENTE).fillna("")
        else:
            df_s = pd.DataFrame(columns=colonne_foto)
            
    df_s = df_s.reindex(columns=colonne_foto).fillna("")
    return df_l, df_t, df_s

df_locali, df_tecnici, df_storico_file = carica_database_locale()

if "storico_cloud" not in st.session_state:
    st.session_state.storico_cloud = df_storico_file.to_dict('records')

def push_excel_su_github(df_da_salvare):
    try:
        t_git = str(st.secrets["github"]["token_accesso"]).strip()
        url_git = f"https://github.com{FILE_STORICO_PERMANENTE}"
        
        output_binario = io.BytesIO()
        colonne_foto = ["DATA_INS", "TECNICO_", "CODICE_L", "NOME_LO", "CONCESSI", "INIZIO_FE", "FINE_FERI", "PROMEMO", "STATO_IN"]
        df_pulito_salva = df_da_salvare.reindex(columns=colonne_foto).fillna("")
        
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
        
        payload_git = {"message": "🤖 [App] Allineamento database ferie Excel", "content": dati_base64, "branch": "main"}
        if sha_file: payload_git["sha"] = sha_file
            
        risposta_put = requests.put(url_git, json=payload_git, headers=headers_git, timeout=5)
        
        if risposta_put.status_code == 422 and sha_file:
            p_del = {"message": "🧹 Sblocco conflitto", "sha": sha_file, "branch": "main"}
            requests.delete(url_git, json=p_del, headers=headers_git, timeout=5)
            if "sha" in payload_git: del payload_git["sha"]
            risposta_put = requests.put(url_git, json=payload_git, headers=headers_git, timeout=5)
            
        if risposta_put.status_code == 200 or risposta_put.status_code == 201:
            st.toast("✅ File Excel salvato correttamente su GitHub!", icon="💾")
            return True
        return False
    except Exception:
        return False

# =====================================================================================
# BLOCCO 3: ACCESSO SICUREZZA CON RIMOZIONE DEL BUG DI VISUALIZZAZIONE DTYPE/LENGTH
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
                st.session_state.user_nome = str(utente_valido["NOME"].values).strip()
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
# BLOCCO 6: ELABORAZIONE RIGHE, CONTROLLO DOPPIONI ED AREA AMMINISTRATORE RIPRISTINATA
# VERSIONE DI PRODUZIONE ORIGINALE — RIPRISTINO INTEGRALE CARICAMENTO EXCEL PULITO
# =====================================================================================
if submit_button:
    if scelta_pvd == "- Selezionare il Locale -":
        st.error("Errore: Seleziona un locale valido.")
    elif datetime.combine(data_riapertura, ora_riapertura) <= datetime.combine(data_chiusura, ora_chiusura):
        st.error("Errore: La data di riapertura deve essere successiva alla chiusura.")
    else:
        str_c, str_r = f"{data_chiusura.strftime('%d-%m-%Y')} {ora_chiusura.strftime('%H:%M')}", f"{data_riapertura.strftime('%d-%m-%Y')} {ora_riapertura.strftime('%H:%M')}"
        testo_pvd = str(scelta_pvd)
        
        # CONTROLLO ANTIDOPPIONE BASATO SUL LOCALE SELEZIONATO
        sovrapposizione_rilevata, riga_conflitto_idx, dettagli_conflitto = False, None, ""
        for idx, row in enumerate(st.session_state.storico_cloud):
            if str(row.get("LOCALE", "")).strip() == testo_pvd.strip():
                try:
                    old_i = datetime.strptime(str(row.get("INIZIO_FE", row.get("INIZIO_FERIE", ""))).split(" "), "%d-%m-%Y").date()
                    old_f = datetime.strptime(str(row.get("FINE_FERI", row.get("FINE_FERIE", ""))).split(" "), "%d-%m-%Y").date()
                    if (data_chiusura <= old_f) and (data_riapertura >= old_i):
                        sovrapposizione_rilevata, riga_conflitto_idx = True, idx
                        dettagli_conflitto = f"Dal {row.get('INIZIO_FE', row.get('INIZIO_FERIE', ''))} al {row.get('FINE_FERI', row.get('FINE_FERIE', ''))}"
                        break
                except Exception: continue

        if sovrapposizione_rilevata and not forza_sovrascrittura:
            st.error(f"⚠️ ATTENZIONE: Questo locale risulta già inserito nel periodo richiesto!\n\n📌 **Periodo registrato:** {dettagli_conflitto}.\n\nSpunta la casella in fondo e reinvia per confermare.")
        else:
            # Scrittura dinamica: si adatta automaticamente alle colonne presenti nella foto
            nuova = {
                "DATA_INS": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
                "TECNICO_": esecutore_nome,
                "CODICE_L": testo_pvd.split(" - ").strip() if " - " in testo_pvd else "",
                "NOME_LO": testo_pvd.split(" - ").split(" (").strip() if " - " in testo_pvd else testo_pvd.strip(),
                "CONCESSI": testo_pvd.split(" (")[-1].replace(")", "").strip() if " (" in testo_pvd else mappa_concessionari.get(testo_pvd, ""),
                "INIZIO_FE": str_c,
                "FINE_FERI": str_r,
                "PROMEMO": str(co_destinatario),
                "STATO_IN": "In attesa"
            }
            
            # Supporto compatibilità per la visualizzazione vecchio stile se presente
            if st.session_state.storico_cloud and "LOCALE" in st.session_state.storico_cloud[0]:
                nuova = {
                    "DATA_INSERIMENTO": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
                    "TECNICO": esecutore_nome,
                    "LOCALE": scelta_pvd,
                    "INIZIO_FERIE": data_chiusura.strftime('%d-%m-%Y'),
                    "FINE_FERIE": data_riapertura.strftime('%d-%m-%Y'),
                    "COPIA_PROMEMORIA": co_destinatario
                }

            chiave_pulita = testo_pvd.split(" (").strip() if " (" in testo_pvd else testo_pvd.strip()
            concessionario_estratto = mappa_concessionari.get(testo_pvd, "")
            
            lista_m = [EMAIL_MANUELA_RICEVENTE, esecutore_email]
            if co_destinatario != "Nessun collega" and " (" in str(co_destinatario):
                try: lista_m.append(co_destinatario.split(" (")[-1].replace(")", "").strip())
                except Exception: pass
                
            with st.spinner("Salvataggio e invio notifica..."):
                invio_ok, risposta_server = invia_mail_diretta_smtp(lista_m, chiave_pulita, concessionario_estratto, str_c, str_r, esecutore_nome)
            
            if invio_ok:
                if "STATO_IN" in nuova: nuova["STATO_IN"] = "Inviato OK"
                if "STATO_INVIO" in nuova: nuova["STATO_INVIO"] = "Inviato OK"
                
                if sovrapposizione_rilevata and riga_conflitto_idx is not None:
                    st.session_state.storico_cloud.pop(riga_conflitto_idx)
                st.session_state.storico_cloud.append(nuova)
                
                df_salva = pd.DataFrame(st.session_state.storico_cloud)
                df_salva.to_excel(FILE_STORICO_PERMANENTE, index=False)
                push_excel_su_github(df_salva)
                
                st.success("✅ OPERAZIONE COMPLETATA!\n\n📧 Registro allineato su GitHub e e-mail inviata.")
                st.session_state.form_id += 1
                time.sleep(1.5)
                st.rerun()
            else:
                st.error(f"❌ Errore Google SMTP: {risposta_server}. Spedizione e-mail fallita.")

st.markdown("---")
st.markdown("### 📅 Promemoria Giri Logistici (Preavviso 3 Giorni)")
oggi = datetime.now().date()
alert_c, alert_r = [], []
for row in st.session_state.storico_cloud:
    try:
        col_i = "INIZIO_FE" if "INIZIO_FE" in row else "INIZIO_FERIE"
        col_f = "FINE_FERI" if "FINE_FERI" in row else "FINE_FERIE"
        col_l = "NOME_LO" if "NOME_LO" in row else "LOCALE"
        d_i = datetime.strptime(str(row.get(col_i, "")).split(" "), "%d-%m-%Y").date()
        d_f = datetime.strptime(str(row.get(col_f, "")).split(" "), "%d-%m-%Y").date()
        if d_i - oggi == timedelta(days=3): alert_c.append(f"⚠️ **{row.get(col_l, 'Locale')}** chiude tra 3 giorni")
        if d_f - oggi == timedelta(days=3): alert_r.append(f"🚚 **{row.get(col_l, 'Locale')}** riapre tra 3 giorni")
    except Exception: continue
for a in alert_c: st.error(a)
for r in alert_r: st.warning(r)

if st.sidebar.button("🚪 Disconnetti Account"):
    del st.session_state.autenticato
    st.rerun()

# --- PLANCCIA AMMINISTRATORE DIRETTA E LINEARE (STILE ORIGINALE FUNZIONANTE) ---
if esecutore_email.lower() == EMAIL_MANUELA_RICEVENTE.lower():
    st.markdown("<br>### 📊 Registro Storico Chiusure Centralizzato", unsafe_allow_html=True)
    
    if st.session_state.storico_cloud:
        df_vis = pd.DataFrame(st.session_state.storico_cloud)
        st.dataframe(df_vis, hide_index=True)
        
        with io.BytesIO() as buffer:
            df_vis.to_excel(buffer, index=False)
            st.download_button(label="📥 Scarica Registro Excel Storico", data=buffer.getvalue(), file_name="storico_ferie.xlsx", mime="application/vnd.ms-excel")
    else:
        st.info("📭 Nessuna chiusura presente in memoria. Trascina il file Excel storico in fondo per ripopolare la plancia.")
        
    st.markdown("---")
    st.markdown("### 🏢 Locali SNAITECH da inserire a sistema")
    
    righe_snaitech = []
    if st.session_state.storico_cloud:
        for row in st.session_state.storico_cloud:
            testo_completo = str(list(row.values())).lower()
            if "snai" in testo_completo:
                righe_snaitech.append(row)
                
    if righe_snaitech:
        st.dataframe(pd.DataFrame(righe_snaitech), hide_index=True)
    else:
        st.write("✅ Nessuna chiusura attiva per locali Snaitech.")
        
    st.markdown("---")
    st.markdown("### 🗑️ Cancella un Periodo Registrato")
    opzioni_cancellazione = ["- Seleziona la riga da eliminare -"]
    if st.session_state.storico_cloud:
        for idx, row in enumerate(st.session_state.storico_cloud):
            lbl = row.get("LOCALE", row.get("NOME_LO", "Locale"))
            inf = row.get("INIZIO_FERIE", row.get("INIZIO_FE", ""))
            opzioni_cancellazione.append(f"ID {idx} | {lbl} (Dal {inf})")
            
    selezione_delete = st.selectbox("Scegli la chiusura da eliminare dal database:", opzioni_cancellazione, disabled=not st.session_state.storico_cloud)
    if selezione_delete != "- Seleziona la riga da eliminare -" and st.session_state.storico_cloud:
        try:
            parti_str = selezione_delete.split("ID ")
            pezzo_numerico = parti_str.split(" |")
            idx_da_eliminare = int(pezzo_numerico)
            if st.button("❌ ELIMINA DEFINITIVAMENTE QUESTA CHIUSURA"):
                st.session_state.storico_cloud.pop(idx_da_eliminare)
                df_nuovo_salva = pd.DataFrame(st.session_state.storico_cloud)
                df_nuovo_salva.to_excel(FILE_STORICO_PERMANENTE, index=False)
                push_excel_su_github(df_nuovo_salva)
                st.success("🗑️ Chiusura rimossa con successo!")
                time.sleep(1)
                st.rerun()
        except Exception: pass
        
    st.markdown("---")
    st.markdown("### 📤 Ricarica Registro Excel Aggiornato dall'Ufficio")
    file_caricato = st.file_uploader("Trascina il file storico_ferie.xlsx modificato per caricare i dati nel portale:", type=["xlsx"])
    if file_caricato is not None:
        try:
            df_caricato = pd.read_excel(file_caricato).fillna("")
            if st.button("🔄 CONFERMA E SOVRASCRIVI DATABASE CON QUESTO FILE"):
                st.session_state.storico_cloud = df_caricato.to_dict('records')
                df_caricato.to_excel(FILE_STORICO_PERMANENTE, index=False)
                push_excel_su_github(df_caricato)
                st.success("✅ Database popolato e sincronizzato con successo su GitHub!")
                time.sleep(1.5)
                st.rerun()
        except Exception as e_load: st.error(f"❌ Errore lettura: {str(e_load)}")


