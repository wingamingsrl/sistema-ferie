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
import pyotp  # 🛡️ INTEGRATO PER LA GENERAZIONE AUTOMATICA OTP 2FA
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime, timedelta, time as dtime

icona_app = "logo.png" if os.path.exists("logo.png") else "📅"

st.set_page_config(
    page_title="Ferie Gestori - Sandbox Test", 
    page_icon=icona_app, 
    layout="centered",
    initial_sidebar_state="collapsed"
)

# =====================================================================================
# 🛡️ BARRIERA ANTISOPRAVVOSCRIZIONE DI MANUELA: BLOCCA LO SCHERMO SE IL ROBOT STA GIRANDO
# =====================================================================================
if st.session_state.get("congelamento_sincro_attivo", False):
    st.error("🚨 ATTENZIONE: Sincronizzazione forzata in corso su GitHub!")
    st.info("⏳ Per evitare la perdita di dati, la plancia è temporaneamente CONGELATA. Attendi circa 2 minuti che il robot completi gli aggiornamenti e poi rinfresca la pagina.")
    st.spinner("Allineamento database online in corso...")
    if st.button("🔄 VERIFICA SE IL ROBOT HA FINITO (RINFRESCA)"):
        st.session_state.congelamento_sincro_attivo = False
        st.rerun()
    st.stop() # 💥 GHIGLIOTTINA: Blocca lo smartphone qui se il processo è attivo!
# =====================================================================================

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

COLONNE_REALI_UFFICIO = ["DATA_INSERIMENTO", "TECNICO_INSERIMENTO", "CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO", "INIZIO_FERIE", "FINE_FERIE", "PROMEMORIA_IN_COPIA", "STATO_INVIO", "ROBOT_ACTION"]


# =====================================================================================
# 🛡️ PROTEZIONE DI SICUREZZA DI MANUELA: TIMEOUT CON RESET TOTALE ANTI-LOOP LOGOUT
# =====================================================================================
import time as t_lib

# ⏱️ CONFIGURAZIONE UFFICIALE: 7200 secondi corrispondono a 2 ore esatte di autonomia.
# (Mantieni 60 per il tuo test di 1 minuto, poi rimetterai 7200 per i tecnici dell'ufficio!)
SECONDI_MASSIMI_SESSIONE = 7200

if "ora_creazione_sessione" not in st.session_state:
    st.session_state.ora_creazione_sessione = t_lib.time()

# Calcola i secondi passati dal login iniziale
secondi_correnti = t_lib.time()
tempo_trascorso = secondi_correnti - st.session_state.ora_creazione_sessione

if "user_nome" in st.session_state and st.session_state.user_nome:
    if tempo_trascorso > SECONDI_MASSIMI_SESSIONE:
        # 💥 GHIGLIOTTINA IMMEDIATA: Svuota la RAM dello smartphone
        st.session_state.clear()
        st.session_state.autenticato = False
        
        # 🔑 CHIAVE DI VOLTA DI MANUELA: Aggiorna IMMEDIATAMENTE l'ora di creazione al momento del crash,
        # così al prossimo login il contatore ripartirà da zero senza mostrare doppi messaggi di errore!
        st.session_state.ora_creazione_sessione = t_lib.time()
        
        # Mostra l'avviso di sicurezza ed esegue il reset pulito della pagina
        st.warning("🔒 Sessione scaduta per inattività. Effettua nuovamente il login per sicurezza.")
        if "st" in locals() and hasattr(st, "query_params"):
            st.query_params.clear()
        t_lib.sleep(0.5)
        st.rerun()
# =====================================================================================




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
    df_l = pd.read_excel(FILE_LOCALI).fillna("") if os.path.exists(FILE_LOCALI) else pd.DataFrame(columns=["CODICE_LOCALE", "NOME_LOCALE", "IONARIO"])
    df_t = pd.read_excel(FILE_TECNICI).fillna("") if os.path.exists(FILE_TECNICI) else pd.DataFrame(columns=["NOME", "EMAIL", "PASSWORD"])
    
    if st.session_state.get("congelamento_sincro_attivo", False):
        if "storico_cloud" in st.session_state:
            df_s = pd.DataFrame(st.session_state.storico_cloud)
        else:
            df_s = pd.DataFrame(columns=COLONNE_REALI_UFFICIO)
    else:
        df_s = pd.read_excel(FILE_STORICO_PERMANENTE).fillna("") if os.path.exists(FILE_STORICO_PERMANENTE) else pd.DataFrame(columns=COLONNE_REALI_UFFICIO)
            
    df_s = df_s.reindex(columns=COLONNE_REALI_UFFICIO).fillna("")
    return df_l, df_t, df_s
    
    # 🧹 MOTORE AUTOMATICO GIORNALIERO (REPLICA ESATTA DEL TASTO ELIMINA MANUALE)
    # Si attiva in automatico solo se l'utente è loggato e la memoria cloud è pronta
    if "user_nome" in st.session_state and "storico_cloud" in st.session_state:
        oggi_ora = datetime.now()
        indici_da_eliminare = []
        
        # Scansiona lo storico cloud individuando la posizione esatta (ID) dei locali con ferie passate rispetto a oggi (Settembre 2026)
        for idx, row in enumerate(st.session_state.storico_cloud):
            testo_fine = str(row.get("FINE_FERIE", "")).strip()
            if testo_fine:
                try:
                    data_fine_valida = datetime.strptime(testo_fine, "%d-%m-%Y %H:%M")
                    if data_fine_valida < oggi_ora:
                        indici_da_eliminare.append(idx)
                except Exception:
                    try:
                        data_fine_valida = datetime.strptime(testo_fine, "%d-%m-%Y")
                        if data_fine_valida.date() < oggi_ora.date():
                            indici_da_eliminare.append(idx)
                    except Exception: pass
        
        # 🛡️ SE CI SONO SCADENZE: Esegue la rimozione col .pop() identica al comando manuale
        if indici_da_eliminare:
            st.session_state.congelamento_sincro_attivo = True  # Blocca temporaneamente la RAM
            
            # Rimuove i record partendo dall'ultimo per non sfasare gli indici della lista
            for idx in sorted(indici_da_eliminare, reverse=True):
                st.session_state.storico_cloud.pop(idx)
                
            # Rigenera il database aggiornato dall'elenco rimasto
            df_nuovo_salva = pd.DataFrame(st.session_state.storico_cloud)
            
            # Forza la stesura su disco e la spinta cloud con la sequenza collaudata
            df_nuovo_salva.to_excel(FILE_STORICO_PERMANENTE, index=False)
            push_excel_su_github(df_nuovo_salva)
            
            st.session_state.congelamento_sincro_attivo = False  # Sblocca la RAM
            st.toast("🧹 Pulizia automatica: Rimossi i locali che hanno terminato le ferie!")
            time.sleep(0.5)
            st.rerun()
            
    return df_l, df_t, df_s

    
df_locali, df_tecnici, df_storico_file = carica_database_locale()

# 🛡️ AUTOMAZIONE DI MANUELA: Forza l'app a leggere l'Excel reale aggiornato dal robot, distruggendo la cache vecchia
df_aggiornato_reale = pd.read_excel(FILE_STORICO_PERMANENTE).fillna("") if os.path.exists(FILE_STORICO_PERMANENTE) else df_storico_file
st.session_state.storico_cloud = df_aggiornato_reale.to_dict('records')


def push_excel_su_github(df_da_salvare):
    try:
        t_git = str(st.secrets["github"]["token_accesso"]).strip()
        # 🛡️ ENDPOINT API PULITO: Indirizzo nativo privo di _nonce per recuperare lo SHA reale senza errori 404
        url_git = f"https://api.github.com/repos/wingamingsrl/sistema-ferie/contents/{FILE_STORICO_PERMANENTE}"
        
        # Converte rigidamente il database in stringhe testuali pure per bypassare i firewall di GitHub
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
        
        # Recupera lo SHA reale e aggiornato direttamente dal file sul sito
        res_get = requests.get(url_git, headers=headers_git, timeout=5)
        sha_file = res_get.json().get("sha", "") if res_get.status_code == 200 else ""
        
        payload_git = {
            "message": "🤖 [App] Sincronizzazione ed allineamento database ferie", 
            "content": str(dati_base64), 
            "branch": "main"
        }
        if sha_file: 
            payload_git["sha"] = str(sha_file)
            
        risposta_put = requests.put(url_git, json=payload_git, headers=headers_git, timeout=5)
        
        # Gestione di sblocco in caso di collisioni simultanee di rete (Codice 422)
        if risposta_put.status_code == 422:
            res_retry = requests.get(url_git, headers=headers_git, timeout=5)
            if res_retry.status_code == 200:
                payload_git["sha"] = str(res_retry.json().get("sha", ""))
                risposta_put = requests.put(url_git, json=payload_git, headers=headers_git, timeout=5)
            else:
                if "sha" in payload_git: 
                    del payload_git["sha"]
                risposta_put = requests.put(url_git, json=payload_git, headers=headers_git, timeout=5)
                
        if risposta_put.status_code == 200 or risposta_put.status_code == 201:
            # 🛡️ DISATTIVATO POP-UP: Il file si salva in background senza mostrare finestre volanti
            return True
        return False
    except Exception:
        return False
# =====================================================================================
# FUNZIONI INTEGRATE DEL ROBOT AUTOMATICO DI SINCRONIZZAZIONE .SNAI.IT
# =====================================================================================
def genera_codice_otp_automatico():
    chiave_pulita = CHIAVE_SEGRETA_2FA.strip().upper().replace(" ", "")
    totp = pyotp.TOTP(chiave_pulita)
    return totp.now()

def esegui_sincronizzazione_robot_snai():
    # 🛡️ BLINDATURA TOTALE: Cambiamo il testo in 'Fase Finale' per verificare l'effettivo aggiornamento del file
    st.info("🎯 TELECOMANDO CLOUD — Fase Finale: Verifica credenziali...")
    try:
        t_git = str(st.secrets["github"]["token_accesso"]).strip()
        st.write("📝 Fase 1a: Gettone di sicurezza rintracciato in memoria.")
            
        # 🛡️ COSTRUZIONE STRUTTURALE PEZZO PER PEZZO: Impedisce la sovrascrittura o il troncamento della cache di Streamlit
        protocollo = "https://"
        dominio_api = "api.github.com"
        percorso_repo = "/repos/wingamingsrl/sistema-ferie"    
        percorso_workflow = "/actions/workflows/cron_robot_snai.yml/dispatches"
        
        # Unisce i blocchi creando la stringa estesa senza rischiare tagli
        url_workflow = f"{protocollo}{dominio_api}{percorso_repo}{percorso_workflow}"
        st.write(f"🔍 Fase 2: Indirizzo di rete del Workflow configurato -> `{url_workflow}`")
        
        headers_dispatch = {
            "Authorization": f"token {t_git}",
            "Accept": "application/vnd.github.v3+json",
            "User-Agent": "WinGaming-Cloud-App"
        }
        
        payload_dispatch = {
            "ref": "main"
        }
        
        st.info("🛰️ Fase 3: Spedizione del segnale di innesco a GitHub Actions...")
        risposta_remota = requests.post(url_workflow, json=payload_dispatch, headers=headers_dispatch, timeout=10)
        
        st.warning(f"📊 Fase 4: Riscontro del server di automazione. Codice numerico -> {risposta_remota.status_code}")
        
        if risposta_remota.status_code == 204 or risposta_remota.status_code == 202:
            st.success("🚀 Fase 5: ROBOT AUTOMATICO ATTIVATO!\n\nIl server esterno si è acceso correttamente ed ha avviato Chrome. Tra circa due minuti le ferie inserite saranno visibili sul portale Snaitech.")
            st.warning("⏱️ Schermo congelato per 10 secondi per consentire la lettura...")
            time.sleep(10)
            return True
        else:
            st.error(f"❌ Fase 5: Attivazione respinta dal server. Dettaglio: {risposta_remota.text}")
            st.warning("⏱️ Schermo congelato per 10 secondi...")
            time.sleep(10)
            return False
            
    except Exception as e_step:
        st.error(f"💥 FASE FALLITA: Errore interno di sistema -> {str(e_step)}")
        st.warning("⏱️ Schermo congelato per 10 secondi...")
        time.sleep(10)
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
        
        if st.button("🚀 ACCEDI AL PORTALE"):
            # Verifica le credenziali inserite dall'ufficio
            utente_trovato = df_tecnici[(df_tecnici["EMAIL"].astype(str).str.lower().str.strip() == input_email) & (df_tecnici["PASSWORD"].astype(str).str.strip() == input_password)]
            if not utente_trovato.empty:
                st.session_state.user_nome = str(utente_trovato.iloc[0]["NOME"]).strip()
                st.session_state.user_email = str(utente_trovato.iloc[0]["EMAIL"]).strip()
                st.session_state.autenticato = True
                # 🛡️ ALLINEAMENTO LIVE DI MANUELA: Scrive il token nell'URL della barra internet per mantenere la sessione 2 ore
                st.query_params["token_sessione"] = f"{st.session_state.user_email}_attivo"

                # 🧹 INNESTO AUTOMATICO DI MANUELA: Spazzino istantaneo nativo al momento del Login
                if os.path.exists(FILE_STORICO_PERMANENTE):
                    df_s_login = pd.read_excel(FILE_STORICO_PERMANENTE).fillna("")
                    st.session_state.storico_cloud = df_s_login.to_dict('records')
                    
                    oggi_ora = datetime.now()
                    indici_da_eliminare = []
                    
                    # Individua gli indici esatti delle scadenze passate (locali che hanno già riaperto)
                    for idx, row in enumerate(st.session_state.storico_cloud):
                        testo_fine = str(row.get("FINE_FERIE", "")).strip()
                        if testo_fine:
                            try:
                                data_fine_valida = datetime.strptime(testo_fine, "%d-%m-%Y %H:%M")
                                if data_fine_valida < oggi_ora:
                                    indici_da_eliminare.append(idx)
                            except Exception:
                                try:
                                    data_fine_valida = datetime.strptime(testo_fine, "%d-%m-%Y")
                                    if data_fine_valida.date() < oggi_ora.date():
                                        indici_da_eliminare.append(idx)
                                except Exception: pass
                    
                    # Se rileva record scaduti, applica la sequenza nativa esatta del tasto elimina manuale
                    if indici_da_eliminare:
                        st.session_state.congelamento_sincro_attivo = True
                        for idx in sorted(indici_da_eliminare, reverse=True):
                            st.session_state.storico_cloud.pop(idx)
                            
                        df_nuovo_salva = pd.DataFrame(st.session_state.storico_cloud)
                        # Scrittura fisica obbligatoria su disco sul server prima della spinta cloud
                        df_nuovo_salva.to_excel(FILE_STORICO_PERMANENTE, index=False)
                        push_excel_su_github(df_nuovo_salva)
                        st.session_state.congelamento_sincro_attivo = False
                
                st.success(f"🔓 Benvenuta {st.session_state.user_nome}!")
                time.sleep(1.0)
                st.rerun()
            else:
                st.error("❌ Credenziali errate. Riprova.")
    st.stop()


esecutore_nome = st.session_state.user_nome
esecutore_email = st.session_state.user_email

st.markdown("<h1>🧳 PORTALE FERIE ESERCENTI</h1>", unsafe_allow_html=True)
st.markdown(f"<div class='user-badge'>👤 {esecutore_nome} ({esecutore_email})</div>", unsafe_allow_html=True)



# =====================================================================================
# BLOCCO 4: MOTORE NOTIFICA EMAIL SMTP GOOGLE CON CONVERSIONE ROTTA IP RIGIDA
# AGGIRA MANUALMENTE I BLACKOUT DELLE RETI PROTETTE DEI SERVER CLOUD DI STREAMLIT
# =====================================================================================
def invia_mail_diretta_smtp(lista_m, locale, ionario_testo, chiusura, riapertura, esecutore):
    try:
        pass_gmail = str(st.secrets["gmail"]["password_applicativa"]).strip()
        msg = MIMEMultipart()
        msg['From'] = EMAIL_MITTENTE_GMAIL
        msg['To'] = ", ".join(lista_m)
        msg['Subject'] = f"🛡️ Registrazione Chiusura Ferie - {locale}"
        
        linee_ionari = ""
        elenco_conc = [c.strip() for c in ionario_testo.split(",") if c.strip()]
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
    with col1: 
        # 🛡️ BLINDATURA CALENDARIO DI MANUELA: Ripristino layout classico con sblocco tendina dei mesi rapida sul telefono
        data_chiusura = st.date_input(
            "Giorno Chiusura:", 
            value=datetime.now().date(), 
            min_value=datetime(2025, 1, 1).date(), 
            max_value=datetime(2030, 12, 31).date(),
            format="DD-MM-YYYY",
            key="cal_chiusura_definitivo_manuela"
        )
    with col2: ora_chiusura = st.time_input("Ora Chiusura:", dtime(6, 0))
    
    st.markdown("---")
    col3, col4 = st.columns(2)
    #with col3: data_riapertura = st.date_input("Giorno Riapertura:", datetime.now() + timedelta(days=14), format="DD-MM-YYYY")
    with col3: 
        # 🛡️ BLINDATURA CALENDARIO DI MANUELA: Ripristino layout classico con sblocco tendina dei mesi rapida sul telefono
        data_riapertura = st.date_input(
            "Giorno Riapertura:", 
            value=datetime.now().date(), 
            min_value=datetime(2025, 1, 1).date(), 
            max_value=datetime(2030, 12, 31).date(),
            format="DD-MM-YYYY",
            key="cal_riapertura_definitivo_manuela"
        )
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
            # 🛡️ FIX DATA INSERIMENTO ALL'ITALIANA: Formato Giorno-Mese-Anno con secondi reali
            data_inserimento_it = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

            # 🛡️ AUTOMAZIONE DI MANUELA: Calcola l'azione esatta usando solo la variabile nativa dell'App
            tipo_azione_snai = "MODIFICA" if forza_sovrascrittura else "NUOVA"

            # 🛡️ ARCHITETTURA DI MANUELA: Scansiona l'anagrafica (df_locali) per trovare tutti i provider di questo codice locale
            codice_cercato_target = str(codice_estratto).strip()
            
            try:
                df_filtro_anagrafica = df_locali[df_locali["CODICE_LOCALE"].astype(str).str.strip() == codice_cercato_target]
                lista_concessionari_rilevati = df_filtro_anagrafica["CONCESSIONARIO"].astype(str).str.strip().unique().tolist()
                lista_concessionari_rilevati = [c for c in lista_concessionari_rilevati if c and c != "nan"]
            except Exception:
                lista_concessionari_rilevati = []

            if not lista_concessionari_rilevati:
                lista_concessionari_rilevati = [str(concessionario_estratto).strip()] if concessionario_estratto else ["Snaitech Spa WG"]

            # Mappa e standardizza i nomi dei provider per l'ufficio
            lista_provider_puliti = []
            for pvd in lista_concessionari_rilevati:
                pvd_up = pvd.upper()
                if "SNAI" in pvd_up: lista_provider_puliti.append("Snaitech Spa WG")
                elif "NTS" in pvd_up: lista_provider_puliti.append("NTS Networks")
                elif "SISAL" in pvd_up: lista_provider_puliti.append("Sisal")
                elif "GLOBAL" in pvd_up: lista_provider_puliti.append("Global Starnet")
                else: lista_provider_puliti.append(pvd)

            # 🛡️ IL MOLTIPLICATORE PER RIGHE MULTIPLE: Genera un record nelle ferie per OGNI riga provider trovata in anagrafica!
            righe_sdoppiate_da_salvare = []
            for provider_singolo in lista_provider_puliti:
                riga_singola = {
                    "DATA_INSERIMENTO": str(data_inserimento_it),
                    "TECNICO_INSERIMENTO": str(esecutore_nome),
                    "CODICE_LOCALE": str(codice_estratto),
                    "NOME_LOCALE": str(nome_puro_locale),
                    "CONCESSIONARIO": str(provider_singolo),
                    "INIZIO_FERIE": str(str_c),
                    "FINE_FERIE": str(str_r),
                    "PROMEMORIA_IN_COPIA": str(co_destinatario),
                    "STATO_INVIO": "In attesa",
                    "ROBOT_ACTION": "MODIFICA" if forza_sovrascrittura else "NUOVA"
                }
                righe_sdoppiate_da_salvare.append(riga_singola)
            
            lista_m = [EMAIL_MANUELA_RICEVENTE, esecutore_email]
            if co_destinatario != "Nessun collega" and " (" in str(co_destinatario):
                try: lista_m.append(co_destinatario.split(" (")[-1].replace(")", "").strip())
                except Exception: pass
            
            titolo_azione = "Modifica Chiusura" if (sovrapposizione_rilevata and forza_sovrascrittura) else "Registrazione Chiusura"
            invio_ok = False
            risposta_server = "OK"
            
            with st.spinner("Salvataggio e invio notifica..."):
                try:
                    pass_gmail = str(st.secrets["gmail"]["password_applicativa"]).strip()
                    msg = MIMEMultipart()
                    msg['From'] = EMAIL_MITTENTE_GMAIL
                    msg['To'] = ", ".join(lista_m)
                    msg['Subject'] = f"🛡️ {titolo_azione} Ferie - {nome_puro_locale}"
                    
                    # 🚨 UNIFICAZIONE E-MAIL: Mostra tutti i concessionari separati da virgola nell'unica mail riassuntiva
                    testo_concessionari_mail = ", ".join(lista_provider_puliti)
                    corpo = f"Rilevato aggiornamento chiusura ferie nel sistema WinGaming.\n\nDettagli della pratica:\n--------------------------------------------------\n🔔 Stato Operazione:  {titolo_azione.upper()}\n👤 Tecnico Esecutore: {esecutore_nome}\n📍 Locale Coinvolto:  {nome_puro_locale}\n🏢 Concessionario/i:  {testo_concessionari_mail}\n📅 Inizio Chiusura:   {str_c}\n🚚 Data Riapertura:   {str_r}\n--------------------------------------------------\n\nWINGAMING SRL"
                    msg.attach(MIMEText(corpo, 'plain'))
                    
                    server = smtplib.SMTP_SSL('64.233.184.108', 465, timeout=10)
                    server.login(EMAIL_MITTENTE_GMAIL, pass_gmail)
                    server.sendmail(EMAIL_MITTENTE_GMAIL, lista_m, msg.as_string())
                    server.quit()
                    invio_ok = True
                except Exception as e_mail:
                    risposta_server = str(e_mail)
            
            if invio_ok:
                if sovrapposizione_rilevata and riga_conflitto_idx is not None:
                    try: st.session_state.storico_cloud.pop(riga_conflitto_idx)
                    except Exception: pass
                
                # Inietta tutte le righe sdoppiate in tempo reale nella memoria dello smartphone
                for record_sdoppiato in righe_sdoppiate_da_salvare:
                    record_sdoppiato["STATO_INVIO"] = "Inviato OK"
                    st.session_state.storico_cloud.append(record_sdoppiato)
                    
                df_salva = pd.DataFrame(st.session_state.storico_cloud)
                
                if "ROBOT_ACTION" not in df_salva.columns:
                    df_salva["ROBOT_ACTION"] = ""
                    
                df_salva.to_excel(FILE_STORICO_PERMANENTE, index=False)

                
                # Sblocca ed aggiorna istantaneamente GitHub spingendo i dati online
                push_excel_su_github(df_salva)
                
                st.success("✅ OPERAZIONE COMPLETATA!\n\nPratica registrata correttamente a sistema e notifica e-mail inviata.")
                st.session_state.form_id += 1
                time.sleep(2.0)
                st.rerun()


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
        # 🛡️ FILTRO DI MANUELA INTEGRALE: Scansiona la RAM e nasconde i locali da eliminare senza usare Pandas
        lista_visibile = [
            riga for riga in st.session_state.storico_cloud 
            if str(riga.get("ROBOT_ACTION", "")).strip().upper() != "ELIMINA" and 
               str(riga.get("robot_action", "")).strip().upper() != "ELIMINA"
        ]
        
        # Genera la tabella solo con i locali rimasti attivi
        df_vis = pd.DataFrame(lista_visibile)
        
        if not df_vis.empty:
            df_vis = df_vis.reindex(columns=colonne_reali_ufficio).fillna("")
            st.dataframe(df_vis, hide_index=True)
        else:
            st.info("📭 Nessuna chiusura attiva presente nel registro storico.")



        
        with io.BytesIO() as buffer:
            df_vis.to_excel(buffer, index=False)
            st.download_button(label="📥 Scarica Registro Excel Storico", data=buffer.getvalue(), file_name="storico_ferie.xlsx", mime="application/vnd.ms-excel")
    else:
        st.info("📭 Nessuna chiusura presente in memoria. Trascina il file Excel storico in fondo per ripopolare la plancia.")
        

    # =====================================================================================
    # TABELLA SINCRO PORTALE SNAITECH - MOSTRA SOLO I LOCALI CON UN'ACTION DA FARE
    # =====================================================================================
    st.markdown("---")
    st.markdown("### 🏢 Concessionari pronti da inviare a sistema")
    st.write("Questo comando attiva il robot che effettua l'invio delle e-mail dirette per NTS e la sincronizzazione automatica su .snai.it.")
    
    if st.button("🚀 AVVIA SINCRONIZZAZIONE FORZATA SU PORTALI / EMAIL"):
        with st.spinner("Robot in azione sui sistemi dei Concessionari... Non chiudere la pagina..."):
            esegui_sincronizzazione_robot_snai()
            
            # 🛡️ RE-SHAPE DI MANUELA: Pausa per dare tempo a GitHub di digerire il file Excel inviato dal robot
            time.sleep(5)
            
            # Svuota lo stato precedente e costringe lo smartphone a ricaricare l'Excel pulito dal server cloud
            if os.path.exists(FILE_STORICO_PERMANENTE):
                # Rilegge il file fisico aggiornato dallo spazzino del robot
                df_aggiornato_cloud = pd.read_excel(FILE_STORICO_PERMANENTE).fillna("")
                st.session_state.storico_cloud = df_aggiornato_cloud.to_dict('records')
            
            # Rinfresca l'interfaccia eliminando le righe azzerate
            st.rerun()


    # 🛡️ FILTRO INTERCETTATORE DI MANUELA: Mostra in tabella TUTTI i locali pronti (SNAI + NTS) con un'azione reale da compiere
    righe_lavorazione_generiche = [
        row for row in st.session_state.storico_cloud 
        if str(row.get("ROBOT_ACTION", "")).strip().upper() in ["NUOVA", "MODIFICA", "ELIMINA"]
    ] if st.session_state.storico_cloud else []
    
    if righe_lavorazione_generiche:
        df_lavorazione = pd.DataFrame(righe_lavorazione_generiche)
        # 🛡️ COLONNA INSERITA: Aggiunta la colonna CONCESSIONARIO nel tabellone visivo dello smartphone
        colonne_visibili = ["CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO", "INIZIO_FERIE", "FINE_FERIE", "ROBOT_ACTION", "TECNICO_INSERIMENTO"]
        df_visibile_pulito = df_lavorazione.reindex(columns=colonne_visibili).fillna("")
        st.dataframe(df_visibile_pulito, hide_index=True)
    else:
        st.success("✅ Nessun locale in attesa. Tutte le chiusure dei Concessionari sono allineate!")

    # =====================================================================================
    # PANNELLO CANCELLAZIONE - COMPRESSIONE MENÙ A TENDINA E SPARIZIONE TASTO SMARTPHONE
    # =====================================================================================
    st.markdown("---")
    st.markdown("### 🗑️ Cancella un Periodo Registrato")
    
    opzioni_cancellazione = ["- Seleziona la riga da eliminare -"]
    mappa_indici_reali = {}
    
    # 🛡️ FILTRO MENÙ DI MANUELA: Scansiona la RAM e inserisce nella tendina SOLO i locali che non sono già in stato ELIMINA
    if st.session_state.storico_cloud:
        for idx, row in enumerate(st.session_state.storico_cloud):
            azione_corrente = str(row.get("ROBOT_ACTION", "")).strip().upper()
            if azione_corrente != "ELIMINA":
                testo_opzione = f"ID {idx} | {row.get('CODICE_LOCALE', '')} - {row.get('NOME_LOCALE', '')} (Dal {row.get('INIZIO_FERIE', '')})"
                opzioni_cancellazione.append(testo_opzione)
                mappa_indici_reali[testo_opzione] = idx
            
    selezione_delete = st.selectbox("Scegli la chiusura da eliminare dal database:", opzioni_cancellazione, disabled=len(opzioni_cancellazione) <= 1)
    
    # 🛡️ BOTTONE FANTASMA DI MANUELA: Il tasto compare SOLO se hai selezionato un locale valido, se rimetti la voce standard sparisce nel nulla!
    if selezione_delete != "- Seleziona la riga da eliminare -" and selezione_delete in mappa_indici_reali:
        try:
            idx_selezionato = mappa_indici_reali[selezione_delete]
            
            # Estrae il codice del locale selezionato nel menù a tendina
            riga_scelta = st.session_state.storico_cloud[idx_selezionato]
            codice_locale_target = str(riga_scelta.get("CODICE_LOCALE", "")).strip()
            nome_locale_target = str(riga_scelta.get("NOME_LOCALE", "")).strip()
                
            if st.button("❌ ELIMINA DEFINITIVAMENTE QUESTA CHIUSURA (PER TUTTI I PROVIDER)"):
                st.session_state.congelamento_sincro_attivo = True  # Protezione RAM
                
                # 🛡️ SCANSIONE GLOBALE DI MANUELA: Cerca e marchia come ELIMINA tutti i provider dello stesso locale
                for riga_cloud in st.session_state.storico_cloud:
                    if str(riga_cloud.get("CODICE_LOCALE", "")).strip() == codice_locale_target:
                        riga_cloud["ROBOT_ACTION"] = "ELIMINA"
                
                df_nuovo_salva = pd.DataFrame(st.session_state.storico_cloud)
                
                # Forza la scrittura fisica dell'Excel su disco prima di inviarlo
                df_nuovo_salva.to_excel(FILE_STORICO_PERMANENTE, index=False)
                
                # Spinge il file modificato su GitHub
                push_excel_su_github(df_nuovo_salva)
                
                st.session_state.congelamento_sincro_attivo = False  # Sblocca RAM
                st.success(f"🗑️ Richiesta di eliminazione inviata per tutti i provider del locale: {nome_locale_target}! Le righe sono state nascoste.")
                time.sleep(1.5)
                st.rerun()
        except Exception as e_del: 
            st.error(f"❌ Errore durante la rimozione: {str(e_del)}")

        
    st.markdown("---")
    st.markdown("### 📤 Ricarica Registro Excel Aggiornato dall'Ufficio")
    file_caricato = st.file_uploader("Trascina il file storico_ferie.xlsx modificato per caricare i dati nel portale:", type=["xlsx"])
    if file_caricato is not None:
        try:
            df_caricato = pd.read_excel(file_caricato).fillna("")
            if "CODICE_LOCALE" in df_caricato.columns:
                if st.button("🔄 CONFERMA E SOVRASCRIVI DATABASE CON QUESTO FILE"):
                    for col_data in ["DATA_INSERIMENTO", "INIZIO_FERIE", "FINE_FERIE"]:
                        if col_data in df_caricato.columns:
                            try:
                                if col_data == "DATA_INSERIMENTO":
                                    df_caricato[col_data] = pd.to_datetime(df_caricato[col_data]).dt.strftime('%d-%m-%Y %H:%M:%S')
                                else:
                                    df_caricato[col_data] = pd.to_datetime(df_caricato[col_data]).dt.strftime('%d-%m-%Y %H:%M')
                            except Exception:
                                df_caricato[col_data] = df_caricato[col_data].astype(str).str.strip()
                    
                    st.session_state.storico_cloud = df_caricato.to_dict('records')
                    df_caricato.to_excel(FILE_STORICO_PERMANENTE, index=False)
                    push_excel_su_github(df_caricato)
                    st.success("✅ Database aziendale aggiornato e sincronizzato con successo!")
                    time.sleep(2.0)
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

