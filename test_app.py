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
# 🛡️ BLOCCO DI RETE DI MANUELA: DISATTIVA I BOTTONI DI TUTTI I TECNICI SE IL ROBOT GIRA
# =====================================================================================
robot_sta_girando_ora = False
try:
    # Rilegge il file fisico presente sul server di GitHub per verificare i processi
    df_lock_rete = pd.read_excel(FILE_STORICO_PERMANENTE).fillna("")
    
    # Condizione di sicurezza: Se nel cloud ci sono righe con lo STATO_INVIO impostato dal robot 
    # o se la sincronizzazione forzata visiva è attiva sulla rete, blocca tutti i telefoni!
    if not df_lock_rete.empty and any(str(row.get("STATO_INVIO", "")).strip() == "In elaborazione" for _, row in df_lock_rete.iterrows()):
        robot_sta_girando_ora = True
except Exception:
    pass
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
    .stButton>button { background-color: #f1f5f9 !important; color: #0f172a !important; border: 1px solid #cbd5e1 !important; font-weight: 600 !important; font-size: 15px !important; width: 100%; border-radius: 8px !important; height: 42px !important; box-shadow: none !important; transition: all 0.2s ease; }
    .stButton>button:hover { background-color: #e2e8f0 !important; border-color: #94a3b8 !important; color: #020617 !important; }
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
    df_l = pd.read_excel(FILE_LOCALI).fillna("") if os.path.exists(FILE_LOCALI) else pd.DataFrame(columns=["CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO"])
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

esecutore_nome = st.session_state.get("user_nome", "UFFICIO")
esecutore_email = st.session_state.get("user_email", "manuela.arigoni@wingaming.it")


st.markdown("<h1>🧳 PORTALE FERIE ESERCENTI</h1>", unsafe_allow_html=True)
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
    submit_button = st.form_submit_button("💾 INVIA CHIUSURA TEMPORANEA", disabled=robot_sta_girando_ora)

# =====================================================================================
# BLOCCO 6: ELABORAZIONE INSERIMENTI (VERSIONE PULITA - INIEZIONE SINGOLA)
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
        
        for idx, row in enumerate(st.session_state.get("storico_cloud", [])):
            if str(row.get("NOME_LOCALE", "")).strip() in testo_pvd or testo_pvd.strip() in str(row.get("NOME_LOCALE", "")):
                try:
                    old_i = datetime.strptime(str(row.get("INIZIO_FERIE", "")).strip(), "%d-%m-%Y %H:%M")
                    old_f = datetime.strptime(str(row.get("FINE_FERIE", "")).strip(), "%d-%m-%Y %H:%M")
                    if (data_inizio_nuova <= old_f) and (data_fine_nuova >= old_i):
                        sovrapposizione_rilevata, riga_conflitto_idx = True, idx
                        dettagli_conflitto = f"Dal {row.get('INIZIO_FERIE','')} al {row.get('FINE_FERIE','')}"
                        break
                except Exception:
                    pass

        if './' in str(testo_pvd) or '/' in str(testo_pvd):
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
            
            data_inserimento_it = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

            try:
                df_filtro_anagrafica = df_locali[df_locali["CODICE_LOCALE"].astype(str).str.strip() == str(codice_estratto).strip()]
                lista_concessionari_rilevati = df_filtro_anagrafica["CONCESSIONARIO"].astype(str).str.strip().unique().tolist()
                lista_concessionari_rilevati = [c for c in lista_concessionari_rilevati if c and c != "nan"]
            except Exception:
                lista_concessionari_rilevati = []

            if not lista_concessionari_rilevati:
                lista_concessionari_rilevati = [str(concessionario_estratto).strip()] if concessionario_estratto else ["Snaitech Spa WG"]

            lista_provider_puliti = []
            for pvd in lista_concessionari_rilevati:
                pvd_up = pvd.upper()
                if "SNAI" in pvd_up: lista_provider_puliti.append("Snaitech Spa WG")
                elif "NTS" in pvd_up: lista_provider_puliti.append("NTS Networks")
                elif "SISAL" in pvd_up: lista_provider_puliti.append("Sisal")
                elif "GLOBAL" in pvd_up: lista_provider_puliti.append("Global Starnet")
                else: lista_provider_puliti.append(pvd)

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
                    
                    testo_concessionari_mail = ", ".join(lista_provider_puliti)
                    corpo = f"Rilevato aggiornamento chiusura ferie nel sistema WinGaming.\n\nDettagli della prima nota:\n--------------------------------------------------\n🔔 Stato Operazione:  {titolo_azione.upper()}\n👤 Tecnico Esecutore: {esecutore_nome}\n📍 Locale Coinvolto:  {nome_puro_locale}\n🏢 Concessionario/i:  {testo_concessionari_mail}\n📅 Inizio Chiusura:   {str_c}\n🚚 Data Riapertura:   {str_r}\n--------------------------------------------------\n\nWINGAMING SRL"
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
                
                # 🛡️ INIEZIONE UNICA GARANTITA: Una sola volta senza raddoppi!
                for record_sdoppiato in righe_sdoppiate_da_salvare:
                    record_sdoppiato["STATO_INVIO"] = "Inviato OK"
                    if "storico_cloud" not in st.session_state:
                        st.session_state.storico_cloud = []
                    st.session_state.storico_cloud.append(record_sdoppiato)
                    
                df_salva = pd.DataFrame(st.session_state.storico_cloud)
                if "ROBOT_ACTION" not in df_salva.columns:
                    df_salva["ROBOT_ACTION"] = ""
        
                df_salva.to_excel(FILE_STORICO_PERMANENTE, index=False)
                push_excel_su_github(df_salva)

                utente_corrente_maiuscolo = str(st.session_state.get("user_nome", "UFFICIO")).strip().upper()
                if utente_corrente_maiuscolo in ["MANUELA ARIGONI", "ADMIN", "UFFICIO"] or utente_corrente_maiuscolo == str(esecutore_nome).strip().upper():
                    st.success(f"✅ OPERAZIONE COMPLETATA!\n\nPratica registrata correttamente a sistema per tutti i provider di: {nome_puro_locale}.")
                
                st.session_state.form_id += 1
                time.sleep(4.0)
                st.rerun()
            else:
                st.error(f"❌ Errore Google SMTP: {risposta_server}. Spedizione e-mail fallita.")
# =====================================================================================
# BLOCCO 7: PROMEMORIA LOGISTICI 3 GG E ALLERTA GIALLA VISIVA DI MANUELA
# =====================================================================================
st.markdown("### 🔔 Scadenze Logistiche Imminenti (3 Giorni)")
try:
    oggi_plancia = datetime.now().date()
    ha_avvisi = False
    
    email_corrente_check = str(st.session_state.get("user_email", "")).strip().lower()
    utente_corrente_check = str(st.session_state.get("user_nome", "UFFICIO")).strip().upper()
    
    if st.session_state.get("storico_cloud", []):
        for row in st.session_state.storico_cloud:
            tecnico_riga = str(row.get("TECNICO_INSERIMENTO", "")).strip().upper()
            
            if "manuela" in email_corrente_check or "admin" in email_corrente_check or "ufficio" in email_corrente_check or utente_corrente_check == tecnico_riga:
                data_in_raw = str(row.get("INIZIO_FERIE", "")).strip()
                data_fi_raw = str(row.get("FINE_FERIE", "")).strip()
                nome_loc_avviso = str(row.get("NOME_LOCALE", "")).strip()
                
                str_in_pax = data_in_raw.replace("-", "/").replace(".", "/").strip()[:10]
                str_fi_pax = data_fi_raw.replace("-", "/").replace(".", "/").strip()[:10]
                
                try:
                    dt_in_check = datetime.strptime(str_in_pax, "%d/%m/%Y").date()
                    dt_fi_check = datetime.strptime(str_fi_pax, "%d/%m/%Y").date()
                    giorni_chiusura = (dt_in_check - oggi_plancia).days
                    giorni_riapertura = (dt_fi_check - oggi_plancia).days
                    
                    if giorni_chiusura == 3:
                        st.warning(f"⚠️ **PROMEMORIA CHIUSURA (TRA 3 GG):** Il locale **{nome_loc_avviso}** chiude il {dt_in_check.strftime('%d/%m/%Y')} (Tecnico: {row.get('TECNICO_INSERIMENTO','')})")
                        ha_avvisi = True
                    if giorni_riapertura == 3:
                        st.warning(f"🚚 **PROMEMORIA RIAPERTURA (TRA 3 GG):** Il locale **{nome_loc_avviso}** riapre il {dt_fi_check.strftime('%d/%m/%Y')} (Tecnico: {row.get('TECNICO_INSERIMENTO','')})")
                        ha_avvisi = True
                except Exception: pass
                    
    if not ha_avvisi:
        st.info("💡 Nessun locale in scadenza a 3 giorni per la tua utenza.")
except Exception: pass

# =====================================================================================
# BLOCCO 8: TABELLONE VISIVO DI MANUELA: PRIVILEGI ADMIN / TECNICI (VISTA IN ATTESA)
# =====================================================================================
st.markdown("---")
email_tab_check = str(st.session_state.get("user_email", "")).strip().lower()
utente_tab_check = str(st.session_state.get("user_nome", "UFFICIO")).strip().upper()

if "manuela" in email_tab_check or "admin" in email_tab_check or "ufficio" in email_tab_check:
    st.markdown("### 📊 [VISTA ADMIN] Tutte le chiusure della flotta in attesa")
    righe_lavorazione_generiche = [
        row for row in st.session_state.get("storico_cloud", []) 
        if str(row.get("ROBOT_ACTION", "")).strip().upper() in ["NUOVA", "MODIFICA", "ELIMINA"]
    ]
else:
    st.markdown("### 📊 Le tue chiusure in attesa di allineamento")
    righe_lavorazione_generiche = [
        row for row in st.session_state.get("storico_cloud", []) 
        if str(row.get("ROBOT_ACTION", "")).strip().upper() in ["NUOVA", "MODIFICA", "ELIMINA"]
        and str(row.get("TECNICO_INSERIMENTO", "")).strip().upper() == utente_tab_check
    ]

if righe_lavorazione_generiche:
    df_lavorazione = pd.DataFrame(righe_lavorazione_generiche)
    if "manuela" in email_tab_check or "admin" in email_tab_check or "ufficio" in email_tab_check:
        colonne_visibili = ["CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO", "INIZIO_FERIE", "FINE_FERIE", "ROBOT_ACTION", "TECNICO_INSERIMENTO"]
    else:
        colonne_visibili = ["CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO", "INIZIO_FERIE", "FINE_FERIE", "ROBOT_ACTION"]
    df_visibile_pulito = df_lavorazione.reindex(columns=colonne_visibili).fillna("")
    st.dataframe(df_visibile_pulito, hide_index=True)
else:
    st.success(f"✅ Nessuna pratica in coda per la tua visualizzazione, {utente_tab_check}!")
# =====================================================================================
# BLOCCO 9: TABELLONE GIRI LOGISTICI DI MANUELA: PRIVILEGI GERARCHICI TOTALI (STORICO INCLUSO)
# =====================================================================================
st.markdown("---")
email_loggata_pulita = str(st.session_state.get("user_email", "")).strip().lower()
utente_loggato_maiuscolo = str(st.session_state.get("user_nome", "UFFICIO")).strip().upper()

# 🛡️ PARACADUTE DI SICUREZZA DI MANUELA: Previene il NameError in caso di logout azzerando la coda
righe_giri_logistici = []

if "manuela" in email_loggata_pulita or "admin" in email_loggata_pulita or "ufficio" in email_loggata_pulita:
    st.markdown("### 📊 [VISTA ADMIN] Tutti i Promemoria Giri Logistici della Flotta")
    righe_giri_logistici = st.session_state.get("storico_cloud", []) if st.session_state.get("storico_cloud", []) else []
else:
    if st.session_state.get("storico_cloud", []):
        righe_giri_logistici = [
            row for row in st.session_state.get("storico_cloud", []) 
            if str(row.get("TECNICO_INSERIMENTO", "")).strip().upper() == utente_loggato_maiuscolo
        ]

if righe_giri_logistici:
    df_lavorazione_giri = pd.DataFrame(righe_giri_logistici)
    if "manuela" in email_loggata_pulita or "admin" in email_loggata_pulita or "ufficio" in email_loggata_pulita:
        colonne_visibili_giri = ["CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO", "INIZIO_FERIE", "FINE_FERIE", "ROBOT_ACTION", "TECNICO_INSERIMENTO"]
    else:
        colonne_visibili_giri = ["CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO", "INIZIO_FERIE", "FINE_FERIE", "ROBOT_ACTION"]
    df_visibile_giri_pulito = df_lavorazione_giri.reindex(columns=colonne_visibili_giri).fillna("")
    st.dataframe(df_visibile_giri_pulito, hide_index=True)
else:
    st.success(f"✅ Nessun promemoria giro logistico registrato a sistema, {utente_loggato_maiuscolo}.")

# =====================================================================================
# BLOCCO 10: PANNELLO CANCELLAZIONE - SELEZIONE E RIMOZIONE RIGHE
# =====================================================================================
st.markdown("---")
st.markdown("### 🗑️ Cancella un Periodo Registrato")

opzioni_cancellazione = ["- Seleziona la riga da eliminare -"]
mappa_indici_reali = {}

if "storico_cloud" in st.session_state and st.session_state.storico_cloud:
    for idx, row in enumerate(st.session_state.storico_cloud):
        azione_corrente = str(row.get("ROBOT_ACTION", "")).strip().upper()
        if azione_corrente != "ELIMINA":
            testo_opzione = f"ID {idx} | {row.get('CODICE_LOCALE', '')} - {row.get('NOME_LOCALE', '')} [{row.get('CONCESSIONARIO','')}] (Dal {row.get('INIZIO_FERIE', '')})"
            opzioni_cancellazione.append(testo_opzione)
            mappa_indici_reali[testo_opzione] = idx
        
selezione_delete = st.selectbox("Scegli la chiusura da eliminare dal database:", opzioni_cancellazione, disabled=len(opzioni_cancellazione) <= 1)

if selezione_delete != "- Seleziona la riga da eliminare -" and selezione_delete in mappa_indici_reali:
    try:
        idx_selezionato = mappa_indici_reali[selezione_delete]
        riga_scelta = st.session_state.storico_cloud[idx_selezionato]
        codice_locale_target = str(riga_scelta.get("CODICE_LOCALE", "")).strip()
        nome_locale_target = str(riga_scelta.get("NOME_LOCALE", "")).strip()
            
        if st.button("❌ ELIMINA DEFINITIVAMENTE QUESTA CHIUSURA"):
            for riga_cloud in st.session_state.storico_cloud:
                if str(riga_cloud.get("CODICE_LOCALE", "")).strip() == codice_locale_target:
                    riga_cloud["ROBOT_ACTION"] = "ELIMINA"
            
            df_nuovo_salva = pd.DataFrame(st.session_state.storico_cloud)
            df_nuovo_salva.to_excel(FILE_STORICO_PERMANENTE, index=False)
            try: push_excel_su_github(df_nuovo_salva)
            except Exception: pass
            st.success(f"🗑️ Richiesta di eliminazione inviata per tutti i provider di: {nome_locale_target}!")
            time.sleep(1.5)
            st.rerun()
    except Exception as e_del: 
        st.error(f"❌ Errore durante la rimozione: {str(e_del)}")

# =====================================================================================
# AREA AMMINISTRATORE: UPLOADER EXCEL (VERSIONE INTEGRALE CONVERTITRICE)
# =====================================================================================
if "manuela" in email_loggata_pulita or "admin" in email_loggata_pulita or "ufficio" in email_loggata_pulita:
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
        except Exception as e_load: 
            st.error(f"❌ Errore lettura: {str(e_load)}")
# =====================================================================================
# BLOCCO 11: PRIVILEGI STRUTTURALI AUTOMATICI DA EMAIL (ESCLUSIVA ADMIN BLINDATA)
# =====================================================================================
st.markdown("---")
email_finale_pulita = str(st.session_state.get("user_email", "")).strip().lower()

# 🛡️ PRIVILEGIO ADMIN DI MANUELA: Se l'email contiene le chiavi dell'ufficio, apre il pannello completo
if "manuela" in email_finale_pulita or "admin" in email_finale_pulita or "ufficio" in email_finale_pulita:
    st.markdown("### 🏢 Concessionari pronti da inviare a sistema")
    st.write("Questo comando attiva il robot che effettua l'invio delle e-mail dirette per NTS e la sincronizzazione automatica su .snai.it.")

    if robot_sta_girando_ora:
        st.button("⚙️ ROBOT IN MARCIA SUI PORTALI... ATTENDI", disabled=True)
        st.warning("⏳ Un altro utente o l'Admin ha avviato il robot. La plancia è protetta. I tasti si riaccenderanno DA SOLI in automatico tra circa 2 minuti.")
        import time as t_sys
        t_sys.sleep(5)
        st.rerun()
    else:
        # Pulsante originale intatto dell'Admin
        if st.button("🚀 AVVIA SINCRONIZZAZIONE FORZATA", key="palo_sincro_admin_elastico_assoluto"):
            with st.spinner("Blindatura database aziendale e avvio server..."):
                try:
                    for riga_ram in st.session_state.storico_cloud:
                        if str(riga_ram.get("ROBOT_ACTION", "")).strip().upper() in ["NUOVA", "MODIFICA", "ELIMINA"]:
                            riga_ram["STATO_INVIO"] = "In elaborazione"
                    df_spingi_lock = pd.DataFrame(st.session_state.storico_cloud)
                    df_spingi_lock.to_excel(FILE_STORICO_PERMANENTE, index=False)
                    push_excel_su_github(df_spingi_lock)
                    esegui_sincronizzazione_robot_snai()
                    time.sleep(2)
                except Exception: pass
            st.rerun()
            
        # Il tasto disconnetti dell'Admin posizionato sotto il blu a sinistra
        if st.button("🚪 DISCONNETTI ACCOUNT", key="palo_logout_admin_elastico_assoluto"):
            st.session_state.clear()
            if "st" in locals() and hasattr(st, "query_params"):
                st.query_params.clear()
            st.rerun()

else:
    # 📱 VISTA TECNICI STANDARD: Qualsiasi altro indirizzo email vede SOLO ed ESCLUSIVAMENTE il logout a sinistra!
    if st.button("🚪 DISCONNETTI ACCOUNT", key="palo_logout_tecnico_elastico_assoluto"):
        st.session_state.clear()
        if "st" in locals() and hasattr(st, "query_params"):
            st.query_params.clear()
        st.rerun()
    st.stop() # 💥 GHIGLIOTTINA ASSOLUTA: Impedisce fisicamente al telefono del tecnico di leggere qualsiasi altra riga!
