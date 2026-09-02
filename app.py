# =====================================================================================
# SW GENERALE DI INSERIMENTO FERIE GESTORI WIN GAMING — PRODUZIONE INTEGRALE
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
import openpyxl  # 🛡️ FIX CHIRURGICO: Caricato correttamente il motore Excel in cima al file
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
# BLOCCO 2: COLLEGAMENTO ED INVIO RETE SU GITHUB MAIN (RIPRISTINATO)
# VERSIONE DI PRODUZIONE ALLINEATA CON TRACCIAMENTO OPENPYXL PROTETTO
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

if "storico_cloud" not in st.session_state:
    st.session_state.storico_cloud = df_storico_file.to_dict('records')

def push_excel_su_github(lista_records_da_salvare):
    try:
        t_git = str(st.secrets["github"]["token_accesso"]).strip()
        parte1 = "https://github.com"
        parte2 = "repos/wingamingsrl/sistema-ferie/contents"
        url_git = parte1 + "/" + parte2 + "/" + FILE_STORICO_PERMANENTE
        
        # 🛡️ Generazione nativa strutturata con openpyxl (Ora importato correttamente)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        
        colonne_reali = ["DATA_INSERIMENTO", "TECNICO_INSERIMENTO", "CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO", "INIZIO_FERIE", "FINE_FERIE", "PROMEMORIA_IN_COPIA", "STATO_INVIO"]
        ws.append(colonne_reali)
        
        for row in lista_records_da_salvare:
            ws.append([str(row.get(col, "")).strip() for col in colonne_reali])
            
        output_binario = io.BytesIO()
        wb.save(output_binario)
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
            "message": "🤖 [App] Sincronizzazione permanente database Excel", 
            "content": dati_base64,
            "branch": "main"
        }
        if sha_file: 
            payload_git["sha"] = sha_file
                        
        risposta_put = requests.put(url_git, json=payload_git, headers=headers_git, timeout=5)
        
        # 🛡️ BYPASS AUTOMATICO ERRORE 422: Se c'è un conflitto di SHA, pialla il file e lo rigenera nuovo da zero
        if risposta_put.status_code == 422 and sha_file:
            payload_delete = {"message": "🧹 Rimozione conflitto per sblocco 422", "sha": sha_file, "branch": "main"}
            requests.delete(url_git, json=payload_delete, headers=headers_git, timeout=5)
            if "sha" in payload_git: del payload_git["sha"]
            risposta_put = requests.put(url_git, json=payload_git, headers=headers_git, timeout=5)
            
        if risposta_put.status_code == 200 or risposta_put.status_code == 201:
            st.toast("✅ File Excel aggiornato stabilmente su GitHub!", icon="💾")
            return True
        else:
            st.error(f"❌ Rifiuto Scrittura GitHub. Stato: {risposta_put.status_code}")
            return False
    except Exception as e_err:
        st.error(f"💥 Errore Interno durante il salvataggio: {str(e_err)}")
        return False


# =====================================================================================
# BLOCCO 3: AUTENTICAZIONE E GESTIONE CREDENZIALI DINAMICHE DA EXCEL
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
                st.session_state.user_ruolo = str(utente_valido["RUOLO"].values[0]).strip().lower()
                st.rerun()
            else:
                st.error("❌ Credenziali errate. Riprova.")
    st.stop()

esecutore_nome = st.session_state.user_nome
esecutore_email = st.session_state.user_email
esecutore_ruolo = st.session_state.user_ruolo

st.markdown("<h1>🛡️ SATELLITE FERIE GESTORI</h1>", unsafe_allow_html=True)
st.markdown(f"<div class='user-badge'>👤 {esecutore_nome} ({esecutore_email}) — Ruolo: {esecutore_ruolo.upper()}</div>", unsafe_allow_html=True)
# =====================================================================================
# BLOCCO 4: MOTORE NOTIFICA EMAIL SMTP GOOGLE CON CONNESSA DI RETE ORIGINARIA
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
        server = smtplib.SMTP_SSL('://gmail.com', 465, timeout=10)
        server.login(EMAIL_MITTENTE_GMAIL, pass_gmail)
        server.sendmail(EMAIL_MITTENTE_GMAIL, lista_m, msg.as_string())
        server.quit()
        return True, "OK"
    except Exception as e:
        return False, str(e)
# =====================================================================================
# BLOCCO 5: MODULO DI COMPILAZIONE (FORM CENTRALE)
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
# BLOCCO 6: CONTROLLO CRONOLOGIE, REALE INVIO RETE E CRUSCOTTO ADMIN ESTESO
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
            if str(row.get("NOME_LOCALE", "")).strip() == str(scelta_pvd).strip():
                try:
                    data_inizio_estratta = str(row["INIZIO_FERIE"]).split(" ")
                    data_fine_estratta = str(row["FINE_FERIE"]).split(" ")
                    old_inizio = datetime.strptime(data_inizio_estratta[0], "%d-%m-%Y").date()
                    old_fine = datetime.strptime(data_fine_estratta[0], "%d-%m-%Y").date()
                    if (new_inizio <= old_fine) and (new_fine >= old_inizio):
                        sovrapposizione_rilevata, riga_conflitto_idx = True, idx
                        dettagli_conflitto = f"Dal {row['INIZIO_FERIE']} al {row['FINE_FERIE']} (Inserito da: {row.get('TECNICO_INSERIMENTO', 'Tecnico')})"
                        break
                except Exception: continue

        if sovrapposizione_rilevata and not forza_sovrascrittura:
            st.error(f"⚠️ ATTENZIONE: Questo locale risulta già chiuso nel periodo richiesto!\n\n📌 **Periodo registrato:** {dettagli_conflitto}.\n\nSe desideri modificare o aggiornare questo periodo con le nuove date, spunta la casella di conferma in fondo al modulo e primi di nuovo il pulsante di invio.")
        else:
            str_c = f"{data_chiusura.strftime('%d-%m-%Y')} {ora_chiusura.strftime('%H:%M')}"
            str_r = f"{data_riapertura.strftime('%d-%m-%Y')} {ora_riapertura.strftime('%H:%M')}"
            
            testo_pvd = str(scelta_pvd)
            
            # Scomposizione sicura delle stringhe basata su elementi fissi di testo
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
            
            nuova = {
                "DATA_INSERIMENTO": datetime.now().strftime("%d-%m-%Y %H:%M:%S"), 
                "TECNICO_INSERIMENTO": esecutore_nome, 
                "CODICE_LOCALE": codice_estratto,
                "NOME_LOCALE": nome_puro_locale,
                "CONCESSIONARIO": concessionario_estratto,
                "INIZIO_FERIE": str_c,   
                "FINE_FERIE": str_r,     
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
            else:
                nuova["STATO_INVIO"] = f"Errore Mail: {risposta_server}"
                
            if sovrapposizione_rilevata and riga_conflitto_idx is not None:
                st.session_state.storico_cloud.pop(riga_conflitto_idx)
            st.session_state.storico_cloud.append(nuova)
            
            df_salva = pd.DataFrame(st.session_state.storico_cloud)
            
            # Scrittura forzata ed immediata via API sul file di GitHub principale
            if push_excel_su_github(df_salva):
                st.success("✅ OPERAZIONE COMPLETATA!\n\n📧 Registro allineato su GitHub e notifica mail inviata.")
                st.session_state.form_id += 1
                time.sleep(1.5)
                st.rerun()
            else:
                st.error("❌ Errore critico durante la sincronizzazione dell'Excel su GitHub. Verifica i permessi.")

st.markdown("---")
st.markdown("### 📅 Promemoria Giri Logistici (Preavviso 3 Giorni)")
oggi = datetime.now().date()
alert_c, alert_r = [], []
for row in st.session_state.storico_cloud:
    try:
        data_inizio_estratta = str(row["INIZIO_FERIE"]).split(" ")
        data_fine_estratta = str(row["FINE_FERIE"]).split(" ")
        d_i = datetime.strptime(data_inizio_estratta[0], "%d-%m-%Y").date()
        d_f = datetime.strptime(data_fine_estratta[0], "%d-%m-%Y").date()
        if d_i - oggi == timedelta(days=3): alert_c.append(f"⚠️ **{row.get('NOME_LOCALE', 'Locale')}** chiude tra 3 giorni")
        if d_f - oggi == timedelta(days=3): alert_r.append(f"🚚 **{row.get('NOME_LOCALE', 'Locale')}** riapre tra 3 giorni")
    except Exception: continue
for a in alert_c: st.error(a)
for r in alert_r: st.warning(r)

if st.sidebar.button("🚪 Disconnetti Account"):
    del st.session_state.autenticato
    st.rerun()

if esecutore_ruolo == "admin":
    st.markdown("<br>### 📊 Registro Storico Chiusure Centralizzato", unsafe_allow_html=True)
    if st.session_state.storico_cloud:
        df_vis = pd.DataFrame(st.session_state.storico_cloud)
        st.dataframe(df_vis, hide_index=True)
        
        with io.BytesIO() as buffer:
            df_vis.to_excel(buffer, index=False, engine='openpyxl')
            st.download_button(label="📥 Scarica Registro Excel Storico", data=buffer.getvalue(), file_name="storico_ferie.xlsx", mime="application/vnd.ms-excel")
            
        st.markdown("---")
        st.markdown("### 🏢 Locali SNAITECH da inserire a sistema")
        # 🛡️ FILTRO ESTENDED ATTIVO: Isola i record se la parola Snai compare in qualsiasi campo
        righe_snaitech = []
        for row in st.session_state.storico_cloud:
            testo_unito = (str(row.get("CONCESSIONARIO", "")) + " " + str(row.get("NOME_LOCALE", ""))).lower()
            if "snaitech" in testo_unito or "snai" in testo_unito:
                righe_snaitech.append(row)
                
        if righe_snaitech:
            df_snai = pd.DataFrame(righe_snaitech)
            st.dataframe(df_snai[["CODICE_LOCALE", "NOME_LOCALE", "INIZIO_FERIE", "FINE_FERIE", "TECNICO_INSERIMENTO"]], hide_index=True)
        else:
            st.write("✅ Nessuna chiusura attiva per locali Snaitech.")
            
        st.markdown("---")
        st.markdown("### 🗑️ Cancella un Periodo Registrato")
        opzioni_cancellazione = ["- Seleziona la riga da eliminare -"]
        for idx, row in enumerate(st.session_state.storico_cloud):
            opzioni_cancellazione.append(f"ID {idx} | {row.get('CODICE_LOCALE', '')} - {row.get('NOME_LOCALE', '')} (Dal {row['INIZIO_FERIE']} al {row['FINE_FERIE']})")
            
        selezione_delete = st.selectbox("Scegli la chiusura da eliminare dal database:", opzioni_cancellazione)
        if selezione_delete != "- Seleziona la riga da eliminare -":
            try:
                parti_stringa = selezione_delete.split("ID ")
                pezzo_numerico = parti_stringa[1].split(" |")
                idx_da_eliminare = int(pezzo_numerico[0])
                
                if st.button("❌ ELIMINA DEFINITIVAMENTE QUESTA CHIUSURA"):
                    with st.spinner("Rimozione e riallineamento database..."):
                        st.session_state.storico_cloud.pop(idx_da_eliminare)
                        df_nuovo_salva = pd.DataFrame(st.session_state.storico_cloud)
                        push_excel_su_github(df_nuovo_salva)
                    st.success("🗑️ Chiusura eliminata con successo!")
                    time.sleep(1)
                    st.rerun()
            except Exception: pass
            
        st.markdown("---")
        st.markdown("### 📤 Ricarica Registro Excel Aggiornato dall'Ufficio")
        file_caricato = st.file_uploader("Trascina qui il file storico_ferie.xlsx modificato per caricare i dati nel portale:", type=["xlsx"])
        if file_caricato is not None:
            try:
                df_caricato = pd.read_excel(file_caricato).fillna("")
                if "CODICE_LOCALE" in df_caricato.columns and "INIZIO_FERIE" in df_caricato.columns:
                    if st.button("🔄 CONFERMA E SOVRASCRIVI DATABASE CON QUESTO FILE"):
                        st.session_state.storico_cloud = df_caricato.to_dict('records')
                        push_excel_su_github(df_caricato)
                        st.success("✅ Database sovrascritto con successo!")
                        time.sleep(1.5)
                        st.rerun()
                else:
                    st.error("❌ Struttura file non valida. Controlla i nomi delle colonne.")
            except Exception as e_load:
                st.error(f"❌ Errore durante la lettura del file: {str(e_load)}")
    else:
                st.write("Nessuna chiusura presente nel registro.")
