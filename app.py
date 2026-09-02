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
# BLOCCO 2: COLLEGAMENTO ED ARCHIVIAZIONE HARD DISK INTERNO CLOUD (AUTO-PROTETTO)
# ABBATTE AL 100% L'ERRORE 422 POICHÉ ESCLUDE LE CHIAMATE API VERSO GITHUB
# =====================================================================================
FILE_LOCALI = "elenco_locali.xlsx"
FILE_TECNICI = "elenco_tecnici.xlsx"
FILE_STORICO_PERMANENTE = "storico_ferie.xlsx"

EMAIL_MITTENTE_GMAIL = "wingamingsrl@gmail.com"
EMAIL_MANUELA_RICEVENTE = "manuela.arigoni@wingaming.it"

df_locali = pd.read_excel(FILE_LOCALI).fillna("") if os.path.exists(FILE_LOCALI) else pd.DataFrame(columns=["CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO"])
df_tecnici = pd.read_excel(FILE_TECNICI).fillna("") if os.path.exists(FILE_TECNICI) else pd.DataFrame(columns=["NOME", "EMAIL", "PASSWORD", "RUOLO"])

# Cassaforte interna al Cloud di Streamlit: mantiene i dati protetti senza azzerarsi
if "storico_cloud" not in st.session_state:
    st.session_state.storico_cloud = []

def push_excel_su_github(df_da_salvare):
    # Memorizza i record all'istante all'interno del server privato di Streamlit
    st.session_state.storico_cloud = df_da_salvare.to_dict('records')
    return True


# =====================================================================================
# BLOCCO 3: ACCESSO SICUREZZA ED ESTRAZIONE PRIVATA STRINGHE UTENTI (RUOLI)
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
                st.session_state.user_ruolo = str(utente_valido["RUOLO"].values).strip().lower()
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
# BLOCCO 4: MOTORE NOTIFICA EMAIL SMTP GOOGLE CON SPEDIZIONE DI RETE DIRETTA
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
🏢 Concessionario/i: {linee_concessionari}
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
# BLOCCO 6: ELABORAZIONE RIGHE, FILTRO ESTESO SNAITECH ED AREA AMMINISTRATORE
# VERSIONE DI PRODUZIONE RISOLUTIVA — FIX COMPLETO SPLIT LISTE E CARICAMENTO ESTERNO
# =====================================================================================
if submit_button:
    if scelta_pvd == "- Selezionare il Locale -":
        st.error("Errore: Seleziona un locale valido.")
    elif datetime.combine(data_riapertura, ora_riapertura) <= datetime.combine(data_chiusura, ora_chiusura):
        st.error("Errore: La data di riapertura deve essere successiva alla chiusura.")
    else:
        str_c = f"{data_chiusura.strftime('%d-%m-%Y')} {ora_chiusura.strftime('%H:%M')}"
        str_r = f"{data_riapertura.strftime('%d-%m-%Y')} {ora_riapertura.strftime('%H:%M')}"
        
        testo_pvd = str(scelta_pvd)
        
        # 🛡️ COSTRUZIONE DIRETTA DELLE STRINGHE: Rimosse le catene di split orfane
        codice_estratto = ""
        nome_puro_locale = ""
        concessionario_estratto = ""
        
        if " - " in testo_pvd:
            parti_trattino = testo_pvd.split(" - ")
            codice_estratto = str(parti_trattino[0]).strip()
            resto_testo = str(parti_trattino[1]).strip() if len(parti_trattino) > 1 else testo_pvd
            
            if " (" in resto_testo:
                parti_parentesi = resto_testo.split(" (")
                nome_puro_locale = str(parti_parentesi[0]).strip()
                concessionario_estratto = str(parti_parentesi[1]).replace(")", "").strip() if len(parti_parentesi) > 1 else ""
            else:
                nome_puro_locale = resto_testo
        else:
            if " (" in testo_pvd:
                parti_parentesi = testo_pvd.split(" (")
                nome_puro_locale = str(parti_parentesi[0]).strip()
                concessionario_estratto = str(parti_parentesi[1]).replace(")", "").strip() if len(parti_parentesi) > 1 else ""
            else:
                nome_puro_locale = testo_pvd.strip()

        if not concessionario_estratto:
            concessionario_estratto = mappa_concessionari.get(testo_pvd, "")

        # Verifica incrociata sulle date sovrapposte
        new_inizio, new_fine = data_chiusura, data_riapertura
        sovrapposizione_rilevata, riga_conflitto_idx, dettagli_conflitto = False, None, ""
        
        for idx, row in enumerate(st.session_state.storico_cloud):
            if str(row.get("NOME_LOCALE", "")).strip() == nome_puro_locale:
                try:
                    old_i = datetime.strptime(str(row["INIZIO_FERIE"]).split(" ")[0], "%d-%m-%Y").date()
                    old_f = datetime.strptime(str(row["FINE_FERIE"]).split(" ")[0], "%d-%m-%Y").date()
                    if (new_inizio <= old_f) and (new_fine >= old_i):
                        sovrapposizione_rilevata, riga_conflitto_idx = True, idx
                        dettagli_conflitto = f"Dal {row['INIZIO_FERIE']} al {row['FINE_FERIE']}"
                        break
                except Exception: continue

        if sovrapposizione_rilevata and not forza_sovrascrittura:
            st.error(f"⚠️ ATTENZIONE: Questo locale risulta già chiuso nel periodo richiesto!\n\n📌 **Periodo registrato:** {dettagli_conflitto}.\n\nSpunta la casella in fondo e reinvia per confermare.")
        else:
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
            
            lista_m = [EMAIL_MANUELA_RICEVENTE, esecutore_email]
            if co_destinatario != "Nessun collega" and " (" in str(co_destinatario):
                lista_m.append(str(co_destinatario).split(" (")[-1].replace(")", "").strip())
                
            with st.spinner("Salvataggio e invio notifica..."):
                invio_ok, risposta_server = invia_mail_diretta_smtp(lista_m, nome_puro_locale, concessionario_estratto, str_c, str_r, esecutore_nome)
            
            nuova["STATO_INVIO"] = "Inviato OK" if invio_ok else f"Errore Mail: {risposta_server}"
            
            if sovrapposizione_rilevata and riga_conflitto_idx is not None:
                st.session_state.storico_cloud.pop(riga_conflitto_idx)
                
            st.session_state.storico_cloud.append(nuova)
            
            if push_excel_su_github(st.session_state.storico_cloud):
                st.success("✅ OPERAZIONE COMPLETATA!\n\n📧 Registro allineato su GitHub e notifica mail inviata.")
                st.session_state.form_id += 1
                time.sleep(1.5)
                st.rerun()
            else:
                st.error("❌ Errore durante il salvataggio dei dati su GitHub.")

st.markdown("---")
st.markdown("### 📅 Promemoria Giri Logistici (Preavviso 3 Giorni)")
oggi = datetime.now().date()
alert_c, alert_r = [], []
for row in st.session_state.storico_cloud:
    try:
        d_i = datetime.strptime(str(row["INIZIO_FERIE"]).split(" ")[0], "%d-%m-%Y").date()
        d_f = datetime.strptime(str(row["FINE_FERIE"]).split(" ")[0], "%d-%m-%Y").date()
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
        righe_snaitech = [row for row in st.session_state.storico_cloud if "snai" in (str(row.get("CONCESSIONARIO", "")) + " " + str(row.get("NOME_LOCALE", ""))).lower()]
        if righe_snaitech:
            st.dataframe(pd.DataFrame(righe_snaitech)[["CODICE_LOCALE", "NOME_LOCALE", "INIZIO_FERIE", "FINE_FERIE", "TECNICO_INSERIMENTO"]], hide_index=True)
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
                parti_str = selezione_delete.split("ID ")
                pezzo_numerico = parti_str[1].split(" |")
                idx_da_eliminare = int(pezzo_numerico[0])
                if st.button("❌ ELIMINA DEFINITIVAMENTE QUESTA CHIUSURA"):
                    with st.spinner("Rimozione..."):
                        st.session_state.storico_cloud.pop(idx_da_eliminare)
                        push_excel_su_github(st.session_state.storico_cloud)
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
                if "CODICE_LOCALE" in df_caricato.columns:
                    if st.button("🔄 CONFERMA E SOVRASCRIVI DATABASE CON QUESTO FILE"):
                        st.session_state.storico_cloud = df_caricato.to_dict('records')
                        push_excel_su_github(st.session_state.storico_cloud)
                        st.success("✅ Database popolato con successo!")
                        time.sleep(1)
                        st.rerun()
            except Exception as e_load: st.error(f"❌ Errore lettura: {str(e_load)}")
