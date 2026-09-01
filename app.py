# =====================================================================================
# BLOCCO 2: COLLEGAMENTO FILE EXCEL PERMANENTI E FUNZIONI DI SCRITTURA SU GITHUB
# VERSIONE DI PRODUZIONE - PROTEZIONE CRASH IN CASO DI FILE EXCEL VUOTO O 0 BYTE
# =====================================================================================
FILE_LOCALI = "elenco_locali.xlsx"
FILE_TECNICI = "elenco_tecnici.xlsx"
FILE_STORICO_PERMANENTE = "storico_ferie.xlsx"

EMAIL_MITTENTE_GMAIL = "wingamingsrl@gmail.com"
EMAIL_MANUELA_RICEVENTE = "manuela.arigoni@wingaming.it"

def scarica_file_da_github_se_esiste(nome_file):
    try:
        t_git = str(st.secrets["github"]["token_accesso"]).strip()
        url_lettura = f"https://github.com{nome_file}?t={int(time.time())}"
        
        h = {
            "Authorization": f"Bearer {t_git}", 
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "WinGaming-Cloud-App"
        }
        r = requests.get(url_lettura, headers=h, timeout=5)
        if r.status_code == 200:
            b64_content = r.json().get("content", "")
            # Legge l'Excel in memoria RAM
            return pd.read_excel(io.BytesIO(base64.b64decode(b64_content)))
    except Exception:
        pass
    return None

def carica_database_locale():
    df_l = pd.read_excel(FILE_LOCALI).fillna("") if os.path.exists(FILE_LOCALI) else pd.DataFrame(columns=["CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO"])
    df_t = pd.read_excel(FILE_TECNICI).fillna("") if os.path.exists(FILE_TECNICI) else pd.DataFrame(columns=["NOME", "EMAIL", "PASSWORD", "RUOLO"])
    
    # 🛡️ GESTIONE ANTICRASH: Tenta lo scaricamento sicuro da GitHub
    df_s = scarica_file_da_github_se_esiste(FILE_STORICO_PERMANENTE)
    
    # Se il file su GitHub è vuoto, dà errore o non esiste, creiamo la tabella vuota con le colonne corrette
    if df_s is None or df_s.empty:
        try:
            if os.path.exists(FILE_STORICO_PERMANENTE):
                df_s = pd.read_excel(FILE_STORICO_PERMANENTE).fillna("")
            else:
                df_s = pd.DataFrame(columns=["DATA_INSERIMENTO", "TECNICO", "LOCALE", "INIZIO_FERIE", "FINE_FERIE", "COPIA_PROMEMORIA"])
        except Exception:
            # Ruota di scorta estrema: se anche il file locale sul server dà errore (perché vuoto), genera la struttura pulita
            df_s = pd.DataFrame(columns=["DATA_INSERIMENTO", "TECNICO", "LOCALE", "INIZIO_FERIE", "FINE_FERIE", "COPIA_PROMEMORIA"])
            
    return df_l, df_t, df_s.fillna("")

df_locali, df_tecnici, df_storico_file = carica_database_locale()

# Sincronizzazione immediata della memoria RAM di Streamlit
st.session_state.storico_cloud = df_storico_file.to_dict('records')

def push_excel_su_github(df_da_salvare):
    try:
        t_git = str(st.secrets["github"]["token_accesso"]).strip()
        url_git = f"https://github.com{FILE_STORICO_PERMANENTE}"
        
        output_binario = io.BytesIO()
        with pd.ExcelWriter(output_binario, engine='openpyxl') as writer:
            df_da_salvare.to_excel(writer, index=False)
        dati_base64 = base64.b64encode(output_binario.getvalue()).decode('utf-8')
        
        headers_git = {
            "Authorization": f"Bearer {t_git}", 
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
        
        if risposta_put.status_code == 200 or risposta_put.status_code == 201:
            return True
        else:
            st.toast(f"⚠️ Errore di Scrittura GitHub: {risposta_put.status_code}.", icon="❌")
            return False
    except Exception as e_err:
        st.toast(f"⚠️ Errore di Sistema Interno: {str(e_err)}", icon="❌")
        return False
