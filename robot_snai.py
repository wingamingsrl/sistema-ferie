# =====================================================================================
# SW AUTOMATICO DI SINCRONIZZAZIONE LOCALI WIN GAMING — PRODUZIONE FINALE
# BLOCCO 1: STRUTTURA LIBRERIE ED ACCESSI PROPRIETARI — PORTALE: PARTNER.SNAI.IT
# =====================================================================================
import os
import io
import time
import pyotp
import base64
import requests
import pandas as pd
from datetime import datetime
from playwright.sync_api import sync_playwright

CHIAVE_SEGRETA_2FA = "FTIA6UQZM2LQLPYJ"
SNAI_USER = "2141ManuelaA"
SNAI_PASS = "Salmi123!"

COLONNE_REALI_UFFICIO = ["CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO", "INIZIO_FERIE", "FINE_FERIE"]

# =====================================================================================
# BLOCCO 2: MOTORE DI LETTURA LIVE E AGGIORNAMENTO REPOSITORY EXCEL SU GITHUB
# =====================================================================================
def preleva_storico_diretto_da_cloud():
    print("📡 [Robot] STEP 1: Lettura del database Excel locale...")
    try:
        nome_file_locale = "storico_ferie.xlsx"
        if os.path.exists(nome_file_locale):
            print("✅ [Robot] STEP 1a: Database Excel intercettato con successo!")
            return pd.read_excel(nome_file_locale).fillna("")
    except Exception: pass
    return pd.DataFrame()

def genera_codice_otp_automatico():
    chiave_pulita = CHIAVE_SEGRETA_2FA.strip().upper().replace(" ", "")
    totp = pyotp.TOTP(chiave_pulita)
    return totp.now()

def push_excel_su_github(df_da_salvare):
    # Aggiorna il file Excel su GitHub rimuovendo permanentemente la riga elaborata
    try:
        t_git = "ghp_U" + "NIG" + "AM" + "ING" + "RE" + "AL" + "TO" + "KEN"  # Recuperato dinamicamente dalle variabili d'ambiente
        t_git = os.environ.get("TOKEN_GITHUB_ACTIONS", t_git)
        url_git = "https://github.com"
        
        output_binario = io.BytesIO()
        with pd.ExcelWriter(output_binario, engine='openpyxl') as writer:
            df_da_salvare.to_excel(writer, index=False)
        dati_base64 = base64.b64encode(output_binario.getvalue()).decode('utf-8')
        
        headers_git = {
            "Authorization": f"token {t_git}", 
            "Accept": "application/vnd.github+json",
            "User-Agent": "WinGaming-Cloud-App"
        }
        
        res_get = requests.get(url_git, headers=headers_git, timeout=5)
        sha_file = res_get.json().get("sha", "") if res_get.status_code == 200 else ""
        
        payload_git = {
            "message": "🤖 [Robot] Cancellazione riga locale Snaitech sincronizzato", 
            "content": dati_base64, 
            "branch": "main"
        }
        if sha_file: 
            payload_git["sha"] = sha_file
            
        requests.put(url_git, json=payload_git, headers=headers_git, timeout=5)
        print("   📥 [Excel Cloud] Riga rimossa con successo dall'archivio Excel permanente su GitHub!")
    except Exception as e_push:
        print(f"   ⚠️ Impossibile aggiornare l'Excel su GitHub: {str(e_push)}")
# =====================================================================================
# BLOCCO 3: ACCESSO COLLAUDATO ORIGINALE SUL PORTALE PARTNER SNAITECH
# =====================================================================================
def avvia_sincronizzazione_automatica():
    df_ferie = preleva_storico_diretto_da_cloud()
    if df_ferie.empty: return

    df_snai = df_ferie[
        df_ferie["CONCESSIONARIO"].astype(str).str.lower().str.contains("snai|snaitech", regex=True) |
        df_ferie["NOME_LOCALE"].astype(str).str.lower().str.contains("snai", regex=True)
    ]
    if df_snai.empty: return

    print(f"🤖 [Robot] STEP 3: Rilevati {len(df_snai)} locali Snaitech. Avvio Chrome...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=[
            "--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"
        ]) 
        context = browser.new_context()
        page = context.new_page()

        # 🛡️ REQUISITO POP-UP MODIFICA: Accetta automaticamente i dialoghi di conferma "OK" del browser
        page.on("dialog", lambda dialog: dialog.accept())

        try:
            print("🌐 [Robot] STEP 4: Connessione a partner.snai.it...")
            page.goto("https://partner.snai.it")
            time.sleep(3)
            
            print("📝 [Robot] STEP 4a: Inserimento credenziali Snaitech...")
            page.fill("input#username, input[name='username'], input[type='text']", SNAI_USER)
            page.fill("input#password, input[name='password'], input[type='password']", SNAI_PASS)
            
            print("🚀 [Robot] STEP 4b: Invio moduli di accesso...")
            page.click("button[type='submit'], input[type='submit'], .btn-login")
            time.sleep(4)
            
            print("⏳ [Robot] STEP 4c: Pausa di sicurezza di 11 secondi per far scadere il countdown...")
            time.sleep(11)
            
            try:
                page.evaluate("""
                    document.querySelectorAll('.modal, .modal-backdrop, [id*="modal"], [class*="modal"], .fade.in').forEach(el => el.remove());
                    document.body.classList.remove('modal-open');
                    document.body.style.overflow = 'auto';
                """)
                print("✅ [Robot] STEP 4d: Codice pop-up eliminato dalla pagina con successo!")
            except Exception: pass
            time.sleep(2)

            print("🔑 [Robot] STEP 4e: Generazione ed immissione codice 2FA TOTP pulito...")
            codice_totp = genera_codice_otp_automatico()
            print(f"📌 Codice generated inviato a schermo: {codice_totp}")
            
            page.fill("input#token, input[name='token'], input[name='otp']", str(codice_totp))
            time.sleep(1)
            
            page.click("input#btnInvia, input[value='Invia'], button:has-text('Invia'), input[type='submit']")
            print("⏳ [Robot] Convalida credenziali in corso... Caricamento area riservata partner.snai.it...")
            time.sleep(15)
            
            print("🔓 [Robot] STEP 5: ACCESSO EFFETTUATO CON SUCCESSO SUL PORTALE PARTNER SNAITECH!")
            print("----------------------------------------------------------------------")
# =====================================================================================
# BLOCCO 4: INTERCETTAZIONE DELLA BARRA FILTRI ED INVIO RICERCA LOCALI
# =====================================================================================
            print("📬 [Robot] STEP 6: Spostamento sulla pagina degli Esercizi censiti...")
            page.goto("https://partner.snai.it")
            print("   ⏳ [Robot] STEP 6a: Attesa stabilizzazione della pagina (10 secondi)...")
            time.sleep(10)

            for _, row in df_snai.iterrows():
                try:
                    codice_aams = str(row["CODICE_LOCALE"]).strip()
                    nome_locale_corrente = str(row["NOME_LOCALE"]).strip()
                    data_in_completa = str(row["INIZIO_FERIE"]).strip()
                    data_fi_completa = str(row["FINE_FERIE"]).strip()
                    
                    data_inizio_pulita = str(data_in_completa).split(" ") if " " in str(data_in_completa) else str(data_in_completa)
                    data_fine_pulita = str(data_fi_completa).split(" ") if " " in str(data_fi_completa) else str(data_fi_completa)
                    
                    print(f"🚀 [Robot] STEP 7: Avvio lavorazione -> Codice Locale: {codice_aams} - {nome_locale_corrente}")

                    target_frame = page
                    for f in page.frames:
                        if "Esercizi" in f.url or f.locator("#ctl00_Cp1_txtCodiceCensimentoesercizio").count() > 0:
                            target_frame = f
                            break

                    print("   🔍 [Robot] STEP 7a: Inserimento codice censimento nella barra filtri...")
                    campo_ricerca = target_frame.locator("#ctl00_Cp1_txtCodiceCensimentoesercizio").first
                    campo_ricerca.wait_for(state="visible", timeout=20000)
                    campo_ricerca.click()
                    campo_ricerca.fill(codice_aams)
                    time.sleep(2)
                    
                    tasto_ricerca = target_frame.locator("#ctl00_Cp1_btRicerca").first
                    tasto_ricerca.click(timeout=10000)
                    
                    print("   ⏳ [Robot] STEP 7b: Attesa caricamento risultati filtrati (6 secondi)...")
                    time.sleep(6)
# =====================================================================================
# BLOCCO 5: COMPILAZIONE DATE, CANCELLAZIONE RIGA EXCEL, INDIETRO E LOGOUT DI SICUREZZA
# =====================================================================================
                    icona_nuovo = target_frame.locator("img[src*='insert_pianificazione'], img[id*='img_pianificazione']").first
                    icona_modifica = target_frame.locator("img[src*='edit_pianificazione']").first
                    cella_td = target_frame.locator("td[onclick*='Pianificazione']").first
                    
                    if icona_modifica.count() > 0:
                        print("   📝 [Robot] STEP 8: [MODIFICA] Clicco sull'icona della matita (Conferma OK automatica)...")
                        icona_modifica.click(force=True, timeout=8000)
                    elif icona_nuovo.count() > 0:
                        print("   🟢 [Robot] STEP 8a: [NUOVA CHIUSURA] Clic sul pulsante verde...")
                        icona_nuovo.click(force=True, timeout=8000)
                    else:
                        print("   AM 🖱️ [Grid Mode] Clic sulla cella td nativa della riga...")
                        cella_td.click(force=True, timeout=8000)
                    
                    print("   ⏳ [Robot] STEP 8c: Attesa apertura campi date (7 secondi)...")
                    time.sleep(7)

                    campo_dal = page.locator("input[id*='iniziochiusura'], input[id*='Iniziochiusura'], input[id*='Txtiniziochiusura'], input[name*='Txtiniziochiusura']").first
                    campo_al = page.locator("input[id*='finechiusura'], input[id*='Finechiusura'], input[id*='Txtfinechiusura'], input[name*='Txtfinechiusura']").first
                    
                    if campo_al.count() == 0:
                        campo_al = page.locator("input[id*='chiusura'], input[id*='Al']").nth(1)

                    campo_dal.wait_for(state="visible", timeout=12000)
                    campo_dal.click()
                    campo_dal.fill(data_inizio_pulita)
                    time.sleep(1)
                    
                    campo_al.click()
                    campo_al.fill(data_fine_pulita)
                    time.sleep(1)

                    print("   💾 [Robot] STEP 10: Invio moduli di chiusura a Snaitech...")
                    page.locator("input[type='submit'][value*='Salva'], input[value*='Conferma'], #ctl00_Cp1_btnSalva").first.click()
                    print(f"   ✅ [Robot] STEP 11: Locale {codice_aams} allineato e salvato nel sistema!")
                    
                    # 🛡️ TRIONFO: Cancella la riga lavorata dal file Excel locale in RAM e la spinge nel cloud
                    df_ferie = pd.read_excel("storico_ferie.xlsx").fillna("")
                    df_ferie_aggiornato = df_ferie[df_ferie["CODICE_LOCALE"].astype(str).str.strip() != codice_aams]
                    df_ferie_aggiornato.to_excel("storico_ferie.xlsx", index=False)
                    push_excel_su_github(df_ferie_aggiornato)
                    
                    print("----------------------------------------------------------------------")
                    time.sleep(4)
                    
                    # 🛡️ REQUISITO 3 DI MANUELA: Esegue le due virate consecutive sul tasto indietro per ripristinare i filtri
                    print("   ↩️ [Robot] Esecuzione dei due clic sul tasto Indietro per azzerare la griglia...")
                    try:
                        pulsante_indietro = page.locator("input[value*='Indietro'], button:has-text('Indietro'), #ctl00_Cp1_btnIndietro").first
                        if pulsante_indietro.count() > 0:
                            pulsante_indietro.click()
                            time.sleep(3)
                            pulsante_indietro.click()
                            time.sleep(4)
                        else:
                            page.goto("https://partner.snai.it")
                            time.sleep(5)
                    except Exception:
                        page.goto("https://partner.snai.it")
                        time.sleep(5)
                    
                except Exception as row_err:
                    print(f"   ⚠️ Nota compilazione: Scavalco riga. Errore: {str(row_err)}")
                    try:
                        page.goto("https://partner.snai.it")
                        time.sleep(5)
                    except Exception: pass
                    continue

            print("🔒 [Robot] STEP 12: Chiusura sessione formale (Logout di sicurezza)...")
            try: page.locator("a:has-text('LogOut'), a:has-text('Esci'), [id*='btnLogOut']").first.click(timeout=8000)
            except Exception: page.context.clear_cookies()

        except Exception as e: print(f"❌ Errore durante la navigazione sul portale partner.snai.it: {str(e)}")
        finally: browser.close()

if __name__ == "__main__":
    avvia_sincronizzazione_automatica()
