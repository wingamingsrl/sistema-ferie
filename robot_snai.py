import os
import io
import time
import pyotp
import requests
import pandas as pd
from datetime import datetime
from playwright.sync_api import sync_playwright

CHIAVE_SEGRETA_2FA = "FTIA6UQZM2LQLPYJ"
SNAI_USER = "2141ManuelaA"
SNAI_PASS = "Salmi123!"

def preleva_storico_diretto_da_cloud():
    print("📡 [Robot] STEP 1: Lettura del database Excel locale...")
    try:
        nome_file_locale = "storico_ferie.xlsx"
        if os.path.exists(nome_file_locale):
            return pd.read_excel(nome_file_locale).fillna("")
    except Exception: pass
    return pd.DataFrame()

def genera_codice_otp_automatico():
    chiave_pulita = CHIAVE_SEGRETA_2FA.strip().upper().replace(" ", "")
    totp = pyotp.TOTP(chiave_pulita)
    return totp.now()

def scarica_e_aggiorna_excel_su_github(codice_locale_successo):
    try:
        print(f"💾 [Cloud Excel] Allineamento riuscito. Svuoto ROBOT_ACTION per: {codice_locale_successo}...")
        t_git = os.environ.get("TOKEN_GITHUB_ACTIONS", "")
        if not t_git:
            try: t_git = str(pd.read_excel("token.xlsx").iloc).strip()
            except Exception: return
            
        url_git = f"https://github.com{FILE_STORICO_PERMANENTE}"
        headers_git = {"Authorization": f"token {t_git}", "Accept": "application/vnd.github+json"}
        
        if os.path.exists(FILE_STORICO_PERMANENTE):
            df_file = pd.read_excel(FILE_STORICO_PERMANENTE)
            df_file.loc[df_file["CODICE_LOCALE"].astype(str).str.strip() == str(codice_locale_successo).strip(), "ROBOT_ACTION"] = ""
            df_file.to_excel(FILE_STORICO_PERMANENTE, index=False)
            
            with open(FILE_STORICO_PERMANENTE, "rb") as f_in:
                dati_b64 = base64.b64encode(f_in.read()).decode('utf-8')
                
            res_get = requests.get(url_git, headers=headers_git, timeout=5)
            sha_file = res_get.json().get("sha", "") if res_get.status_code == 200 else ""
            
            payload_git = {
                "message": f"🤖 [Robot] Allineamento Snaitech OK. Reset azione locale {codice_locale_successo}", 
                "content": dati_b64, "branch": "main"
            }
            if sha_file: payload_git["sha"] = sha_file
            requests.put(url_git, json=payload_git, headers=headers_git, timeout=5)
            print("   ✅ [Cloud Excel] Database ripulito e sincronizzato con successo su GitHub!")
    except Exception: pass


def avvia_sincronizzazione_automatica():
    df_ferie = preleva_storico_diretto_da_cloud()
    if df_ferie.empty: return

    df_snai = df_ferie[
        df_ferie["CONCESSIONARIO"].astype(str).str.strip() == "Snaitech Spa WG"
    ]
    if df_snai.empty: return

    print(f"🤖 [Robot] STEP 3: Rilevati {len(df_snai)} locali Snaitech Spa WG. Avvio Chrome...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=[
            "--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"
        ]) 
        context = browser.new_context()
        page = context.new_page()

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

            print("📬 [Robot] STEP 6: Spostamento sulla pagina degli Esercizi censiti...")
            page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx")
            print("   ⏳ [Robot] STEP 6a: Attesa stabilizzazione della pagina (10 secondi)...")
            time.sleep(10)

            for _, row in df_snai.iterrows():
                try:
                    codice_aams = str(row["CODICE_LOCALE"]).strip()
                    nome_locale_corrente = str(row["NOME_LOCALE"]).strip()
                    data_in_completa = str(row["INIZIO_FERIE"]).strip()
                    data_fi_completa = str(row["FINE_FERIE"]).strip()
                    mirino_azione = str(row.get("ROBOT_ACTION", "")).strip().upper()
                    
                    data_inizio_pulita = str(data_in_completa).replace("-", "/").strip()
                    data_fine_pulita = str(data_fi_completa).replace("-", "/").strip()
                    
                    # 🛡️ INTERCETTATORE DI MANUELA: Se la cella non contiene NUOVA o MODIFICA, salta la riga all'istante
                    if mirino_azione not in ["NUOVA", "MODIFICA"]:
                        print(f"⏩ [Robot] Locale {codice_aams} - {nome_locale_corrente}: Nessuna azione richiesta. Salto riga.")
                        continue
                        
                    print(f"🚀 [Robot] STEP 7: Avvio lavorazione ({mirino_azione}) -> Codice Locale: {codice_aams} - {nome_locale_corrente}")

                    # 🛡️ TUO CODICE NATIVO ORIGINALE DEI RAGAZZI AL 100% — MAI PIÙ TOCCATO
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

                    icona_nuovo = target_frame.locator("img[src*='insert_pianificazione.jpg'], img[src*='insert_pianificazione'], img[id*='img_pianificazione']").first
                    icona_modifica = target_frame.locator("img[src*='edit_pianificazione']").first
                    cella_td = target_frame.locator("td[onclick*='Pianificazione']").first
                    
                    if icona_modifica.count() > 0:
                        print("   📝 [Robot] STEP 8: [MODIFICA] Rilevato cambio URL ChiusuraEsercizio.aspx. Clicco...")
                        icona_modifica.click(force=True, timeout=8000)
                    elif icona_nuovo.count() > 0:
                        print("   🟢 [Robot] STEP 8a: [NUOVA CHIUSURA] Clic sul pallino verde...")
                        icona_nuovo.click(force=True, timeout=8000)
                    else:
                        print("   AM 🖱️ [Grid Mode] Clic sulla cella td nativa della riga...")
                        cella_td.click(force=True, timeout=8000)
                    
                    print("   ⏳ [Robot] STEP 8c: Attesa apertura campi date (7 secondi)...")
                    time.sleep(7)

# =====================================================================================
# BLOCCO 5: AGGIORNAMENTO AUTOMATICO VIA JS CON ID RETTIFICATO E RESET REALE EXCEL
# =====================================================================================
                    # 🛡️ DIGITAZIONE REALE DI MANUELA: Sveglia i validatori di Snaitech inserendo le date tasto per tasto
                    campo_dal = frame_date.locator("#ctl00_Cp1_Txtiniziochiusura, input[id*='Txtiniziochiusura']").first
                    campo_al = frame_date.locator("#ctl00_Cp1_txtfinechiusura, input[id*='txtfinechiusura']").first

                    print("   📝 [Robot] STEP 9: Digitazione sequenziale data inizio...")
                    campo_dal.wait_for(state="visible", timeout=12000)
                    campo_dal.click()
                    campo_dal.press("Control+A")
                    campo_dal.press("Backspace")
                    time.sleep(1)
                    campo_dal.press_sequentially(data_inizio_pulita, delay=100)
                    time.sleep(1)
                    campo_dal.press("Tab")
                    time.sleep(1)
                    
                    try: frame_date.locator("#ctl00_Cp1_fascia_from").select_option("00:00")
                    except Exception: pass
                    time.sleep(1)
                    
                    print("   📝 [Robot] STEP 9a: Digitazione sequenziale data fine...")
                    campo_al.click()
                    campo_al.press("Control+A")
                    campo_al.press("Backspace")
                    time.sleep(1)
                    campo_al.press_sequentially(data_fine_pulita, delay=100)
                    time.sleep(1)
                    campo_al.press("Tab")
                    time.sleep(1)
                    
                    try: frame_date.locator("#ctl00_Cp1_fascia_to").select_option("23:30")
                    except Exception: pass
                    time.sleep(2)


                    print("   💾 [Robot] STEP 10: Invio moduli di chiusura a Snaitech (Clic su Tasto Salva)...")
                    frame_date.locator("#ctl00_Cp1_BtnOk").first.click(timeout=10000)
                    print(f"   ✅ [Robot] STEP 11: Invio completato. Pausa di stabilizzazione di 8 secondi...")
                    time.sleep(4)
                    
                    # 🛡️ PULIZIA AUTOMATICA EXCEL: Svuota la parola dall'Excel e aggiorna GitHub ad inserimento riuscito
                    scarica_e_aggiorna_excel_su_github(codice_aams)
                    time.sleep(4)
                    
                    page.goto("https://partner.snai.it")
                    time.sleep(6)
                    
                except Exception as row_err:
                    print(f"   ⚠️ Nota compilazione: Scavalco riga. Errore: {str(row_err)}")
                    try:
                        page.goto("https://partner.snai.it")
                        time.sleep(6)
                    except Exception: pass
                    continue

            print("🔒 [Robot] STEP 12: Chiusura sessione formale (Logout di sicurezza)...")
            try: page.locator("a:has-text('LogOut'), a:has-text('Esci'), [id*='btnLogOut']").first.click(timeout=8000)
            except Exception: page.context.clear_cookies()

        except Exception as e: print(f"❌ Errore durante la navigazione sul portale partner.snai.it: {str(e)}")
        finally: browser.close()

if __name__ == "__main__":
    avvia_sincronizzazione_automatica()
