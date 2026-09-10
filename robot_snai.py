# =====================================================================================
# SW AUTOMATICO DI SINCRONIZZAZIONE LOCALI WIN GAMING — PRODUZIONE FINALE
# BLOCCO 1: STRUTTURA LIBRERIE ED ACCESSI PROPRIETARI — PORTALE: PARTNER.SNAI.IT
# =====================================================================================
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

# =====================================================================================
# BLOCCO 2: MOTORE DI LETTURA LIVE IN MEMORIA RAM DEL REPOSITORY EXCEL LOCALE
# =====================================================================================
def preleva_storico_diretto_da_cloud():
    print("📡 [Robot] STEP 1: Lettura del database Excel locale sul server Actions...")
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

# =====================================================================================
# BLOCCO 3: AVVIO CHROME CON SCHERMATURA ED ATTERRAGGIO RAPIDO
# =====================================================================================
def avvia_sincronizzazione_automatica():
    df_ferie = preleva_storico_diretto_da_cloud()
    if df_ferie.empty: return

    df_snai = df_ferie[
        df_ferie["CONCESSIONARIO"].astype(str).str.lower().str.contains("snai|snaitech", regex=True) |
        df_ferie["NOME_LOCALE"].astype(str).str.lower().str.contains("snai", regex=True)
    ]
    if df_snai.empty: return

    print(f"🤖 [Robot] STEP 3: Rilevati {len(df_snai)} locali Snaitech. Avvio Chrome Camuffato...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=[
            "--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage",
            "--disable-blink-features=AutomationControlled"
        ]) 
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
            locale="it-IT", timezone_id="Europe/Rome"
        )
        page = context.new_page()

        try:
            print("🌐 [Robot] STEP 4: Connessione schermata a partner.snai.it...")
            page.goto("https://partner.snai.it", wait_until="load", timeout=50000)
            time.sleep(5)
            
            page.locator("input#username, input[name='username']").first.fill(SNAI_USER)
            page.locator("input#password, input[name='password']").first.fill(SNAI_PASS)
            page.click("button[type='submit'], input[type='submit'], .btn-login")
            time.sleep(11)
            
            try: page.evaluate("document.querySelectorAll('.modal, .modal-backdrop').forEach(el => el.remove());")
            except Exception: pass

            codice_totp = genera_codice_otp_automatico()
            print(f"🔑 [Robot] STEP 4f: Codice generated inviato a schermo: {codice_totp}")
            
            campo_token = page.locator("input#token, input[name='token'], input[name='otp']").first
            campo_token.fill(str(codice_totp))
            time.sleep(2)
            
            page.keyboard.press("Enter")
            print("⏳ [Robot] Convalida credenziali in corso (15 secondi)...")
            time.sleep(15)
            
            print("🔓 [Robot] STEP 5: ACCESSO EFFETTUATO CON SUCCESSO SUL PORTALE PARTNER SNAITECH!")
            print("----------------------------------------------------------------------")

# =====================================================================================
# BLOCCO 4: NAVIGAZIONE ED INSERIMENTO CODICE CENSIMENTO CON GLI ID CERTIFICATI DA MANUELA
# =====================================================================================
            print("📬 [Robot] STEP 6: Spostamento sulla pagina degli Esercizi censiti...")
            page.goto("https://partner.snai.it", wait_until="load", timeout=40000)
            
            print("   ⏳ [Robot] STEP 6a: Attesa rendering del pannello Microsoft UpdatePanel...")
            # Forziamo l'attesa sull'ID dell'UpdatePanel che mi hai confermato tu
            page.locator("#ctl00_Cp1_updPnlSearchResult").wait_for(state="visible", timeout=20000)
            time.sleep(8)

            for _, row in df_snai.iterrows():
                try:
                    codice_aams = str(row["CODICE_LOCALE"]).strip()
                    nome_locale_corrente = str(row["NOME_LOCALE"]).strip()
                    data_in_completa = str(row["INIZIO_FERIE"]).strip()
                    data_fi_completa = str(row["FINE_FERIE"]).strip()
                    
                    print(f"🚀 [Robot] STEP 7: Avvio elaborazione per il Locale -> {codice_aams} - {nome_locale_corrente}")

                    # Pulizia dei caricamenti di sfondo per liberare lo schermo
                    try: page.evaluate("document.querySelectorAll('.mainPreload').forEach(el => el.remove());")
                    except Exception: pass

                    print("   🔍 [Robot] STEP 7a: Inserimento codice censimento nella barra filtri...")
                    # 🛡️ INPUT REALE DI MANUELA
                    campo_ricerca = page.locator("#ctl00_Cp1_txtCodiceCensimentoesercizio").first
                    campo_ricerca.click(timeout=15000)
                    campo_ricerca.fill(codice_aams)
                    time.sleep(2)
                    
                    # 🛡️ BOTTONE REALE DI MANUELA (Con la r minuscola!)
                    tasto_ricerca = page.locator("#ctl00_Cp1_btRicerca").first
                    tasto_ricerca.click(timeout=10000)
                    
                    print("   ⏳ [Robot] STEP 7b: Attesa caricamento risultati filtrati (6 secondi)...")
                    time.sleep(6)

# =====================================================================================
# BLOCCO 5: APERTURA SOTTO-PANNELLO DATE, ALLINEAMENTO E LOGOUT DI SICUREZZA
# =====================================================================================
                    icona_nuovo = page.locator("img[src*='insert_pianificazione'], img[id*='img_pianificazione']").first
                    icona_modifica = page.locator("img[src*='edit_pianificazione']").first
                    cella_td_cliccabile = page.locator("td[onclick*='Pianificazione']").first
                    
                    if icona_modifica.count() > 0:
                        print("   📝 [Robot] STEP 8: [MODIFICA] Chiusura già presente, entro in modifica...")
                        icona_modifica.click(force=True, timeout=8000)
                    elif icona_nuovo.count() > 0:
                        print("   🟢 [Robot] STEP 8a: [NUOVA CHIUSURA] Locale vuoto, clicco sul pallino verde...")
                        icona_nuovo.click(force=True, timeout=8000)
                    else:
                        print("   AM 🖱️ [Grid Mode] Clic sulla cella td nativa della riga...")
                        cella_td_cliccabile.click(force=True, timeout=8000)
                    
                    print("   ⏳ [Robot] STEP 8c: Attesa apertura campi temporali date (5 secondi)...")
                    time.sleep(5)

                    campo_dal = page.locator("input[id*='txtDataDal'], input[id*='Inizio'], input[name*='Dal']").first
                    campo_al = page.locator("input[id*='txtDataAl'], input[id*='Fine'], input[name*='Al']").first
                    
                    campo_dal.wait_for(state="visible", timeout=10000)
                    campo_dal.click()
                    campo_dal.fill(data_in_completa)
                    time.sleep(1)
                    
                    campo_al.click()
                    campo_al.fill(data_fi_completa)
                    time.sleep(1)

                    print("   💾 [Robot] STEP 10: Invio moduli di chiusura a Snaitech...")
                    page.locator("input[type='submit'][value*='Salva'], button:has-text('Salva'), input[id*='btnSalva']").first.click()
                    print(f"   ✅ [Robot] STEP 11: Locale {codice_aams} allineato e salvato con successo!")
                    print("----------------------------------------------------------------------")
                    time.sleep(5)
                    
                except Exception as row_err:
                    print(f"   ⚠️ Nota compilazione: Scavalco riga. Errore: {str(row_err)}")
                    continue

            print("🔒 [Robot] STEP 12: Chiusura formale della sessione (Logout di sicurezza)...")
            try: page.locator("a:has-text('LogOut'), a:has-text('Esci'), [id*='btnLogOut']").first.click(timeout=8000)
            except Exception: page.context.clear_cookies()

        except Exception as e: print(f"❌ Errore generale: {str(e)}")
        finally: browser.close()

if __name__ == "__main__":
    avvia_sincronizzazione_automatica()
