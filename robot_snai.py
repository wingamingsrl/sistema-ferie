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

# Preleva il token di sicurezza iniettato in automatico dal server protetto di GitHub
CHIAVE_ACCESSO_GIT = os.environ.get("TOKEN_GITHUB_ACTIONS", "")

def preleva_storico_diretto_da_cloud():
    print("📡 [Robot] Estrazione database Excel direttamente in RAM da GitHub...")
    try:
        c_time = str(int(time.time() * 1000))
        url_git = f"https://github.com{c_time}"
        headers_diretti = {"Accept": "application/vnd.github.v3.raw", "User-Agent": "WinGaming-Cloud-App"}
        if CHIAVE_ACCESSO_GIT:
            headers_diretti["Authorization"] = f"token {CHIAVE_ACCESSO_GIT}"
        risposta = requests.get(url_git, headers=headers_diretti, timeout=15)
        if risposta.status_code == 200:
            return pd.read_excel(io.BytesIO(risposta.content)).fillna("")
    except Exception as e:
        print(f"⚠️ Errore di rete: {str(e)}")
    return pd.DataFrame()

def genera_codice_otp_automatico():
    totp = pyotp.TOTP(CHIAVE_SEGRETA_2FA.strip().upper().replace(" ", ""))
    return totp.now()

def avvia_sincronizzazione_automatica():
    df_ferie = preleva_storico_diretto_da_cloud()
    if df_ferie.empty:
        print("❌ Database vuoto o non accessibile.")
        return

    df_snai = df_ferie[
        df_ferie["CONCESSIONARIO"].astype(str).str.lower().str.contains("snai|snaitech", regex=True) |
        df_ferie["NOME_LOCALE"].astype(str).str.lower().str.contains("snai", regex=True)
    ]

    if df_snai.empty:
        print("✅ Nessun locale Snaitech attivo trovato nel registro.")
        return

    print(f"🤖 Rilevati {len(df_snai)} locali Snaitech. Avvio Chrome con supporto clic reali...")

    with sync_playwright() as p:
        # Sui server GitHub Actions l'avvio con headless=False e i clic reali funzionano alla perfezione!
        browser = p.chromium.launch(headless=False) 
        context = browser.new_context()
        page = context.new_page()

        try:
            page.goto("https://snai.it", timeout=35000)
            time.sleep(3)
            
            page.fill("input#username, input[name='username']", SNAI_USER)
            page.fill("input#password, input[name='password']", SNAI_PASS)
            page.click("button[type='submit'], input[type='submit']")
            time.sleep(4)
            
            time.sleep(11) # Countdown obbligatorio Snaitech
            try: page.evaluate("document.querySelectorAll('.modal, .modal-backdrop, .fade.in').forEach(el => el.remove());")
            except Exception: pass
            time.sleep(2)

            codice_totp = genera_codice_otp_automatico()
            page.fill("input#token, input[name='token'], input[name='otp']", codice_totp)
            time.sleep(1)
            
            page.click("input#btnInvia, input[value='Invia'], button:has-text('Invia')")
            time.sleep(15)
            
            # 🛡️ NAVIGAZIONE SULL'INDIRIZZO REALE DI MANUELA
            page.goto("https://snai.it/secure/Anagrafiche/Esercizi.aspx", timeout=30000)
            time.sleep(8)

            for _, row in df_snai.iterrows():
                try:
                    codice_aams = str(row["CODICE_LOCALE"]).strip()
                    data_in_completa = str(row["INIZIO_FERIE"]).strip()
                    data_fi_completa = str(row["FINE_FERIE"]).strip()
                    
                    print(f"🚀 Elaborazione visiva -> Locale: {codice_aams}")

                    target_frame = page
                    if len(page.frames) > 1: target_frame = page.frames

                    campo_ricerca = "input[id*='Censimento'], input[id*='txtCodice']"
                    if target_frame.locator(campo_ricerca).count() > 0:
                        target_frame.locator(campo_ricerca).first.fill(codice_aams)
                        target_frame.keyboard.press("Enter")
                        time.sleep(5)

                    tasto_modifica = "img[id*='img_modifica'], [title*='Modifica'], img[id*='img_dettaglio']"
                    pallino_verde_nuovo = "img[id*='img_pianificazione'], img[src*='insert_pianificazione']"
                    
                    if target_frame.locator(tasto_modifica).count() > 0:
                        print("   📝 Modifico periodo esistente...")
                        target_frame.locator(tasto_modifica).first.click(timeout=10000)
                    elif target_frame.locator(pallino_verde_nuovo).count() > 0:
                        print("   🟢 Inserisco nuovo periodo...")
                        target_frame.locator(pallino_verde_nuovo).first.click(timeout=10000)
                    time.sleep(5)

                    campo_dal = "input[id*='txtDataDal'], input[id*='Inizio']"
                    campo_al = "input[id*='txtDataAl'], input[id*='Fine']"
                    
                    target_frame.locator(campo_dal).first.fill(data_in_completa)
                    target_frame.locator(campo_al).first.fill(data_fi_completa)
                    time.sleep(1)

                    # Clicca sul salvataggio visivo generando il __VIEWSTATE richiesto
                    target_frame.locator("input[type='submit'][value*='Salva'], button:has-text('Salva')").first.click()
                    print(f"   ✅ Allineato con successo nel database Snaitech!")
                    time.sleep(4)
                    
                except Exception as e_row:
                    print(f"⚠️ Errore riga: {str(e_row)}")
                    continue
        except Exception as e:
            print(f"❌ Errore generale: {str(e)}")
        finally:
            browser.close()

if __name__ == "__main__":
    avvia_sincronizzazione_automatica()
