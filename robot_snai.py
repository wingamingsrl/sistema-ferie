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
    try:
        nome_file_locale = "storico_ferie.xlsx"
        if os.path.exists(nome_file_locale):
            return pd.read_excel(nome_file_locale).fillna("")
    except Exception: pass
    return pd.DataFrame()

def genera_codice_otp_automatico():
    totp = pyotp.TOTP(CHIAVE_SEGRETA_2FA.strip().upper().replace(" ", ""))
    return totp.now()

def avvia_sincronizzazione_automatica():
    df_ferie = preleva_storico_diretto_da_cloud()
    if df_ferie.empty: return

    df_snai = df_ferie[
        df_ferie["CONCESSIONARIO"].astype(str).str.lower().str.contains("snai|snaitech", regex=True) |
        df_ferie["NOME_LOCALE"].astype(str).str.lower().str.contains("snai", regex=True)
    ]
    if df_snai.empty: return

    print(f"🤖 Rilevati {len(df_snai)} locali Snaitech. Avvio allineamento finale su codici HTML...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]) 
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            locale="it-IT",
            timezone_id="Europe/Rome"
        )
        page = context.new_page()

        try:
            page.goto("https://partner.snai.it", wait_until="networkidle", timeout=60000)
            time.sleep(5)
            
            page.locator("input[id*='username'], input[name='username'], input[type='text']").first.fill(SNAI_USER)
            page.locator("input#password, input[name='password'], input[type='password']").first.fill(SNAI_PASS)
            page.locator("button[type='submit'], input[type='submit'], .btn-login").first.click()
            time.sleep(11)
            
            try: page.evaluate("document.querySelectorAll('.modal, .modal-backdrop, .fade.in').forEach(el => el.remove());")
            except Exception: pass

            codice_totp = genera_codice_otp_automatico()
            input_token = page.locator("input[id*='otp']:not([type='hidden']), input[name*='otp']:not([type='hidden']), input[type='text']").first
            input_token.fill(str(codice_totp))
            time.sleep(1)
            page.keyboard.press("Enter")
            time.sleep(15)
            
            print("📬 Spostamento sulla pagina degli Esercizi censiti...")
            page.goto("https://snai.it/secure/Anagrafiche/Esercizi.aspx", wait_until="networkidle", timeout=50000)
            time.sleep(10)

            for _, row in df_snai.iterrows():
                try:
                    codice_aams = str(row["CODICE_LOCALE"]).strip()
                    data_in_completa = str(row["INIZIO_FERIE"]).strip()
                    data_fi_completa = str(row["FINE_FERIE"]).strip()
                    
                    print(f"🚀 Ispezione visiva Locale Snaitech -> {codice_aams}")

                    campo_ricerca = page.locator("input[id*='Censimento'], input[name*='Censimento'], input[id*='txtCodiceCensimento'], input[id*='txtCodice']").first
                    campo_ricerca.click(timeout=10000)
                    campo_ricerca.fill(codice_aams)
                    time.sleep(1)
                    
                    tasto_ricerca = page.locator("input[type='submit'][value='Ricerca'], input[value='Ricerca'], button:has-text('Ricerca'), input[id*='Ricerca']").first
                    tasto_ricerca.click(timeout=10000)
                    
                    print("   ⏳ Attesa caricamento griglia dei risultati (6 secondi)...")
                    time.sleep(6)

                    # 🛡️ MAPPA SENSORIALE HTML CERTIFICATA DA MANUELA:
                    icona_nuovo_inserimento = page.locator("img[src*='insert_pianificazione'], img[id*='img_pianificazione']").first
                    icona_modifica_esistente = page.locator("img[src*='edit_pianificazione']").first
                    
                    if icona_modifica_esistente.count() > 0:
                        print("   📝 [EDIT_PIANIFICAZIONE DETECTED] Trovata chiusura esistente, entro in modifica...")
                        icona_modifica_esistente.click(timeout=10000)
                    elif icona_nuovo_inserimento.count() > 0:
                        print("   🟢 [INSERT_PIANIFICAZIONE DETECTED] Nuovo locale vuoto, inserisco da zero...")
                        icona_nuovo_inserimento.click(timeout=10000)
                    else:
                        print("   ⚠️ Icona specifica non isolata, eseguo il clic sulla cella td della riga...")
                        page.locator("td[onclick*='Pianificazione']").first.click(timeout=10000)
                    time.sleep(6)

                    # Compilazione dei campi dal sotto-pannello sbloccato
                    campo_dal = page.locator("input[id*='txtDataDal'], input[id*='Inizio'], input[name*='Dal']").first
                    campo_al = page.locator("input[id*='txtDataAl'], input[id*='Fine'], input[name*='Al']").first
                    
                    # Controlla se il valore inserito coincide già per non fare salvataggi a vuoto
                    valore_attuale_dal = campo_dal.input_value() if campo_dal.count() > 0 else ""
                    if valore_attuale_dal == data_in_completa:
                        print(f"   ℹ️ Le date inserite ({data_in_completa}) coincidono già sul portale Snaitech. Salto il locale.")
                        page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", timeout=30000)
                        time.sleep(5)
                        continue

                    campo_dal.fill(data_in_completa)
                    time.sleep(1)
                    campo_al.fill(data_fi_completa)
                    time.sleep(1)

                    # Pressione del tasto Salva reale del portale Snaitech
                    tasto_salva = page.locator("input[type='submit'][value*='Salva'], button:has-text('Salva'), input[id*='btnSalva']").first
                    tasto_salva.click(timeout=10000)
                    print(f"   ✅ Locale {codice_aams} allineato e salvato con successo nel database Snaitech!")
                    time.sleep(5)
                    
                    page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", timeout=30000)
                    time.sleep(5)
                except Exception as e_row:
                    print(f"   ⚠️ Nota compilazione riga: Sposto focus per riga successiva.")
                    page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", timeout=30000)
                    time.sleep(4)
                    continue
        except Exception as e: print(f"❌ Errore generale: {str(e)}")
        finally: browser.close()

if __name__ == "__main__":
    avvia_sincronizzazione_automatica()
