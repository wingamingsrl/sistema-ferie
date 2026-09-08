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

CHIAVE_ACCESSO_GIT = os.environ.get("TOKEN_GITHUB_ACTIONS", "")

def preleva_storico_diretto_da_cloud():
    print("📡 [Robot] Lettura del database Excel locale sul server Actions...")
    try:
        nome_file_locale = "storico_ferie.xlsx"
        if os.path.exists(nome_file_locale):
            return pd.read_excel(nome_file_locale).fillna("")
        else:
            print(f"❌ File {nome_file_locale} non trovato sul server.")
    except Exception as e:
        print(f"⚠️ Errore lettura file: {str(e)}")
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

    print(f"🤖 Rilevati {len(df_snai)} locali Snaitech. Avvio Chrome con SCHERMATURA ANTI-BOT...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]) 
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            locale="it-IT",
            timezone_id="Europe/Rome"
        )
        page = context.new_page()

        try:
            print("🌐 Connessione schermata a partner.snai.it...")
            page.goto("https://partner.snai.it", wait_until="networkidle", timeout=60000)
            time.sleep(6)
            
            try: page.mouse.click(100, 100)
            except Exception: pass
            time.sleep(2)
            
            print("📝 Inserimento credenziali sul portale...")
            input_user = page.locator("input[id*='username'], input[name='username'], input[type='text']").first
            input_user.click(timeout=15000)
            input_user.fill(SNAI_USER)
            time.sleep(1)
            
            input_pass = page.locator("input[id*='password'], input[name='password'], input[type='password']").first
            input_pass.click(timeout=15000)
            input_pass.fill(SNAI_PASS)
            time.sleep(1)
            
            print("🚀 Invio moduli di accesso...")
            page.locator("button[type='submit'], input[type='submit'], .btn-login, .button").first.click()
            time.sleep(5)
            
            print("⏳ Attesa del countdown di sicurezza Snaitech (11 secondi)...")
            time.sleep(11)
            try: page.evaluate("document.querySelectorAll('.modal, .modal-backdrop, .fade.in').forEach(el => el.remove());")
            except Exception: pass
            time.sleep(2)

            codice_totp = genera_codice_otp_automatico()
            print(f"🔑 Codice OTP calcolato -> {codice_totp}")
            
            input_token = page.locator("input[id*='otp']:not([type='hidden']), input[name*='otp']:not([type='hidden']), input[id*='code']:not([type='hidden']), input[type='text']:not([type='hidden'])").first
            input_token.click(timeout=15000)
            input_token.fill(str(codice_totp))
            time.sleep(2)
            
            print("⌨️ Pressione del tasto Enter da tastiera...")
            page.keyboard.press("Enter")
            
            print("⏳ Caricamento area riservata partner.snai.it (15 secondi)...")
            time.sleep(15)
            
            print("📬 Spostamento sulla pagina degli Esercizi censiti...")
            page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", wait_until="networkidle", timeout=50000)
            time.sleep(10)

            for _, row in df_snai.iterrows():
                try:
                    codice_aams = str(row["CODICE_LOCALE"]).strip()
                    data_in_completa = str(row["INIZIO_FERIE"]).strip()
                    data_fi_completa = str(row["FINE_FERIE"]).strip()
                    
                    print(f"🚀 Ispezione visiva Locale Snaitech -> {codice_aams}")

                    target_frame = page
                    for f in page.frames:
                        if "Esercizi" in f.url or f.locator("input").count() > 0:
                            target_frame = f
                            break

                    # 🛡️ PUNTATORE RETTIFICATO DA MANUELA: Cerca l'input abbinato al testo visibile 'Codice Censimento'
                    campo_ricerca = target_frame.locator("input[id*='Censimento'], input[name*='Censimento'], input[id*='txtCodiceCensimento'], input[placeholder*='Censimento'], input[type='text']").first
                    campo_ricerca.click(timeout=15000)
                    campo_ricerca.fill(codice_aams)
                    time.sleep(1)
                    
                    # Forza la ricerca cliccando sul tasto reale Cerca/Filtra del pannello Microsoft
                    tasto_cerca = target_frame.locator("input[type='submit'][value*='Cerca'], input[id*='Cerca'], button[id*='Cerca'], input[value*='Filtra']").first
                    if tasto_cerca.count() > 0:
                        tasto_cerca.click()
                    else:
                        target_frame.keyboard.press("Enter")
                    
                    print("   ⏳ Attesa caricamento riga esercizio (5 secondi)...")
                    time.sleep(5)

                    icona_agenda_matita = "img[id*='img_modifica'], img[id*='img_dettaglio'], img[src*='agenda'], img[src*='edit'], [title*='Modifica'], img[id*='Pianificazione']"
                    pallino_verde_nuovo = "img[id*='img_pianificazione'], img[src*='insert_pianificazione'], img[src*='plus']"
                    
                    if target_frame.locator(icona_agenda_matita).count() > 0:
                        print("   📝 [AGENDA/MATITA DETECTED] Chiusura già presente. Clic per entrare in modifica...")
                        target_frame.locator(icona_agenda_matita).first.click(timeout=10000)
                    elif target_frame.locator(pallino_verde_nuovo).count() > 0:
                        print("   🟢 [PALLINO VERDE DETECTED] Nuovo locale vuoto. Clic per inserire da zero...")
                        target_frame.locator(pallino_verde_nuovo).first.click(timeout=10000)
                    else:
                        print("   ⚠️ Icona specifica non vista, tento il clic sulla prima immagine utile della riga...")
                        target_frame.locator("td img, tr img, table img").first.click(timeout=10000)
                    time.sleep(6)

                    campo_dal = "input[id*='txtDataDal'], input[id*='Inizio'], input[name*='Dal']"
                    campo_al = "input[id*='txtDataAl'], input[id*='Fine'], input[name*='Al']"
                    
                    valore_attuale_dal = target_frame.locator(campo_dal).first.input_value() if target_frame.locator(campo_dal).count() > 0 else ""
                    if valore_attuale_dal == data_in_completa:
                        print(f"   ℹ️ Le date inserite coincidono già ({data_in_completa}). Salto il salvataggio per sicurezza.")
                        page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", timeout=30000)
                        time.sleep(5)
                        continue

                    target_frame.locator(campo_dal).first.fill(data_in_completa)
                    time.sleep(1)
                    target_frame.locator(campo_al).first.fill(data_fi_completa)
                    time.sleep(1)

                    tasto_salva = "input[type='submit'][value*='Salva'], button:has-text('Salva'), input[id*='btnSalva']"
                    target_frame.locator(tasto_salva).first.click(timeout=10000)
                    print(f"   ✅ Allineato e salvato correttamente nel database Snaitech!")
                    time.sleep(5)
                    
                    page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", timeout=30000)
                    time.sleep(5)
                    
                except Exception as e_row:
                    print(f"⚠️ Errore riga: {str(e_row)}")
                    page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", timeout=30000)
                    time.sleep(5)
                    continue
        except Exception as e:
            print(f"❌ Errore generale di navigazione: {str(e)}")
        finally:
            browser.close()

if __name__ == "__main__":
    avvia_sincronizzazione_automatica()
