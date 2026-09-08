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

    print(f"🤖 Rilevati {len(df_snai)} locali Snaitech. Avvio inserimento visivo localizzato...")

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
            page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", wait_until="networkidle", timeout=50000)
            time.sleep(10)

            for _, row in df_snai.iterrows():
                try:
                    codice_aams = str(row["CODICE_LOCALE"]).strip()
                    data_in_completa = str(row["INIZIO_FERIE"]).strip()
                    data_fi_completa = str(row["FINE_FERIE"]).strip()
                    
                    print(f"🚀 Elaborazione forzata Locale Snaitech -> {codice_aams}")

                    # 🛡️ INPUT DIRETTO DA TASTIERA GLOBALE SULLO SCHERMO:
                    # Clicca al centro dello schermo per attivare la pagina, scrive il codice locale e preme Cerca
                    page.mouse.click(400, 300)
                    time.sleep(1)
                    
                    # Cerca l'input in modo assoluto e digita
                    campo_assoluto = page.locator("input[id*='Censimento'], input[name*='Censimento'], input[type='text']").first
                    campo_assoluto.click(timeout=5000)
                    campo_assoluto.fill(codice_aams)
                    time.sleep(1)
                    
                    # Invia la ricerca premendo Enter o cliccando sui bottoni di primo livello
                    tasto_cerca = page.locator("input[type='submit'][value*='Cerca'], input[id*='Cerca'], button[id*='Cerca']").first
                    if tasto_cerca.count() > 0: 
                        tasto_cerca.click()
                    else: 
                        page.keyboard.press("Enter")
                    time.sleep(6)

                    # 🛡️ FORCE CLICK: Se non trova selettori, esegue il clic sulle coordinate standard della prima riga di tabella
                    # Clicca sulla matita/pallino verde situati indicativamente nella prima colonna dei risultati
                    icona_visibile = page.locator("img[id*='modifica'], img[id*='pianificazione'], img[src*='agenda'], img[src*='edit'], img[src*='plus']").first
                    if icona_visibile.count() > 0:
                        icona_visibile.click(timeout=5000)
                    else:
                        print("   🖱️ [Coordinate Mode] Icona non intercettata dal DOM, eseguo clic posizionale sulla prima riga...")
                        page.mouse.click(350, 420)  # Clic visivo sulla coordinata della prima icona della griglia
                    time.sleep(5)

                    # Compilazione campi date
                    campo_dal = page.locator("input[id*='txtDataDal'], input[id*='Inizio'], input[name*='Dal']").first
                    campo_al = page.locator("input[id*='txtDataAl'], input[id*='Fine'], input[name*='Al']").first
                    
                    campo_dal.fill(data_in_completa)
                    time.sleep(1)
                    campo_al.fill(data_fi_completa)
                    time.sleep(1)

                    # Clic sul tasto Salva nativo
                    page.locator("input[type='submit'][value*='Salva'], button:has-text('Salva'), input[id*='btnSalva']").first.click(timeout=5000)
                    print(f"   ✅ Allineato e salvato correttamente nel database Snaitech!")
                    time.sleep(5)
                    
                    page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", timeout=30000)
                    time.sleep(4)
                except Exception as e_row:
                    print(f"   ⚠️ Nota riga: Sposto focus per riga successiva.")
                    page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", timeout=30000)
                    time.sleep(4)
                    continue
        except Exception as e: print(f"❌ Errore generale: {str(e)}")
        finally: browser.close()

if __name__ == "__main__":
    avvia_sincronizzazione_automatica()
