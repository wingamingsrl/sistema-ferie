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
    print("📡 [Robot] Lettura del database Excel locale sul server Actions...")
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

    print(f"🤖 Rilevati {len(df_snai)} locali Snaitech. Avvio navigazione inter-frame...")

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
                    
                    print(f"🚀 Ispezione visiva Locale Snaitech -> {codice_aams}")

                    # 🛡️ SCANSIONE REALE DI TUTTI I SOTTO-SCHERMI (FRAME) DISPONIBILI NELLA PAGINA
                    for target_frame in page.frames:
                        try:
                            campo_ricerca = target_frame.locator("input[id*='Censimento'], input[name*='Censimento'], input[id*='txtCodiceCensimento']").first
                            if campo_ricerca.count() > 0:
                                campo_ricerca.click(timeout=3000)
                                campo_ricerca.fill(codice_aams)
                                
                                tasto_cerca = target_frame.locator("input[type='submit'][value*='Cerca'], input[id*='Cerca'], input[value*='Filtra']").first
                                if tasto_cerca.count() > 0: tasto_cerca.click()
                                else: page.keyboard.press("Enter")
                                time.sleep(6)

                                # Selettori espansi comprensivi di tutte le icone grafiche di Snaitech
                                icona_agenda_matita = target_frame.locator("img[id*='img_modifica'], img[id*='img_dettaglio'], img[src*='agenda'], img[src*='edit'], [title*='Modifica']")
                                pallino_verde_nuovo = target_frame.locator("img[id*='img_pianificazione'], img[src*='insert_pianificazione'], img[src*='plus']")
                                
                                if icona_agenda_matita.count() > 0:
                                    print("   📝 [MODIFICA] Clic sull'agenda/matita...")
                                    icona_agenda_matita.first.click()
                                elif pallino_verde_nuovo.count() > 0:
                                    print("   🟢 [NUOVO] Clic sul pallino verde...")
                                    pallino_verde_nuovo.first.click()
                                else:
                                    # Se non vede le icone specifiche, forza il clic sul primo link/immagine utile della riga di risultato
                                    target_frame.locator("td a img, tr td a, .Grid img").first.click()
                                time.sleep(5)

                                campo_dal = target_frame.locator("input[id*='txtDataDal'], input[id*='Inizio']").first
                                campo_al = target_frame.locator("input[id*='txtDataAl'], input[id*='Fine']").first
                                
                                campo_dal.fill(data_in_completa)
                                time.sleep(1)
                                campo_al.fill(data_fi_completa)
                                time.sleep(1)

                                target_frame.locator("input[type='submit'][value*='Salva'], button:has-text('Salva'), input[id*='btnSalva']").first.click()
                                print(f"   ✅ Allineato e salvato con successo nel pannello Snaitech!")
                                time.sleep(4)
                                break
                        except Exception: continue
                    
                    page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", timeout=30000)
                    time.sleep(5)
                except Exception as e_row:
                    print(f"⚠️ Errore riga: {str(e_row)}")
                    page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", timeout=30000)
                    time.sleep(4)
        except Exception as e: print(f"❌ Errore generale: {str(e)}")
        finally: browser.close()

if __name__ == "__main__":
    avvia_sincronizzazione_automatica()
