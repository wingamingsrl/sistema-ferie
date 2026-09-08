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
        # 🛡️ FIX DEFINITIVO: Legge il file dall'hard disk virtuale locale senza chiamate di rete (Evita il 406)
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

    print(f"🤖 Rilevati {len(df_snai)} locali Snaitech. Avvio Chrome con gestione Agenda/Matita...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]) 
        context = browser.new_context()
        page = context.new_page()

        try:
            print("🌐 Connessione a partner.snai.it...")
            page.goto("https://snai.it", timeout=45000)
            time.sleep(4)
            
            page.fill("input#username, input[name='username']", SNAI_USER)
            page.fill("input#password, input[name='password']", SNAI_PASS)
            page.click("button[type='submit'], input[type='submit'], .btn-login")
            time.sleep(5)
            
            print("⏳ Attesa del countdown di sicurezza Snaitech (11 secondi)...")
            time.sleep(11)
            try: page.evaluate("document.querySelectorAll('.modal, .modal-backdrop, .fade.in').forEach(el => el.remove());")
            except Exception: pass
            time.sleep(2)

            codice_totp = genera_codice_otp_automatico()
            print(f"🔑 Codice OTP calcolato -> {codice_totp}")
            page.fill("input#token, input[name='token'], input[name='otp']", codice_totp)
            time.sleep(1)
            
            page.click("input#btnInvia, input[value='Invia'], button:has-text('Invia')")
            print("⏳ Caricamento area riservata (15 secondi)...")
            time.sleep(15)
            
            # 🛡️ NAVIGAZIONE SULLA PAGINA DEGLI ESERCIZI DI MANUELA
            print("📬 Spostamento sulla pagina degli Esercizi censiti...")
            page.goto("https://snai.it/secure/Anagrafiche/Esercizi.aspx", timeout=40000)
            time.sleep(10)

            for _, row in df_snai.iterrows():
                try:
                    codice_aams = str(row["CODICE_LOCALE"]).strip()
                    data_in_completa = str(row["INIZIO_FERIE"]).strip()
                    data_fi_completa = str(row["FINE_FERIE"]).strip()
                    
                    print(f"🚀 Ispezione visiva Locale Snaitech -> {codice_aams}")

                    target_frame = page
                    if len(page.frames) > 1: target_frame = page.frames

                    campo_ricerca = "input[id*='Censimento'], input[id*='txtCodice'], input[name*='Codice']"
                    if target_frame.locator(campo_ricerca).count() > 0:
                        target_frame.locator(campo_ricerca).first.fill(codice_aams)
                        target_frame.keyboard.press("Enter")
                        time.sleep(6)

                    # 🛡️ LOGICA SBLOCCATA DI MANUELA (RICERCA ICONE AGENDA / MATITA / PALLINO)
                    icona_agenda_matita = "img[id*='img_modifica'], img[id*='img_dettaglio'], img[src*='agenda'], img[src*='edit'], [title*='Modifica']"
                    pallino_verde_nuovo = "img[id*='img_pianificazione'], img[src*='insert_pianificazione'], img[src*='plus']"
                    
                    if target_frame.locator(icona_agenda_matita).count() > 0:
                        print("   📝 [AGENDA/MATITA DETECTED] Chiusura già presente. Clic per entrare in modifica...")
                        target_frame.locator(icona_agenda_matita).first.click(timeout=10000)
                    elif target_frame.locator(pallino_verde_nuovo).count() > 0:
                        print("   🟢 [PALLINO VERDE DETECTED] Nuovo locale vuoto. Clic per inserire da zero...")
                        target_frame.locator(pallino_verde_nuovo).first.click(timeout=10000)
                    else:
                        print("   ⚠️ Icona non identificata, tento il clic d'emergenza sulla prima riga...")
                        target_frame.locator("td img").first.click(timeout=10000)
                    time.sleep(6)

                    # Compilazione dei campi della data interni al sotto-pannello sbloccato
                    campo_dal = "input[id*='txtDataDal'], input[id*='Inizio'], input[name*='Dal']"
                    campo_al = "input[id*='txtDataAl'], input[id*='Fine'], input[name*='Al']"
                    
                    # Legge se i campi hanno già lo stesso valore per non sovrascrivere a vuoto
                    valore_attuale_dal = target_frame.locator(campo_dal).first.input_value() if target_frame.locator(campo_dal).count() > 0 else ""
                    if valore_attuale_dal == data_in_completa:
                        print(f"   ℹ️ Le date inserite coincido già ({data_in_completa}). Salto il salvataggio per sicurezza.")
                        # Clicca sul tasto Annulla o torna indietro se presente, altrimenti ricarica la pagina principale
                        page.goto("https://snai.it/secure/Anagrafiche/Esercizi.aspx", timeout=30000)
                        time.sleep(5)
                        continue

                    # Se sono diverse o vuote, digita ed allinea
                    target_frame.locator(campo_dal).first.fill(data_in_completa)
                    time.sleep(1)
                    target_frame.locator(campo_al).first.fill(data_fi_completa)
                    time.sleep(1)

                    # Clicca sul salvataggio visivo ufficiale di Snaitech sbloccando la tabella
                    tasto_salva = "input[type='submit'][value*='Salva'], button:has-text('Salva'), input[id*='btnSalva']"
                    target_frame.locator(tasto_salva).first.click(timeout=10000)
                    print(f"   ✅ Allineato e salvato correttamente nel database Snaitech!")
                    time.sleep(5)
                    
                    # Torna alla pagina di ricerca per il locale successivo
                    page.goto("https://snai.it/secure/Anagrafiche/Esercizi.aspx", timeout=30000)
                    time.sleep(5)
                    
                except Exception as e_row:
                    print(f"⚠️ Errore riga: {str(e_row)}")
                    page.goto("https://snai.it/secure/Anagrafiche/Esercizi.aspx", timeout=30000)
                    time.sleep(5)
                    continue
        except Exception as e:
            print(f"❌ Errore generale di navigazione: {str(e)}")
        finally:
            browser.close()

if __name__ == "__main__":
    avvia_sincronizzazione_automatica()

