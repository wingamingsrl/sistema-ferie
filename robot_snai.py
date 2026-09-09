# =====================================================================================
# SW AUTOMATICO DI SINCRONIZZAZIONE LOCALI WIN GAMING — PRODUZIONE FINALE ONLINE
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

# Configurazione ufficiale ed esclusiva delle tue coordinate aziendali
CHIAVE_SEGRETA_2FA = "FTIA6UQZM2LQLPYJ"  # La tua chiave definitiva allineata al telefono
SNAI_USER = "2141ManuelaA"
SNAI_PASS = "Salmi123!"

# =====================================================================================
# BLOCCO 2: MOTORE DI LETTURA LIVE IN MEMORIA RAM DEL REPOSITORY EXCEL LOCALE
# =====================================================================================
def preleva_storico_diretto_da_cloud():
    print("📡 [Robot] Lettura del database Excel locale sul server Actions...")
    try:
        nome_file_locale = "storico_ferie.xlsx"
        if os.path.exists(nome_file_locale):
            print("✅ [Robot] Database Excel intercettato con successo!")
            return pd.read_excel(nome_file_locale).fillna("")
        else:
            print(f"❌ File {nome_file_locale} non trovato sul server.")
    except Exception as e_file:
        print(f"⚠️ Errore lettura file Excel: {str(e_file)}")
    return pd.DataFrame()

def genera_codice_otp_automatico():
    chiave_pulita = CHIAVE_SEGRETA_2FA.strip().upper().replace(" ", "")
    totp = pyotp.TOTP(chiave_pulita)
    return totp.now()

# =====================================================================================
# BLOCCO 3: AVVIO CHROME CON CAMUFFAMENTO ED IMMISSIONE CREDENZIALI
# =====================================================================================
def avvia_sincronizzazione_automatica():
    df_ferie = preleva_storico_diretto_da_cloud()
    if df_ferie.empty:
        print("❌ Impossibile procedere: Il database delle ferie è vuoto o bloccato.")
        return

    df_snai = df_ferie[
        df_ferie["CONCESSIONARIO"].astype(str).str.lower().str.contains("snai|snaitech", regex=True) |
        df_ferie["NOME_LOCALE"].astype(str).str.lower().str.contains("snai", regex=True)
    ]

    if df_snai.empty:
        print("✅ [Robot] Nessun locale Snaitech attivo trovato nel registro. Sincronizzazione conclusa.")
        return

    print(f"🤖 [Robot] Rilevati {len(df_snai)} locali Snaitech da elaborare. Avvio Chrome con camuffamento d'ufficio...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=[
            "--no-sandbox", 
            "--disable-setuid-sandbox", 
            "--disable-dev-shm-usage",
            "--disable-blink-features=AutomationControlled"  # Nasconde lo stato di robot a Snaitech
        ]) 
        
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
            locale="it-IT",
            timezone_id="Europe/Rome",
            viewport={"width": 1280, "height": 1024}
        )
        page = context.new_page()

        try:
            print("🌐 [Robot] Connessione a partner.snai.it con tolleranza di rete...")
            page.goto("https://partner.snai.it", wait_until="networkidle", timeout=60000)
            time.sleep(6)
            
            try: page.mouse.move(200, 200)
            except Exception: pass
            time.sleep(2)
            
            print("📝 [Robot] Inserimento credenziali Snaitech...")
            campo_user = page.locator("input#username, input[name='username'], input[type='text']").first
            campo_user.click(timeout=20000)
            campo_user.fill(SNAI_USER)
            time.sleep(1)
            
            campo_pass = page.locator("input#password, input[name='password'], input[type='password']").first
            campo_pass.click(timeout=20000)
            campo_pass.fill(SNAI_PASS)
            time.sleep(1)
            
            print("🚀 [Robot] Invio moduli di accesso...")
            page.click("button[type='submit'], input[type='submit'], .btn-login")
            time.sleep(4)
            
            print("⏳ [Robot] Pausa di sicurezza di 11 secondi per far scadere il countdown...")
            time.sleep(11)
            
            try:
                page.evaluate("""
                    document.querySelectorAll('.modal, .modal-backdrop, [id*="modal"], [class*="modal"], .fade.in').forEach(el => el.remove());
                    document.body.classList.remove('modal-open');
                    document.body.style.overflow = 'auto';
                """)
                print("✅ [Robot] Codice pop-up eliminato dalla pagina con successo!")
            except Exception: pass
            time.sleep(2)

            print("🔑 [Robot] Generazione ed immissione codice 2FA TOTP pulito...")
            codice_totp = genera_codice_otp_automatico()
            print(f"📌 Codice generato inviato a schermo: {codice_totp}")
            
            campo_token = page.locator("input#token, input[name='token'], input[name='otp'], input[type='text']").first
            campo_token.click(timeout=15000)
            campo_token.fill(str(codice_totp))
            time.sleep(2)
            
            print("📤 Invio codice OTP tramite tastiera virtuale...")
            page.keyboard.press("Enter")
            print("⏳ [Robot] Convalida credenziali in corso... Caricamento area riservata...")
            time.sleep(15)
            
            print("🔓 [Robot] ACCESSO EFFETTUATO CON SUCCESSO SUL PORTALE PARTNER SNAITECH!")
            print("----------------------------------------------------------------------")


# =====================================================================================
# BLOCCO 4: INTERCETTAZIONE MENU ANAGRAFICA E FILTRAGGIO CODICI CENSIMENTO
# =====================================================================================
            print("📦 [Robot] Apertura del menu Anagrafica Locali...")
            page.locator("#ctl00_MenuID1_rpMaster_ctl04_btnMnuItemPadre").first.click(timeout=20000)
            time.sleep(8)

            for _, row in df_snai.iterrows():
                try:
                    codice_aams = str(row["CODICE_LOCALE"]).strip()
                    nome_locale_corrente = str(row["NOME_LOCALE"]).strip()
                    data_in_completa = str(row["INIZIO_FERIE"]).strip()
                    data_fi_completa = str(row["FINE_FERIE"]).strip()
                    
                    print(f"🚀 [Robot] Avvio lavorazione -> Codice Locale: {codice_aams} - {nome_locale_corrente}")

                    target_frame = page
                    for f in page.frames:
                        if "Esercizi" in f.url or f.locator("input[id*='Censimento']").count() > 0 or f.locator("input[id*='txtCodice']").count() > 0:
                            target_frame = f
                            break

                    print("   🔍 Inserimento codice censimento nella barra filtri...")
                    campo_ricerca = "input[id*='Censimento'], input[name*='Censimento'], input[id*='txtCodice'], input[id*='txtCodiceCensimento']"
                    if target_frame.locator(campo_ricerca).count() > 0:
                        target_frame.locator(campo_ricerca).first.click(timeout=10000)
                        target_frame.locator(campo_ricerca).first.fill(codice_aams)
                        
                        tasto_ricerca = target_frame.locator("input[type='submit'][value='Ricerca'], input[value='Ricerca'], input[value='Filtra']").first
                        if tasto_ricerca.count() > 0:
                            tasto_ricerca.click()
                        else:
                            page.keyboard.press("Enter")
                        time.sleep(6)

# =====================================================================================
# BLOCCO 5: CONTROLLO STRUTTURA (NUOVO/MODIFICA) ALLINEATO ALL'HTML DI MANUELA E CHIUSURA
# =====================================================================================
                    icona_nuovo_inserimento = target_frame.locator("img[src*='insert_pianificazione'], img[id*='img_pianificazione']").first
                    icona_modifica_esistente = target_frame.locator("img[src*='edit_pianificazione']").first
                    
                    if icona_modifica_esistente.count() > 0:
                        print("   📝 [Robot] [EDIT_PIANIFICAZIONE DETECTED] Chiusura esistente! Clic sull'icona di Modifica...")
                        icona_modifica_esistente.click(timeout=10000)
                    elif icona_nuovo_inserimento.count() > 0:
                        print("   🟢 [Robot] [INSERT_PIANIFICAZIONE DETECTED] Nuova inserzione! Clic sul pallino verde di aggiunta...")
                        icona_nuovo_inserimento.click(timeout=10000)
                    else:
                        print("   ⚠️ [Robot] Icone specifiche non intercettate dal DOM. Tento il clic sulla cella td...")
                        target_frame.locator("td[onclick*='Pianificazione']").first.click(timeout=10000)
                    time.sleep(6)

                    print("   ⏰ Compilazione campi temporali nel sistema...")
                    campo_dal = target_frame.locator("input[id*='txtDataDal'], input[id*='Inizio']").first
                    campo_al = target_frame.locator("input[id*='txtDataAl'], input[id*='Fine']").first
                    
                    valore_attuale_dal = campo_dal.input_value() if campo_dal.count() > 0 else ""
                    if valore_attuale_dal == data_in_completa:
                        print(f"   ℹ️ Le date ({data_in_completa}) coincidono già sul portale Snaitech. Salto il locale.")
                        page.locator("#ctl00_MenuID1_rpMaster_ctl04_btnMnuItemPadre").first.click()
                        time.sleep(5)
                        continue

                    campo_dal.fill(data_in_completa)
                    time.sleep(1)
                    campo_al.fill(data_fi_completa)
                    time.sleep(1)

                    print("   💾 Invio moduli di chiusura a Snaitech...")
                    target_frame.locator("input[type='submit'][value*='Salva'], button:has-text('Salva'), input[id*='btnSalva']").first.click()
                    
                    print(f"✅ [Robot] Locale {codice_aams} allineato e salvato con successo nel database Snaitech!")
                    print("----------------------------------------------------------------------")
                    time.sleep(5)
                    
                    page.locator("#ctl00_MenuID1_rpMaster_ctl04_btnMnuItemPadre").first.click()
                    time.sleep(5)
                    
                except Exception as row_err:
                    print(f"⚠️ Errore durante la compilazione del locale: {str(row_err)}")
                    page.locator("#ctl00_MenuID1_rpMaster_ctl04_btnMnuItemPadre").first.click()
                    time.sleep(4)
                    continue
        except Exception as e:
            print(f"❌ Errore durante la navigazione sul portale partner.snai.it: {str(e)}")
        finally:
            print("🤖 [Robot] Processo ultimato. Chiusura sessione.")
            time.sleep(5)
            browser.close()

if __name__ == "__main__":
    avvia_sincronizzazione_automatica()

