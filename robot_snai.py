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
            print("✅ [Robot] STEP 1a: Database Excel intercettato con successo!")
            return pd.read_excel(nome_file_locale).fillna("")
    except Exception: pass
    return pd.DataFrame()

def genera_codice_otp_automatico():
    chiave_pulita = CHIAVE_SEGRETA_2FA.strip().upper().replace(" ", "")
    totp = pyotp.TOTP(chiave_pulita)
    return totp.now()

def avvia_sincronizzazione_automatica():
    df_ferie = preleva_storico_diretto_da_cloud()
    if df_ferie.empty: return

    # 🛡️ FILTRO REALE DI MANUELA: Isola ed elabora esclusivamente i locali sotto "Snaitech Spa WG"
    df_snai = df_ferie[
        df_ferie["CONCESSIONARIO"].astype(str).str.strip() == "Snaitech Spa WG"
    ]
    if df_snai.empty: 
        print("ℹ️ [Robot] Nessun locale trovato per il concessionario specifico 'Snaitech Spa WG'.")
        return

    print(f"🤖 [Robot] STEP 3: Rilevati {len(df_snai)} locali ufficiali Snaitech Spa WG. Avvio Chrome...")

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
            
            print("🔓 [Robot] STEP 5: ACCESSO EFFETTUATO CON SUCCESSO SUL PORTALE SNAITECH!")
            print("----------------------------------------------------------------------")

            print("📬 [Robot] STEP 6: Spostamento singolo sulla pagina degli Esercizi censiti...")
            page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", wait_until="load", timeout=40000)
            print("   ⏳ [Robot] STEP 6a: Attesa stabilizzazione della pagina (10 secondi)...")
            time.sleep(10)

            for _, row in df_snai.iterrows():
                try:
                    codice_aams = str(row["CODICE_LOCALE"]).strip()
                    nome_locale_corrente = str(row["NOME_LOCALE"]).strip()
                    data_in_completa = str(row["INIZIO_FERIE"]).strip()
                    data_fi_completa = str(row["FINE_FERIE"]).strip()
                    
                    data_inizio_pulita = str(data_in_completa)
                    data_fine_pulita = str(data_fi_completa)
                    
                    print(f"🚀 [Robot] STEP 7: Avvio lavorazione -> Codice Locale: {codice_aams} - {nome_locale_corrente}")

                    target_frame = page

                    print("   🔍 [Robot] STEP 7a: Inserimento codice censimento nella barra filtri...")
                    campo_ricerca = target_frame.locator("#ctl00_Cp1_txtCodiceCensimentoesercizio").first
                    campo_ricerca.wait_for(state="visible", timeout=25000)
                    campo_ricerca.click()
                    campo_ricerca.fill(codice_aams)
                    time.sleep(2)
                    
                    tasto_ricerca = target_frame.locator("#ctl00_Cp1_btRicerca").first
                    tasto_ricerca.click(timeout=10000)
                    
                    print("   ⏳ [Robot] STEP 7b: Attesa caricamento risultati filtrati (6 secondi)...")
                    time.sleep(6)

# =====================================================================================
# BLOCCO 5: COMPILAZIONE DATE CON BARRE, CLIC SU TASTO SALVA E SCATTO CLOUD SPIA
# =====================================================================================
                    icona_modifica = target_frame.locator("img[src*='edit_pianificazione'], img[src*='edit'], img[id*='img_pianificazione'][src*='gif']").first
                    icona_nuovo = target_frame.locator("img[src*='insert_pianificazione'], img[src*='insert'], img[id*='img_pianificazione'][src*='jpg']").first
                    cella_td_cliccabile = target_frame.locator("td[onclick*='Pianificazione_dettagli'], table[id*='lst'] tr td:nth-child(8)").first
                    
                    if icona_modifica.count() > 0:
                        print("   📝 [Robot] STEP 8: [MODIFICA DETECTED] Entro nella pianificazione (Pop-up OK)...")
                        icona_modifica.click(force=True, timeout=8000)
                    elif icona_nuovo.count() > 0:
                        print("   🟢 [Robot] STEP 8a: [NUOVA CHIUSURA DETECTED] Clic sul pallino verde...")
                        icona_nuovo.click(force=True, timeout=8000)
                    elif cella_td_cliccabile.count() > 0:
                        print("   🖱️ [Grid Mode] Clic diretto sulla cella TD nativa della colonna 8...")
                        cella_td_cliccabile.click(force=True, timeout=8000)
                    else:
                        print("   ⚠️ [Grid Mode] Tento il clic forzato sulla prima immagine della riga...")
                        target_frame.locator("table#rounded-corner tbody tr td img, td[onclick*='Pianificazione'] img").first.click(force=True, timeout=8000)
                    
                    print("   ⏳ [Robot] STEP 8c: Attesa apertura campi date (7 secondi)...")
                    time.sleep(7)

                    campo_dal = page.locator("#ctl00_Cp1_Txtiniziochiusura, input[name*='Txtiniziochiusura'], input[id*='Txtiniziochiusura']").first
                    campo_al = page.locator("#ctl00_Cp1_Txtfinechiusura, input[name*='Txtfinechiusura'], input[id*='Txtfinechiusura']").first
                    
                    if campo_al.count() == 0:
                        campo_al = page.locator("input[id*='chiusura'], input[id*='Al']").nth(1)

                    # 🛡️ CONVERSIONE RIGIDA DATA: Trasforma i trattini dell'Excel (31-08) nelle barre di Snaitech (31/08)
                    data_inizio_barre = str(data_inizio_pulita).replace("-", "/").strip()
                    data_fine_barre = str(data_fine_pulita).replace("-", "/").strip()

                    campo_dal.wait_for(state="visible", timeout=12000)
                    campo_dal.click()
                    campo_dal.fill(data_inizio_barre)
                    time.sleep(1)
                    
                    try:
                        target_frame.locator("#ctl00_Cp1_fascia_from, select[name*='fascia_from']").select_option("00:00")
                        time.sleep(1)
                    except Exception: pass
                    
                    campo_al.click()
                    campo_al.fill(data_fine_barre)
                    time.sleep(1)
                    
                    try:
                        target_frame.locator("#ctl00_Cp1_fascia_to, select[name*='fascia_to']").select_option("23:30")
                        time.sleep(1)
                    except Exception: pass

                    print("   💾 [Robot] STEP 10: Invio moduli di chiusura a Snaitech (Clic su Tasto Salva)...")
                    page.locator("#ctl00_Cp1_BtnOk").first.click(timeout=10000)
                    print(f"   ✅ [Robot] STEP 11: Impulso inviato con successo sul portale Snaitech.")
                    time.sleep(6)
                    
                    # 🛡️ RESET DI NAVIGAZIONE VINCENTE: Ricarica l'URL pulito ignorando il tasto Indietro grafico
                    print("   ↩️ [Robot] Ricarico la pagina anagrafica pulita per il locale successivo...")
                    page.goto("https://snai.it", wait_until="load", timeout=30000)
                    time.sleep(6)
                    
                except Exception as row_err:
                    print(f"   ⚠️ Nota compilazione: Scavalco riga. Errore: {str(row_err)}")
                    try:
                        page.goto("https://snai.it", wait_until="load", timeout=30000)
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

