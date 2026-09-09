# =====================================================================================
# SW AUTOMATICO DI SINCRONIZZAZIONE LOCALI WIN GAMING — PRODUZIONE FINALE ONLINE
# BLOCCO 1: STRUTTURA LIBRERIE ED ACCESSI PROPRIETARI — PORTALE: PARTNER.
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
# BLOCCO 3: AVVIO CHROME, LOG IN E SUPERAMENTO BARRIERA DI SICUREZZA 2FA SU PARTNER.
# =====================================================================================
def avvia_sincronizzazione_automatica():
    df_ferie = preleva_storico_diretto_da_cloud()
    if df_ferie.empty:
        print("❌ Impossibile procedere: Il database delle ferie è vuoto o bloccato.")
        return

    # 🛡️ ALLINEAMENTO CONCESSIONARIO: Isola i record che contengono Snai nelle colonne ufficiali
    df_snai = df_ferie[
        df_ferie["CONCESSIONARIO"].astype(str).str.lower().str.contains("snai|snaitech", regex=True) |
        df_ferie["NOME_LOCALE"].astype(str).str.lower().str.contains("snai", regex=True)
    ]

    if df_snai.empty:
        print("✅ [Robot] Nessun locale Snaitech attivo trovato nel registro. Sincronizzazione conclusa.")
        return

    print(f"🤖 [Robot] Rilevati {len(df_snai)} locali Snaitech da elaborare. Avvio Chrome...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False) 
        context = browser.new_context()
        page = context.new_page()

        try:
            print("🌐 [Robot] Connessione a partner....")
            # 🛡️ UNICA MODIFICA: 'load' al posto di 'networkidle' per evitare il blocco dei 60 secondi
            page.goto("https://partner.snai.it", wait_until="load", timeout=45000)
            time.sleep(4)
            
            print("📝 [Robot] Inserimento credenziali Snaitech...")
            page.fill("input#username, input[name='username'], input[type='text']", SNAI_USER)
            page.fill("input#password, input[name='password'], input[type='password']", SNAI_PASS)
            
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
            print(f"📌 Codice generated inviato a schermo: {codice_totp}")
            
            page.fill("input#token, input[name='token'], input[name='otp']", codice_totp)
            time.sleep(1)
            
            page.click("input#btnInvia, input[value='Invia'], button:has-text('Invia'), input[type='submit']")
            print("⏳ [Robot] Convalida credenziali in corso... Caricamento area riservata partner.snai.it...")
            time.sleep(15)
            
            print("🔓 [Robot] ACCESSO EFFETTUATO CON SUCCESSO SUL PORTALE PARTNER SNAITECH!")
            print("----------------------------------------------------------------------")

# =====================================================================================
# BLOCCO 4: INTERCETTAZIONE MENU ANAGRAFICA E FILTRAGGIO CODICI CENSIMENTO
# =====================================================================================
            print("📬 [Robot] STEP 6: Spostamento forzato sulla pagina degli Esercizi censiti...")
            page.goto("https://partner.snai.it", wait_until="load", timeout=40000)
            # 🛡️ FIX CONTENITORE ESTRATTO DA MANUELA: Attende la comparsa reale dell'UpdatePanel dinamico di Snaitech
            print("   ⏳ [Robot] STEP 6a: Attesa rendering del pannello Microsoft UpdatePanel...")
            pannello_ricerca = page.locator("#ctl00_Cp1_updPnlSearchResult, div[id*='updPnlSearchResult']").first
            pannello_ricerca.wait_for(state="visible", timeout=20000)
            time.sleep(5)

            for _, row in df_snai.iterrows():
                try:
                    codice_aams = str(row["CODICE_LOCALE"]).strip()
                    nome_locale_corrente = str(row["NOME_LOCALE"]).strip()
                    data_in_completa = str(row["INIZIO_FERIE"]).strip()
                    data_fi_completa = str(row["FINE_FERIE"]).strip()
                    
                    print(f"🚀 [Robot] STEP 7: Avvio lavorazione -> Codice Locale: {codice_aams} - {nome_locale_corrente}")

                    # Puntiamo al pannello di ricerca estratto da Manuela
                    target_frame = page

                    print("   🔍 [Robot] STEP 7a: Inserimento codice censimento nella barra filtri...")
                    # Puntatore laser sul campo txtCodiceCensimentoesercizio che vive dentro l'UpdatePanel
                    campo_ricerca = target_frame.locator("#ctl00_Cp1_txtCodiceCensimentoesercizio, input[id*='txtCodiceCensimentoesercizio'], input[name*='txtCodiceCensimentoesercizio']").first
                    
                    # Forza l'attesa di stabilità della casella di testo
                    campo_ricerca.wait_for(state="visible", timeout=15000)
                    campo_ricerca.click()
                    campo_ricerca.fill(codice_aams)
                    time.sleep(2)
                    
                    # Pressione del tasto Ricerca visivo
                    tasto_ricerca = target_frame.locator("input[type='submit'][value='Ricerca'], input[value='Ricerca'], input[id*='Ricerca']").first
                    tasto_ricerca.click(timeout=10000)
                    
                    print("   ⏳ [Robot] STEP 7b: Attesa griglia dei risultati (6 secondi)...")
                    time.sleep(6)
# =====================================================================================
# BLOCCO 5: CONTROLLO STRUTTURA (NUOVO/MODIFICA) ALLINEATO ALL'HTML DI MANUELA E CHIUSURA
# =====================================================================================
                    icona_nuovo_inserimento = target_frame.locator("img[src*='insert_pianificazione'], img[id*='img_pianificazione']").first
                    icona_modifica_esistente = target_frame.locator("img[src*='edit_pianificazione']").first
                    
                    if icona_modifica_esistente.count() > 0:
                        print("   📝 [Robot] STEP 8: [EDIT_PIANIFICAZIONE DETECTED] Clic sull'icona di Modifica...")
                        icona_modifica_esistente.click(timeout=10000)
                    elif icona_nuovo_inserimento.count() > 0:
                        print("   🟢 [Robot] STEP 8a: [INSERT_PIANIFICAZIONE DETECTED] Nuovo locale vuoto, inserisco...")
                        icona_nuovo_inserimento.click(timeout=10000)
                    else:
                        print("   ⚠️ [Robot] STEP 8b: Icone specifiche non isolate dal DOM. Tento il clic td...")
                        target_frame.locator("td[onclick*='Pianificazione']").first.click(timeout=10000)
                    time.sleep(6)

                    print("   ⏰ [Robot] STEP 9: Compilazione campi temporali nel sistema...")
                    campo_dal = target_frame.locator("input[id*='txtDataDal'], input[id*='Inizio']").first
                    campo_al = target_frame.locator("input[id*='txtDataAl'], input[id*='Fine']").first
                    
                    valore_attuale_dal = campo_dal.input_value() if campo_dal.count() > 0 else ""
                    if valore_attuale_dal == data_in_completa:
                        print(f"   ℹ️ [Robot] STEP 9a: Le date ({data_in_completa}) coincidono già sul portale. Salto.")
                        continue

                    campo_dal.fill(data_in_completa)
                    time.sleep(1)
                    campo_al.fill(data_fi_completa)
                    time.sleep(1)

                    print("   💾 [Robot] STEP 10: Invio moduli di chiusura a Snaitech...")
                    target_frame.locator("input[type='submit'][value*='Salva'], button:has-text('Salva'), input[id*='btnSalva']").first.click()                
                    print(f"✅ [Robot] STEP 11: Locale {codice_aams} allineato e salvato con successo nel database Snaitech!")
                    print("----------------------------------------------------------------------")
                    time.sleep(5)
                    
                except Exception as row_err:
                    print(f"⚠️ [Robot] STEP ERRORE: Scavalco riga. Errore: {str(row_err)}")
                    continue
        except Exception as e:
            print(f"❌ [Robot] ERRORE GENERALE DI NAVIGAZIONE: {str(e)}")
        finally:
            print("🤖 [Robot] Processo ultimato. Chiusura sessione.")
            time.sleep(5)
            browser.close()

if __name__ == "__main__":
    avvia_sincronizzazione_automatica()
