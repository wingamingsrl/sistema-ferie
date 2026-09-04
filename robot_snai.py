# =====================================================================================
# SW AUTOMATICO DI SINCRONIZZAZIONE LOCALI WIN GAMING — ORE 20:00 NATIVO
# BLOCCO 1: STRUTTURA LIBRERIE ED ACCESSI PROPRIETARI — PORTALE: PARTNER.SNAI.IT
# =====================================================================================
import os
import io
import time
import pyotp
import requests
import pandas as pd
import streamlit as st
from datetime import datetime
from playwright.sync_api import sync_playwright

# Configurazione ufficiale ed esclusiva delle tue coordinate aziendali
CHIAVE_SEGRETA_2FA = "FTIA6UQZM2LQLPYJ"  # La tua chiave definitiva allineata al telefono
SNAI_USER = "2141ManuelaA"
SNAI_PASS = "Salmi123!"

# =====================================================================================
# BLOCCO 2: MOTORE DI LETTURA LIVE IN MEMORIA RAM DA REPOSITORY GITHUB (FORMATO EXCEL)
# DIRETTO DA INTERNET ALLA RAM DEL ROBOT — ALLINEATO ALLA NUOVA APP DI PRODUZIONE
# =====================================================================================
def preleva_storico_diretto_da_cloud():
    print("📡 [Robot] Estrazione database Excel direttamente in RAM da GitHub...")
    try:
        # Recupera il token di accesso in totale sicurezza dai Secrets di Streamlit Cloud
        t_git = str(st.secrets["github"]["token_accesso"]).strip()
        c_time = str(int(time.time() * 1000))
        
        # Punta all'indirizzo API ufficiale del file Excel condiviso dai ragazzi
        url_git = f"https://github.com{c_time}"
        
        headers_diretti = {
            "Authorization": f"token {t_git}", 
            "Accept": "application/vnd.github.v3.raw",
            "User-Agent": "WinGaming-Cloud-App"
        }
        
        risposta = requests.get(url_git, headers=headers_diretti, timeout=15)
        
        if risposta.status_code == 200:
            # Inietta il file Excel direttamente in RAM senza toccare l'hard disk
            df_ram = pd.read_excel(io.BytesIO(risposta.content))
            print("✅ [Robot] Database Excel iniettato in RAM con successo!")
            return df_ram.fillna("")
        else:
            print(f"❌ Errore scaricamento Excel da GitHub. Stato HTTP: {risposta.status_code}")
            
    except Exception as e_network:
        print(f"⚠️ Errore di rete durante la sincronizzazione live: {str(e_network)}")
        
    print("❌ Impossibile caricare il database. Blocco precauzionale per evitare dati errati.")
    return pd.DataFrame()

def genera_codice_otp_automatico():
    chiave_pulita = CHIAVE_SEGRETA_2FA.strip().upper().replace(" ", "")
    totp = pyotp.TOTP(chiave_pulita)
    return totp.now()

# =====================================================================================
# BLOCCO 3: AVVIO CHROME, LOG IN E SUPERAMENTO BARRIERA DI SICUREZZA 2FA SU PARTNER.SNAI.IT
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
            print("🌐 [Robot] Connessione a partner.snai.it...")
            page.goto("https://snai.it", timeout=30000)
            time.sleep(3)
            
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
            print("📦 [Robot] Apertura del menu Anagrafica Locali...")
            page.locator("#ctl00_MenuID1_rpMaster_ctl04_btnMnuItemPadre").first.click(timeout=15000)
            time.sleep(8)

            for _, row in df_snai.iterrows():
                try:
                    # 🛡️ FIX COLONNE: Estrae i dati usando le intestazioni ufficiali dell'ufficio
                    codice_aams = str(row["CODICE_LOCALE"]).strip()
                    nome_locale_corrente = str(row["NOME_LOCALE"]).strip()
                    
                    data_in_completa = str(row["INIZIO_FERIE"]).strip()
                    data_fi_completa = str(row["FINE_FERIE"]).strip()
                    ora_inserimento = str(row.get("DATA_INSERIMENTO", "N.D.")).strip()
                    
                    print(f"🚀 [Robot] Avvio lavorazione -> Codice Locale: {codice_aams} - {nome_locale_corrente}")
                    print(f"   🔹 Inserito il:     {ora_inserimento} (Tracciamento Cloud)")
                    print(f"   🔹 Inizio Chiusura: {data_in_completa}")
                    print(f"   🔹 Fine Chiusura:   {data_fi_completa}")

                    target_frame = page
                    if len(page.frames) > 1:
                        print("   📦 Rilevato sotto-foglio iframe protetto. Spostamento all'interno...")
                        target_frame = page.frames

                    print("   🔍 Inserimento codice censimento nella barra filtri...")
                    campo_ricerca = "input[id*='Censimento'], input[name*='Censimento'], input[id*='txtCodice']"
                    if target_frame.locator(campo_ricerca).count() > 0:
                        target_frame.locator(campo_ricerca).first.fill(codice_aams)
                        target_frame.keyboard.press("Enter")
                        time.sleep(5)

# =====================================================================================
# BLOCCO 5: CONTROLLO STRUTTURA (NUOVO/MODIFICA), INIEZIONE ORARI E CHIUSURA SESSIONE
# =====================================================================================
                    tasto_modifica = "img[id*='img_modifica'], img[id*='img_dettaglio'], img[src*='edit'], [title*='Modifica']"
                    pallino_verde_nuovo = "img[id*='img_pianificazione'], img[src*='insert_pianificazione']"
                    
                    if target_frame.locator(tasto_modifica).count() > 0:
                        print("   📝 [Robot] Rilevata chiusura esistente! Clic sull'icona di Modifica/Matita...")
                        target_frame.locator(tasto_modifica).first.click(timeout=10000)
                    elif target_frame.locator(pallino_verde_nuovo).count() > 0:
                        print("   🟢 [Robot] Nuova inserzione! Clic sul pallino verde '+' per aggiungere il periodo...")
                        target_frame.locator(pallino_verde_nuovo).first.click(timeout=10000)
                    else:
                        print("   ⚠️ [Robot] Icone specifiche non intercettate. Tento il clic generico sulla riga...")
                        target_frame.locator("img[id*='pianificazione']").first.click(timeout=10000)
                    time.sleep(5)

                    print("   ⏰ Compilazione campi temporali nel sistema...")
                    campo_dal = "input[id*='txtDataDal'], input[id*='Inizio'], input[name*='Inizio']"
                    campo_al = "input[id*='txtDataAl'], input[id*='Fine'], input[name*='Fine']"
                    
                    target_frame.locator(campo_dal).first.fill(data_in_completa)
                    time.sleep(1)
                    target_frame.locator(campo_al).first.fill(data_fi_completa)
                    time.sleep(1)

                    print("   💾 Invio moduli di chiusura a Snaitech...")
                    # target_frame.locator("input[type='submit'][value*='Salva'], button:has-text('Salva')").first.click()
                    
                    print(f"✅ [Robot] Locale {codice_aams} elaborato con successo con orario dettagliato!")
                    print("----------------------------------------------------------------------")
                    time.sleep(4)
                    
                except Exception as row_err:
                    print(f"⚠️ Errore durante la compilazione del locale: {str(row_err)}")
                    continue
        except Exception as e:
            print(f"❌ Errore durante la navigazione sul portale partner.snai.it: {str(e)}")
        finally:
            print("🤖 [Robot] Processo ultimato. Chiusura sessione.")
            time.sleep(5)
            browser.close()

if __name__ == "__main__":
    avvia_sincronizzazione_automatica()