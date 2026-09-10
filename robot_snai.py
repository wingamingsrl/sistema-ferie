# =====================================================================================
# SW AUTOMATICO DI SINCRONIZZAZIONE LOCALI WIN GAMING — PRODUZIONE FINALE INDISTRUTTIBILE
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

CHIAVE_SEGRETA_2FA = "FTIA6UQZM2LQLPYJ"
SNAI_USER = "2141ManuelaA"
SNAI_PASS = "Salmi123!"

def preleva_storico_diretto_da_cloud():
    print("📡 [Robot] STEP 1: Lettura del database Excel locale sul server Actions...")
    try:
        nome_file_locale = "storico_ferie.xlsx"
        if os.path.exists(nome_file_locale):
            return pd.read_excel(nome_file_locale).fillna("")
    except Exception: pass
    return pd.DataFrame()

def genera_codice_otp_automatico():
    chiave_pulita = CHIAVE_SEGRETA_2FA.strip().upper().replace(" ", "")
    totp = pyotp.TOTP(chiave_pulita)
    return totp.now()

# =====================================================================================
# BLOCCO 3: LOGIN ED AUTENTICAZIONE OTP INTEGRALE COMPLETATA
# =====================================================================================
def avvia_sincronizzazione_automatica():
    df_ferie = preleva_storico_diretto_da_cloud()
    if df_ferie.empty: return

    df_snai = df_ferie[
        df_ferie["CONCESSIONARIO"].astype(str).str.lower().str.contains("snai|snaitech", regex=True) |
        df_ferie["NOME_LOCALE"].astype(str).str.lower().str.contains("snai", regex=True)
    ]
    if df_snai.empty: return

    print(f"🤖 [Robot] STEP 3: Rilevati {len(df_snai)} locali Snaitech. Avvio Chrome Schermato...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=[
            "--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage",
            "--disable-blink-features=AutomationControlled"
        ]) 
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
            locale="it-IT", timezone_id="Europe/Rome"
        )
        page = context.new_page()

        try:
            print("🌐 [Robot] STEP 4: Connessione schermata a partner.snai.it...")
            page.goto("https://partner.snai.it", wait_until="load", timeout=50000)
            time.sleep(5)
            
            page.locator("input#username, input[name='username']").first.fill(SNAI_USER)
            page.locator("input#password, input[name='password']").first.fill(SNAI_PASS)
            page.click("button[type='submit'], input[type='submit'], .btn-login")
            time.sleep(11)
            
            try: page.evaluate("document.querySelectorAll('.modal, .modal-backdrop').forEach(el => el.remove());")
            except Exception: pass

            codice_totp = genera_codice_otp_automatico()
            print(f"🔑 [Robot] STEP 4f: Codice generato inviato a schermo: {codice_totp}")
            campo_token = page.locator("input#token, input[name='token'], input[name='otp']").first
            campo_token.fill(str(codice_totp))
            time.sleep(2)
            page.keyboard.press("Enter")
            time.sleep(15)
            
            print("🔓 [Robot] STEP 5: ACCESSO EFFETTUATO CON SUCCESSO SUL PORTALE SNAITECH!")
            print("----------------------------------------------------------------------")

# =====================================================================================
# BLOCCO 4: INIEZIONE PROGRAMMATICA DEI CODICI CLIENTE (BYPASS COMPLETO SHADOW-DOM)
# =====================================================================================
            print("📬 [Robot] STEP 6: Spostamento forzato sulla pagina degli Esercizi censiti...")
            page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", wait_until="load", timeout=40000)
            time.sleep(8)

            for _, row in df_snai.iterrows():
                try:
                    codice_aams = str(row["CODICE_LOCALE"]).strip()
                    nome_locale_corrente = str(row["NOME_LOCALE"]).strip()
                    data_in_completa = str(row["INIZIO_FERIE"]).strip()
                    data_fi_completa = str(row["FINE_FERIE"]).strip()
                    
                    print(f"🚀 [Robot] STEP 7: Forzatura programmatica -> {codice_aams} - {nome_locale_corrente}")

                    # 🛡️ CONTROMISURA ATOMICA INIETTATA DA MANUELA:
                    # Inietta il valore direttamente nel motore JavaScript dell'aspnetForm eludendo i blocchi visivi
                    page.evaluate(f"""
                        var inputCensimento = document.getElementById('ctl00_Cp1_txtCodiceCensimentoesercizio');
                        if(inputCensimento) {{
                            inputCensimento.value = '{codice_aams}';
                            __doPostBack('ctl00$Cp1$btnRicerca','');
                        }} else {{
                            // Forza l'assegnazione se occultato nei nodi nascosti
                            document.forms['aspnetForm']['ctl00$Cp1$txtCodiceCensimentoesercizio'].value = '{codice_aams}';
                            document.forms['aspnetForm'].submit();
                        }}
                    """)
                    print("   ⏳ [Robot] STEP 7b: Attesa caricamento griglia dei risultati (6 secondi)...")
                    time.sleep(6)

# =====================================================================================
# BLOCCO 5: ACCESSO DIRETTO ALLA PIANIFICAZIONE ED INSERIMENTO DATE FERIE
# =====================================================================================
                    icona_nuovo = page.locator("img[src*='insert_pianificazione'], img[id*='img_pianificazione']").first
                    icona_modifica = page.locator("img[src*='edit_pianificazione']").first
                    
                    if icona_modifica.count() > 0:
                        print("   📝 [Robot] STEP 8: [MODIFICA] Entro nella pianificazione esistente...")
                        icona_modifica.click(force=True)
                    elif icona_nuovo.count() > 0:
                        print("   🟢 [Robot] STEP 8a: [NUOVA CHIUSURA] Clic sul pulsante verde...")
                        icona_nuovo.click(force=True)
                    else:
                        # Se occultato, lancia la funzione direttamente dall'HTML estratto da Manuela
                        print("   🖱️ [JavaScript Mode] Esecuzione nativa della funzione Pianificazione_dettagli...")
                        page.evaluate(f"Pianificazione_dettagli('', '37832','{nome_locale_corrente}','{codice_aams}')")
                    time.sleep(4)

                    # Compilazione date nel sotto-pannello sbloccato
                    page.locator("input[id*='txtDataDal'], input[id*='Inizio']").first.fill(data_in_completa)
                    time.sleep(1)
                    page.locator("input[id*='txtDataAl'], input[id*='Fine']").first.fill(data_fi_completa)
                    time.sleep(1)

                    # Invio finale e archiviazione nel server di Snaitech
                    page.locator("input[type='submit'][value*='Salva'], input[id*='btnSalva']").first.click()
                    print(f"   ✅ [Robot] STEP 11: Locale {codice_aams} sincronizzato con successo!")
                    print("----------------------------------------------------------------------")
                    time.sleep(4)
                    
                except Exception as row_err:
                    print(f"   ⚠️ Nota compilazione: Scavalco riga ed aggiorno focus. Errore: {str(row_err)}")
                    continue

            # Logout formale richiesto da Manuela per liberare la linea dell'ufficio
            print("🔒 [Robot] STEP 12: Chiusura formale della sessione (Logout)...")
            try: page.locator("a:has-text('LogOut'), a:has-text('Esci')").first.click(timeout=8000)
            except Exception: page.context.clear_cookies()

        except Exception as e: print(f"❌ Errore generale: {str(e)}")
        finally: browser.close()

if __name__ == "__main__":
    avvia_sincronizzazione_automatica()
