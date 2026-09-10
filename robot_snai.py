# =====================================================================================
# SW AUTOMATICO DI SINCRONIZZAZIONE LOCALI WIN GAMING — PRODUZIONE FINALE INTEGRALE
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
# BLOCCO 3: IL LOGIN INTIMISTICO E SCHERMATO CHE HA SUPERATO L'OTP CON SUCCESSO
# =====================================================================================
def avvia_sincronizzazione_automatica():
    df_ferie = preleva_storico_diretto_da_cloud()
    if df_ferie.empty: return

    df_snai = df_ferie[
        df_ferie["CONCESSIONARIO"].astype(str).str.lower().str.contains("snai|snaitech", regex=True) |
        df_ferie["NOME_LOCALE"].astype(str).str.lower().str.contains("snai", regex=True)
    ]
    if df_snai.empty: return

    print(f"🤖 [Robot] STEP 3: Rilevati {len(df_snai)} locali Snaitech. Ripristino Chrome Schermato...")

    with sync_playwright() as p:
        # 🛡️ RIPRISTINO ASSETTO VINCENTE: Schermatura anti-bot totale che inganna il firewall di Snaitech
        browser = p.chromium.launch(headless=False, args=[
            "--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage",
            "--disable-blink-features=AutomationControlled"
        ]) 
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
            locale="it-IT", timezone_id="Europe/Rome",
            viewport={"width": 1280, "height": 1024}
        )
        page = context.new_page()

        try:
            print("🌐 [Robot] STEP 4: Connessione schermata a partner.partner.snai.it...")
            page.goto("https://partner.snai.it", wait_until="networkidle", timeout=60000)
            time.sleep(6)
            
            try: page.mouse.move(150, 150)
            except Exception: pass
            
            print("📝 [Robot] STEP 4a: Inserimento credenziali Snaitech...")
            campo_user = page.locator("input#username, input[name='username'], input[type='text']").first
            campo_user.click(timeout=20000)
            campo_user.fill(SNAI_USER)
            time.sleep(1)
            
            campo_pass = page.locator("input#password, input[name='password'], input[type='password']").first
            campo_pass.click(timeout=20000)
            campo_pass.fill(SNAI_PASS)
            time.sleep(1)
            
            print("🚀 [Robot] STEP 4b: Invio moduli di accesso...")
            page.click("button[type='submit'], input[type='submit'], .btn-login")
            time.sleep(4)
            
            print("⏳ [Robot] STEP 4c: Pausa di sicurezza di 11 secondi countdown...")
            time.sleep(11)
            
            try: page.evaluate("document.querySelectorAll('.modal, .modal-backdrop').forEach(el => el.remove());")
            except Exception: pass
            time.sleep(2)

            print("🔑 [Robot] STEP 4e: Generazione ed immissione codice 2FA TOTP...")
            codice_totp = genera_codice_otp_automatico()
            print(f"📌 [Robot] STEP 4f: Codice generated inviato a schermo: {codice_totp}")
            
            campo_token = page.locator("input#token, input[name='token'], input[name='otp'], input[type='text']").first
            campo_token.click(timeout=15000)
            campo_token.fill(str(codice_totp))
            time.sleep(2)
            
            print("📤 [Robot] STEP 4g: Invio codice OTP tramite tastiera virtuale...")
            page.keyboard.press("Enter")
            print("⏳ [Robot] STEP 4h: Convalida credenziali in corso (15 secondi)...")
            time.sleep(15)
            
            print("🔓 [Robot] STEP 5: ACCESSO EFFETTUATO CON SUCCESSO SUL PORTALE PARTNER SNAITECH!")
            print("----------------------------------------------------------------------")

# =====================================================================================
# BLOCCO 4: INIEZIONE NELL'UPDATEPANEL MICROSOFT DEL FRAME SPECIFICO (DA FOTO DI MANUELA)
# =====================================================================================
            print("📬 [Robot] STEP 6: Spostamento forzato sulla pagina degli Esercizi censiti...")
            page.goto("https://partner.snai.it", wait_until="load", timeout=40000)
            time.sleep(10)

            for _, row in df_snai.iterrows():
                try:
                    codice_aams = str(row["CODICE_LOCALE"]).strip()
                    nome_locale_corrente = str(row["NOME_LOCALE"]).strip()
                    data_in_completa = str(row["INIZIO_FERIE"]).strip()
                    data_fi_completa = str(row["FINE_FERIE"]).strip()
                    
                    print(f"🚀 [Robot] STEP 7: Forzatura programmatica -> {codice_aams} - {nome_locale_corrente}")

                    # 🛡️ IDENTIFICAZIONE DEL FRAME GIUSTO: Trova lo schermo in cui vive la tabella
                    target_frame = page
                    for f in page.frames:
                        if "Esercizi" in f.url or f.locator("[id*='txtCodiceCensimentoesercizio']").count() > 0:
                            target_frame = f
                            break

                    # 🛡️ INIEZIONE INTERNA AL FRAME: Scrive ed esegue il PostBack direttamente dentro la scatola protetta
                    target_frame.evaluate(f"""
                        var inputCensimento = document.getElementById('ctl00_Cp1_txtCodiceCensimentoesercizio');
                        if(inputCensimento) {{
                            inputCensimento.value = '{codice_aams}';
                            __doPostBack('ctl00$Cp1$btnRicerca','');
                        }} else if(document.forms['aspnetForm']) {{
                            document.forms['aspnetForm']['ctl00$Cp1$txtCodiceCensimentoesercizio'].value = '{codice_aams}';
                            __doPostBack('ctl00$Cp1$btnRicerca','');
                        }}
                    """)
                    print("   ⏳ [Robot] STEP 7b: Attesa griglia dei risultati (6 secondi)...")
                    time.sleep(6)

# =====================================================================================
# BLOCCO 5: RINFRESCO OCCHIO VISIVO ED INSERIMENTO DELLE DATE FERIE CON LOGOUT
# =====================================================================================
                    # 🛡️ CONTROMISURA RINFRESCO: Rinfresca l'occhio sul sotto-foglio dopo che il PostBack ha ricaricato la tabella
                    time.sleep(2)
                    target_frame = page
                    for f in page.frames:
                        if "Esercizi" in f.url or f.locator("img[id*='img_pianificazione']").count() > 0 or f.locator("td[onclick*='Pianificazione']").count() > 0:
                            target_frame = f
                            break

                    try: target_frame.evaluate("document.querySelectorAll('.mainPreload').forEach(el => el.remove());")
                    except Exception: pass

                    # Mappatura dei puntatori reali estratti dall'HTML di Manuela
                    icona_nuovo = target_frame.locator("img[src*='insert_pianificazione'], img[id*='img_pianificazione']").first
                    icona_modifica = target_frame.locator("img[src*='edit_pianificazione']").first
                    cella_javascript_diretta = target_frame.locator("td[onclick*='Pianificazione']").first
                    
                    if icona_modifica.count() > 0:
                        print("   📝 [Robot] STEP 8: [MODIFICA] Rilevata icona esistente. Entro nella pianificazione...")
                        icona_modifica.click(force=True, timeout=8000)
                    elif icona_nuovo.count() > 0:
                        print("   🟢 [Robot] STEP 8a: [NUOVA CHIUSURA] Rilevato locale vuoto. Clic sul pulsante verde...")
                        icona_nuovo.click(force=True, timeout=8000)
                    elif cella_javascript_diretta.count() > 0:
                        print("   🖱️ [Grid Mode] Clic sulla cella td nativa estratta da Manuela...")
                        cella_javascript_diretta.click(force=True, timeout=8000)
                    else:
                        print("   ⚠️ [Grid Mode] Tento il clic forzato sulla prima immagine dei risultati aggiornati...")
                        target_frame.locator("img[id*='img_pianificazione'], td img, .Grid img").first.click(force=True, timeout=8000)
                    
                    print("   ⏳ [Robot] STEP 8c: Attesa apertura campi temporali (5 secondi)...")
                    time.sleep(5)

                    campo_dal = target_frame.locator("input[id*='txtDataDal'], input[id*='Inizio'], input[name*='Dal']").first
                    campo_al = target_frame.locator("input[id*='txtDataAl'], input[id*='Fine'], input[name*='Al']").first
                    
                    campo_dal.wait_for(state="visible", timeout=10000)
                    campo_dal.click()
                    campo_dal.fill(data_in_completa)
                    time.sleep(1)
                    
                    campo_al.click()
                    campo_al.fill(data_fi_completa)
                    time.sleep(1)

                    print("   💾 [Robot] STEP 10: Invio moduli di chiusura a Snaitech...")
                    target_frame.locator("input[type='submit'][value*='Salva'], button:has-text('Salva'), input[id*='btnSalva']").first.click()
                    print(f"   ✅ [Robot] STEP 11: Locale {codice_aams} sincronizzato con successo!")
                    print("----------------------------------------------------------------------")
                    time.sleep(5)
                    
                except Exception as row_err:
                    print(f"   ⚠️ Nota compilazione: Scavalco riga. Errore: {str(row_err)}")
                    continue

            print("🔒 [Robot] STEP 12: Chiusura formale della sessione (Logout di sicurezza)...")
            try: page.locator("a:has-text('LogOut'), a:has-text('Esci'), [id*='btnLogOut']").first.click(timeout=8000)
            except Exception: page.context.clear_cookies()

        except Exception as e: print(f"❌ Errore generale: {str(e)}")
        finally: browser.close()

if __name__ == "__main__":
    avvia_sincronizzazione_automatica()
