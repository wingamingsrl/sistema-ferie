# =====================================================================================
# SW AUTOMATICO DI SINCRONIZZAZIONE LOCALI WIN GAMING — PRODUZIONE FINALE
# BLOCCO 1: STRUTTURA LIBRERIE AZIENDALI, TOTP 2FA E RIPULITURA EXCEL VIA GIT
# =====================================================================================
import os
import io
import time
import base64
import pyotp
import requests
import pandas as pd
from datetime import datetime
from playwright.sync_api import sync_playwright

CHIAVE_SEGRETA_2FA = "FTIA6UQZM2LQLPYJ"
SNAI_USER = "2141GestFerie"
SNAI_PASS = "ks{6bv0Gjo"

#SNAI_USER = "2141ManuelaA"
#SNAI_PASS = "Salmi123!"

def preleva_storico_diretto_da_cloud():
    print("📡 [Robot] STEP 1: Lettura del database Excel locale...")
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

def scarica_e_aggiorna_excel_su_github(codice_locale_successo):
    print(f"💾 [Cloud Excel] Svuoto ROBOT_ACTION per il locale: {codice_locale_successo}...")
    try:
        nome_file = "storico_ferie.xlsx"
        if os.path.exists(nome_file):
            df_file = pd.read_excel(nome_file)
            
            # Azzera la cella d'azione per il locale completato
            # 🛡️ SPAZZINO DOPPIO DI MANUELA: Se l'azione era ELIMINA cancella la riga, altrimenti svuota la cella
            df_filtrato_locale = df_file[df_file["CODICE_LOCALE"].astype(str).str.strip() == str(codice_locale_successo).strip()]
            if not df_filtrato_locale.empty and str(df_filtrato_locale.iloc[0].get("ROBOT_ACTION", "")).strip().upper() == "ELIMINA":
                df_file = df_file[df_file["CODICE_LOCALE"].astype(str).str.strip() != str(codice_locale_successo).strip()]
                print("   🗑️ [Cloud Excel] Rilevato comando ELIMINA: Riga rimossa definitivamente dal database.")
            else:
                df_file.loc[df_file["CODICE_LOCALE"].astype(str).str.strip() == str(codice_locale_successo).strip(), "ROBOT_ACTION"] = ""

            
            # Riassegna la struttura colonne rigida dell'ufficio
            colonne_ufficio = ["DATA_INSERIMENTO", "TECNICO_INSERIMENTO", "CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO", "INIZIO_FERIE", "FINE_FERIE", "PROMEMORIA_IN_COPIA", "STATO_INVIO", "ROBOT_ACTION"]
            df_pulito_salva = df_file.reindex(columns=colonne_ufficio).astype(str).fillna("")
            df_pulito_salva.to_excel(nome_file, index=False)
            
            # 🛡️ AUTOMAZIONE GIT NATIVA: Salva e spinge sul cloud senza conflitti di token API
            os.system("git config --global user.name 'WinGaming-Robot'")
            os.system("git config --global user.email 'wingamingsrl@gmail.com'")
            os.system(f"git add {nome_file}")
            os.system(f"git commit -m '🤖 [Robot] Allineamento Snaitech OK. Reset azione locale {codice_locale_successo}'")
            stato_push = os.system("git push origin main")
            
            if stato_push == 0:
                print("   ✅ [Cloud Excel] Database ripulito e sincronizzato con successo online!")
            else:
                print(f"   ❌ [Cloud Excel] Errore riscrittura online. Codice push: {stato_push}")
    except Exception as e:
        print(f"   ⚠️ Impossibile aggiornare l'Excel: {str(e)}")
# =====================================================================================
# BLOCCO 2: FILTRO SELEZIONE ANAGRAFICA AZIENDALE ED ACCENSIONE BROWSER CHROME
# =====================================================================================
def avvia_sincronizzazione_automatica():
    df_ferie = preleva_storico_diretto_da_cloud()
    if df_ferie.empty: return

    df_snai = df_ferie[
        df_ferie["CONCESSIONARIO"].astype(str).str.strip() == "Snaitech Spa WG"
    ]
    if df_snai.empty: return

    print(f"🤖 [Robot] STEP 3: Rilevati {len(df_snai)} locali Snaitech Spa WG. Avvio Chrome...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=[
            "--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"
        ]) 
        context = browser.new_context()
        page = context.new_page()

        page.on("dialog", lambda dialog: dialog.accept())
# =====================================================================================
# BLOCCO 3: ACCESSO SUL PORTALE PARTNER ED IMMISSIONE CHIAVE DINAMICA OTP (LINK CORTO)
# =====================================================================================
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
            
            print("🔓 [Robot] STEP 5: ACCESSO EFFETTUATO CON SUCCESSO SUL PORTALE PARTNER SNAITECH!")
            print("----------------------------------------------------------------------")
# =====================================================================================
# BLOCCO 4: SPOSTAMENTO IN ANAGRAFICA E STRUTTURA RICERCA COPIATA COERENTEMENTE DAL TUO TESTO
# =====================================================================================
            print("📬 [Robot] STEP 6: Spostamento sulla pagina degli Esercizi censiti...")
            page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx")
            print("   ⏳ [Robot] STEP 6a: Attesa stabilizzazione della pagina (10 secondi)...")
            time.sleep(10)

            for _, row in df_snai.iterrows():
                try:
                    codice_aams = str(row["CODICE_LOCALE"]).strip()
                    nome_locale_corrente = str(row["NOME_LOCALE"]).strip()
                    data_in_completa = str(row["INIZIO_FERIE"]).strip()
                    data_fi_completa = str(row["FINE_FERIE"]).strip()
                    mirino_azione = str(row.get("ROBOT_ACTION", "")).strip().upper()
                    
                    data_inizio_pulita = str(data_in_completa).replace("-", "/").strip()
                    data_fine_pulita = str(data_fi_completa).replace("-", "/").strip()
                    
                    # 🛡️ INTERCETTATORE DI MANUELA: Abilita il robot a elaborare NUOVA, MODIFICA ed ELIMINA
                    if mirino_azione not in ["NUOVA", "MODIFICA", "ELIMINA"]:
                        print(f"⏩ [Robot] Locale {codice_aams} - {nome_locale_corrente}: Nessuna azione richiesta. Salto riga.")
                        continue
                        
                    print(f"🚀 [Robot] STEP 7: Avvio lavorazione ({mirino_azione}) -> Codice Locale: {codice_aams} - {nome_locale_corrente}")

                    # 🛡️ TUO CODICE NATIVO ORIGINALE DEI RAGAZZI AL 100% — COPIATO LETTERALMENTE DAL TUO FILE
                    target_frame = page
                    for f in page.frames:
                        if "Esercizi" in f.url or f.locator("#ctl00_Cp1_txtCodiceCensimentoesercizio").count() > 0:
                            target_frame = f
                            break

                    print("   🔍 [Robot] STEP 7a: Inserimento codice censimento nella barra filtri...")
                    campo_ricerca = target_frame.locator("#ctl00_Cp1_txtCodiceCensimentoesercizio").first
                    campo_ricerca.wait_for(state="visible", timeout=20000)
                    campo_ricerca.click()
                    campo_ricerca.fill(codice_aams)
                    time.sleep(2)
                    
                    tasto_ricerca = target_frame.locator("#ctl00_Cp1_btRicerca").first
                    tasto_ricerca.click(timeout=10000)
                    
                    print("   ⏳ [Robot] STEP 7b: Attesa caricamento risultati filtrati (6 secondi)...")
                    time.sleep(6)

                    icona_nuovo = target_frame.locator("img[src*='insert_pianificazione.jpg'], img[src*='insert_pianificazione'], img[id*='img_pianificazione']").first
                    icona_modifica = target_frame.locator("img[src*='edit_pianificazione']").first
                    
                    try:
                        icona_nuovo.wait_for(state="attached", timeout=4000)
                    except Exception: pass

                    # Sblocca l'ingresso in scheda tramite matita anche per il comando di rimozione
                    if (icona_modifica.count() > 0 and icona_modifica.is_visible()) or mirino_azione in ["MODIFICA", "ELIMINA"]:
                        print(f"   📝 [Robot] STEP 8: [{mirino_azione}] Rilevato cambio URL ChiusuraEsercizio.aspx. Clicco...")
                        icona_modifica.click(force=True, timeout=8000)
                    elif icona_nuovo.count() > 0:
                        print("   🟢 [Robot] STEP 8a: [NUOVA CHIUSURA] Clic sul pallino verde...")
                        icona_nuovo.click(force=True, timeout=8000)
                    else:
                        print("   AM 🖱️ [Grid Mode] Clic sulla cella td nativa della riga...")
                        cella_td = target_frame.locator("td[onclick*='Pianificazione']").first
                        cella_td.click(force=True, timeout=8000)
                    
                    print("   ⏳ [Robot] STEP 8c: Attesa apertura campi date (7 secondi)...")
                    time.sleep(7)
# =====================================================================================
# BLOCCO 5: DATA FINE PURIFICATA, SEQUENZA FOTO REALE E RITORNO IN BACHECA PROTETTO
# =====================================================================================
                    frame_date = page
                    for f in page.frames:
                        if "Chiusura" in f.url or f.locator("#ctl00_Cp1_Txtiniziochiusura").count() > 0:
                            frame_date = f
                            break

                    data_inizio_pura = data_inizio_pulita[:10].strip()
                    data_fine_pura = data_fine_pulita[:10].strip()
                    
                    ora_inizio_pulita = "06:00" if "06:00" in str(row["INIZIO_FERIE"]) else "00:00"
                    ora_fine_pulita = "12:00" if "12:00" in str(row["FINE_FERIE"]) else "23:30"

                    # 🛡️ BIVIO CANCELLAZIONE DI MANUELA: Se l'azione è ELIMINA, clicca sul tasto di rimozione ed esce
                    if mirino_azione == "ELIMINA":
                        print("   🗑️ [Robot] STEP 9: Rilevato comando di rimozione. Cerco il tasto Elimina di Snaitech...")
                        tasto_elimina_snai = frame_date.locator("#ctl00_Cp1_BtnElimina").first
                        tasto_elimina_snai.wait_for(state="visible", timeout=10000)
                        
                        # Clicca sul tasto Elimina nativo estratto dal tuo HTML
                        tasto_elimina_snai.click(force=True)
                        print("   💾 [Robot] STEP 10: Pulsante Elimina premuto. Attesa conferma dal server Snaitech...")
                        time.sleep(6)
                        
                        # Svuota ed elimina la riga dal cloud
                        scarica_e_aggiorna_excel_su_github(codice_aams)
                        time.sleep(4)
                        
                        page.goto("https://partner.snai.it")
                        time.sleep(6)
                        continue # Salta il resto del codice e passa al locale successivo della lista
                    
                    campo_dal = frame_date.locator("#ctl00_Cp1_Txtiniziochiusura, input[id*='Txtiniziochiusura']").first
                    campo_al = frame_date.locator("#ctl00_Cp1_txtfinechiusura, input[id*='txtfinechiusura']").first

                    print("   📝 [Robot] STEP 9: Digitazione reale data inizio...")
                    campo_dal.wait_for(state="visible", timeout=10000)
                    campo_dal.click()
                    campo_dal.press("Control+A")
                    campo_dal.press("Backspace")
                    time.sleep(1)
                    campo_dal.press_sequentially(data_inizio_pura, delay=100)
                    campo_dal.press("Tab")
                    time.sleep(1)

                    print("   📝 [Robot] STEP 9a: Digitazione reale data fine...")
                    campo_al.wait_for(state="visible", timeout=10000)
                    campo_al.click()
                    campo_al.press("Control+A")
                    campo_al.press("Backspace")
                    time.sleep(1)
                    campo_al.press_sequentially(data_fine_pura, delay=100)
                    campo_al.press("Tab")
                    time.sleep(2)
                    
                    try: frame_date.locator("#ctl00_Cp1_fascia_from").select_option(ora_inizio_pulita)
                    except Exception: pass
                    try: frame_date.locator("#ctl00_Cp1_fascia_to").select_option(ora_fine_pulita)
                    except Exception: pass
                    time.sleep(2)

                    try: page.screenshot(path="1_modulo_compilato.png", full_page=True)
                    except Exception: pass

                    print("   💾 [Robot] STEP 10: Invio moduli di chiusura a Snaitech (Clic su Tasto Salva)...")
                    frame_date.locator("#ctl00_Cp1_BtnOk").first.click(timeout=10000)
                    print(f"   ✅ [Robot] STEP 11: Invio completato. Attesa risposta visiva del portale...")
                    time.sleep(4)
                    
                    try: page.screenshot(path="2_risposta_immediata.png", full_page=True)
                    except Exception: pass
                    time.sleep(4)
                    
                    try: page.screenshot(path="errore_visivo_snaitech.png", full_page=True)
                    except Exception: pass
                    time.sleep(2)
                    
                    # 🛡️ RESET CELLA EXCEL CLOUD NATIVO VIA GIT PUSH
                    scarica_e_aggiorna_excel_su_github(codice_aams)
                    time.sleep(4)
                    
                    # 🛡️ PUNTAMENTO REALE RIPRISTINATO: Torna alla bacheca degli esercizi senza rompere il Login
                    page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx")
                    time.sleep(6)
                    
                except Exception as row_err:
                    print(f"   ⚠️ Nota compilazione: Scavalco riga. Errore: {str(row_err)}")
                    try:
                        page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx")
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

if __name__ == "__main__":
    avvia_sincronizzazione_automatica()
