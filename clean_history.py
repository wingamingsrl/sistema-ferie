import os
from datetime import datetime
import pandas as pd

FILE_STORICO = "registro_ferie_salvato.xlsx"

def pulisci_storico_notturno():
    if os.path.exists(FILE_STORICO):
        try:
            df = pd.read_excel(FILE_STORICO).fillna("")
            if not df.empty:
                oggi = datetime.now().date()
                righe_valide = []
                for _, riga in df.iterrows():
                    try:
                        # Se la fine delle ferie è oggi o nel futuro, la teniamo
                        data_riap = datetime.strptime(str(riga["FINE_FERIE"]).strip(), "%d-%m-%Y").date()
                        if data_riap >= oggi:
                            righe_valide.append(riga)
                    except Exception:
                        righe_valide.append(riga)
                
                df_nuovo = pd.DataFrame(righe_valide)
                if df_nuovo.empty:
                    df_nuovo = pd.DataFrame(columns=["DATA_INSERIMENTO", "TECNICO", "LOCALE", "INIZIO_FERIE", "FINE_FERIE", "COPIA_PROMEMORIA"])
                
                df_nuovo.to_excel(FILE_STORICO, index=False)
                print("Pulizia notturna completata con successo.")
            else:
                print("Lo storico è già vuoto.")
        except Exception as e:
            print(f"Errore durante la pulizia: {str(e)}")
    else:
        print("File storico non trovato. Nessuna pulizia necessaria.")

if __name__ == "__main__":
    puli_storico_notturno()
