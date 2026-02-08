import pandas as pd
import os

# Configurazione
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_PATH = os.path.join(BASE_DIR, "src", "CARICO_LAVORO_OFFICINA_Rev0.xlsm")
REPORT_PATH = os.path.join(BASE_DIR, "src", "REPORT_PRODUZIONE.txt")

def generate_dashboard_report():
    if not os.path.exists(EXCEL_PATH):
        print("File Excel non trovato.")
        return

    try:
        # Leggi Elenco Commesse
        df = pd.read_excel(EXCEL_PATH, sheet_name='ELENCO_COMMESSE')
        df = df.dropna(subset=['COMMESSA'])

        # 1. Statistiche Generali
        totale_commesse = len(df)
        completate = df['Fine lavorazione TOTALE'].notna().sum()
        in_corso = totale_commesse - completate

        # 2. Carico per Fase
        fasi_cols = [c for c in df.columns if str(c).startswith('ORE ')]
        carico_fasi = df[fasi_cols].sum()

        # 3. Prossime Scadenze
        df_scadenze = df[['COMMESSA', 'CLIENTE', 'DATA CONSEGNA RICHIESTA']].copy()
        df_scadenze = df_scadenze.sort_values(by='DATA CONSEGNA RICHIESTA').head(5)

        # Generazione Testo Report
        report = []
        report.append("="*40)
        report.append("   DASHBOARD DI RIEPILOGO PRODUZIONE")
        report.append("="*40)
        report.append(f"Data Report: {pd.Timestamp.now().strftime('%d/%m/%Y %H:%M')}")
        report.append(f"\\n[STATO COMMESSE]")
        report.append(f"- Totali: {totale_commesse}")
        report.append(f"- Completate: {completate}")
        report.append(f"- In corso: {in_corso}")

        report.append(f"\\n[CARICO ORE PER REPARTO]")
        for fase, ore in carico_fasi.items():
            report.append(f"- {str(fase).replace('ORE ', ''):<25}: {ore:>6} ore")

        report.append(f"\\n[PROSSIME 5 SCADENZE]")
        for _, row in df_scadenze.iterrows():
            data = row['DATA CONSEGNA RICHIESTA']
            data_str = data.strftime('%d/%m/%Y') if pd.notna(data) and hasattr(data, 'strftime') else "N/D"
            report.append(f"- {data_str} | {row['COMMESSA']} ({row['CLIENTE']})")

        report.append("\\n" + "="*40)

        # Scrittura su file e stampa a video
        output = "\\n".join(report)
        with open(REPORT_PATH, "w") as f:
            f.write(output)
        
        print(output)
        print(f"\\nReport salvato in: {REPORT_PATH}")

    except Exception as e:
        print(f"Errore durante la generazione della dashboard: {e}")

if __name__ == "__main__":
    generate_dashboard_report()