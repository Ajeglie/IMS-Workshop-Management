import pandas as pd, os
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_PATH = os.path.join(BASE_DIR, "src", "CARICO_LAVORO_OFFICINA_Rev0.xlsm")
def generate():
    df = pd.read_excel(EXCEL_PATH, sheet_name='ELENCO_COMMESSE').dropna(subset=['COMMESSA'])
    print("--- DASHBOARD ---")
    print(f"Commesse: {len(df)}")
    print(df[[c for c in df.columns if c.startswith('ORE ')]].sum())
if __name__ == '__main__': generate()
