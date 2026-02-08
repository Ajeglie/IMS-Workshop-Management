import shutil, os, from datetime import datetime
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SOURCE_FILE = os.path.join(BASE_DIR, "src", "CARICO_LAVORO_OFFICINA_Rev0.xlsm")
BACKUP_DIR = os.path.join(BASE_DIR, "backups")
def create_backup():
    if not os.path.exists(BACKUP_DIR): os.makedirs(BACKUP_DIR)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest_path = os.path.join(BACKUP_DIR, f"BACKUP_{timestamp}.xlsm")
    shutil.copy2(SOURCE_FILE, dest_path)
    print(f"Backup creato: {dest_path}")
if __name__ == '__main__': create_backup()
