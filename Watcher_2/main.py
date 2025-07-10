from Tratamento.incrementa import incrementa
import time
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import shutil
import os 
from dotenv import load_dotenv

load_dotenv()



path_init = Path(os.getenv('caminho_cru'))
path_oficial = Path(os.getenv('caminho_consolidado'))
arquivos_processados = set()

print(path_init)
class MyHandler(FileSystemEventHandler):
    def on_created(self, event):
        ASCru = None
        for i in path_init.iterdir():
            name = i.name
            if name.startswith("~$") or not name.endswith(".xlsx"):
                continue

            if name == "As_Cru.xlsx":
                ASCru = i
        time.sleep(2.5)

        try:
            a = incrementa(ASCru, path_oficial)
            a.save(r"C:\Users\angelo.alencar\Desktop\TESTES E TESTES\Base_As\Base de AS 3.0.xlsx")

            os.remove(ASCru)
        except Exception as e:
            print(f"Erro no tratamento: {e}")

    def on_deleted(self, event):
        print(f"DELETADO")

def main():
    handler = MyHandler()
    observer = Observer()
    observer.schedule(handler, str(path_init), recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1.5)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    main()
