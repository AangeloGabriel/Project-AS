from Tratamento.TratamentoAS_IQ import tratamento
import time
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import shutil

ValidaBrasil = "AS Brasil Week"
ValidaIQMa = "- IQ Orders"
ValidaIQMi = "- IQ orders"

path_init = Path(r"C:\Users\angelo.alencar\Desktop\TESTES E TESTES\MonitoraPasta")
path_historico = Path(r"C:\Users\angelo.alencar\Desktop\TESTES E TESTES\HistoricoAs")

arquivos_processados = set()

class MyHandler(FileSystemEventHandler):
    def on_created(self, event):
        nome_evento = Path(event.src_path).name
        if nome_evento.startswith("~$"):
            return  

        time.sleep(0.5) 

        IQ = None
        Brasil = None

        for i in path_init.iterdir():
            nome = i.name
            if nome.startswith("~$") or not nome.endswith(".xlsx"):
                continue

            if ValidaBrasil in nome and (ValidaIQMa in nome or ValidaIQMi in nome):
                IQ = i
            elif ValidaBrasil in nome:
                Brasil = i

        if not Brasil or not IQ:
            return


        chave = frozenset([Brasil.name, IQ.name])
        if chave in arquivos_processados:
            return  

        try:
            a = tratamento(Brasil, IQ)
            a.save(r"C:\Users\angelo.alencar\Desktop\TESTES E TESTES\AS_Consolida\AS_Cru.xlsx")

            shutil.move(Brasil, path_historico / Brasil.name)
            shutil.move(IQ, path_historico / IQ.name)

            arquivos_processados.add(chave)

        except Exception as e:
            print(f"Erro no tratamento: {e}")

    def on_deleted(self, event):
        print(f"D-E-L-E-T-A-D-O") # referência nerd

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
