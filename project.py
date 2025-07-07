import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd
from bs4 import BeautifulSoup

def processar_arquivos(caminho_am, caminho_sp, pasta_saida):
    def processar_arquivo_xml(file_path, filial_codigo):
        with open(file_path, "r", encoding="utf-8") as file:
            soup = BeautifulSoup(file, "xml")
        rows = soup.find_all("Row")
        data = []
        current_vendedor = None
        current_nome = None
        for row in rows:
            cells = row.find_all("Cell")
            values = [cell.Data.text if cell.Data else "" for cell in cells]
            if len(values) == 2 and values[0].isdigit():
                current_vendedor = values[0]
                current_nome = values[1]
            elif len(values) >= 16 and values[0] != "Prefixo":
                data.append([filial_codigo, current_vendedor, current_nome] + values)
        return data

    # Processar arquivos
    dados_am = processar_arquivo_xml(caminho_am, "01")
    dados_sp = processar_arquivo_xml(caminho_sp, "03")
    dados_consolidados = dados_am + dados_sp

    # Criar DataFrame
    colunas = [
        "Filial", "Vendedor", "Nome", "Prefixo", "No. Titulo", "Parcela", "Cliente", "Nome Cliente",
        "Dt Comissao", "Vencto Origem", "Dt.Baixa", "Data Pagto", "Pedido",
        "Vlr Titulo", "Vlr Base", "%", "Comissao", "Tipo", "AJUSTE"
    ]
    df = pd.DataFrame(dados_consolidados, columns=colunas)

    # Formatando valores monetários
    col_to_format = ["Comissao", "Vlr Titulo", "Vlr Base"]
    for col in colunas:
        if col in col_to_format:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            df[col] = df[col].map(lambda x: f"{x:07.2f}".replace(".", ","))
    # Ajustar depois.

    # Formatando percentual
    df["%"] = df["%"].str.replace('%', '').str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    df["%"] = pd.to_numeric(df["%"], errors='coerce') / 100
    df["%"] = df["%"].map(lambda x: f"{x:.2%}".replace(".", ",") if pd.notnull(x) else "")

    # Formatando datas
    col_datas = ["Dt Comissao", "Vencto Origem", "Dt.Baixa", "Data Pagto"]
    for col in col_datas:
        df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')
        df[col] = df[col].dt.strftime('%d/%m/%Y').fillna("")

    # Garante que a pasta de saída existe
    os.makedirs(pasta_saida, exist_ok=True)

    # Salvar consolidado
    consolidado_path = os.path.join(pasta_saida, "tabela_padronizada.csv")
    df.to_csv(consolidado_path, index=False, sep=';', encoding='utf-8-sig')

    # Salvar arquivos por vendedor
    for (vendedor, nome), grupo in df.groupby(["Vendedor", "Nome"]):
        nome_arquivo = f"{vendedor}-{nome}.csv".replace("/", "-").replace("\\", "-")
        vendedor_path = os.path.join(pasta_saida, nome_arquivo)
        grupo.to_csv(vendedor_path, index=False, sep=';', encoding='utf-8-sig')

    messagebox.showinfo("Sucesso", f"Arquivos salvos em:\n{pasta_saida}")

def iniciar_interface():
    root = tk.Tk()
    root.title("Importar XML - AM/SP")
    root.geometry("600x400")

    def selecionar_am():
        caminho = filedialog.askopenfilename(title="Selecione o arquivo am.xml", filetypes=[("XML files", "*.xml")])
        entry_am.delete(0, tk.END)
        entry_am.insert(0, caminho)

    def selecionar_sp():
        caminho = filedialog.askopenfilename(title="Selecione o arquivo sp.xml", filetypes=[("XML files", "*.xml")])
        entry_sp.delete(0, tk.END)
        entry_sp.insert(0, caminho)

    def selecionar_saida():
        caminho = filedialog.askdirectory(title="Selecione a pasta de saída")
        entry_saida.delete(0, tk.END)
        entry_saida.insert(0, caminho)

    def executar():
        caminho_am = entry_am.get()
        caminho_sp = entry_sp.get()
        pasta_saida = entry_saida.get()

        if not os.path.isfile(caminho_am) or not os.path.isfile(caminho_sp):
            messagebox.showerror("Erro", "Selecione os dois arquivos XML corretamente.")
            return
        if not os.path.isdir(pasta_saida):
            messagebox.showerror("Erro", "Selecione uma pasta de saída válida.")
            return
        print('etapa confirada')
        processar_arquivos(caminho_am, caminho_sp, pasta_saida)

    # Interface visual
    tk.Label(root, text="Arquivo am.xml").pack(pady=5)
    entry_am = tk.Entry(root, width=80)
    entry_am.pack()
    tk.Button(root, text="Selecionar", command=selecionar_am).pack()

    tk.Label(root, text="Arquivo sp.xml").pack(pady=5)
    entry_sp = tk.Entry(root, width=80)
    entry_sp.pack()
    tk.Button(root, text="Selecionar", command=selecionar_sp).pack()

    tk.Label(root, text="Pasta de saída").pack(pady=5)
    entry_saida = tk.Entry(root, width=80)
    entry_saida.pack()
    tk.Button(root, text="Selecionar pasta", command=selecionar_saida).pack()

    tk.Button(root, text="Processar Arquivos", command=executar, bg="#4CAF50", fg="white", height=2).pack(pady=20)

    root.mainloop()

# Iniciar app
iniciar_interface()


