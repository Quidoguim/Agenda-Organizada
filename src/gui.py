import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import pandas as pd
from processor import contar_atividades_por_pessoa_e_dia
from writer import atualizar_planilha_por_data

# Interface Gráfica
janela = tk.Tk()
janela.title("Organizador de Agenda")

# Variáveis globais
entrada_var = tk.StringVar()
saida_var = tk.StringVar()


def selecionar_arquivo_entrada():
    caminho = filedialog.askopenfilename(
        title="Selecione a planilha da agenda",
        filetypes=[("Planilhas Excel", "*.xlsx *.xls")]
    )
    if caminho:
        entrada_var.set(caminho)


def selecionar_arquivo_saida():
    caminho = filedialog.asksaveasfilename(
        title="Salvar planilha organizada como...",
        defaultextension=".xlsx",
        filetypes=[("Planilhas Excel", "*.xlsx")]
    )
    if caminho:
        saida_var.set(caminho)


def gerar_planilha():
    caminho_agenda = Path(entrada_var.get())
    caminho_saida = Path(saida_var.get())

    if not caminho_agenda.exists():
        messagebox.showerror("Erro", "Arquivo de agenda não encontrado.")
        return
    if not caminho_saida:
        messagebox.showerror(
            "Erro", "Defina um caminho para salvar a planilha.")
        return

    try:
        df = pd.read_excel(caminho_agenda)
        df.columns = df.columns.str.strip()
        dados_pessoas = contar_atividades_por_pessoa_e_dia(df)

        dados_formatados = []
        for nome, df_pessoa in dados_pessoas.items():
            for _, linha in df_pessoa.iterrows():
                registro = {
                    "Data": linha["Data"].strftime("%d/%m/%Y"),
                    "Nome": nome,
                    "Qtd Atividades": linha["Qtd Atividades"],
                }
                if "Qtd Atos" in linha:
                    registro["Qtd Atos"] = linha["Qtd Atos"]
                dados_formatados.append(registro)

        atualizar_planilha_por_data(dados_formatados, caminho_saida)
        messagebox.showinfo(
            "Sucesso", f"Planilha gerada com sucesso em:\n{caminho_saida}")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar a planilha:\n{e}")


# Entrada
tk.Label(janela, text="Arquivo da agenda (.xlsx):").pack(pady=5)
tk.Entry(janela, textvariable=entrada_var, width=50).pack()
tk.Button(janela, text="Selecionar arquivo de entrada",
          command=selecionar_arquivo_entrada).pack(pady=5)

# Saída
tk.Label(janela, text="Destino da planilha organizada:").pack(pady=5)
tk.Entry(janela, textvariable=saida_var, width=50).pack()
tk.Button(janela, text="Selecionar local de saída",
          command=selecionar_arquivo_saida).pack(pady=5)

# Botão principal
tk.Button(janela, text="Gerar Planilha Organizada",
          command=gerar_planilha, bg="#4CAF50", fg="white").pack(pady=10)

janela.mainloop()
