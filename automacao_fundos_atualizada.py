import pandas as pd
import customtkinter as ctk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

# Configurações da interface
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

caminho_csv = None
df_final_resultado = pd.DataFrame()
COLUNAS_OBRIGATORIAS = ['CNPJ_FUNDO', 'CNPJ_ADMIN', 'ADMIN', 'DT_INI_ADMIN', 'DT_FIM_ADMIN']

# ========== Funções Auxiliares ==========
def validar_csv(df):
    for col in COLUNAS_OBRIGATORIAS:
        if col not in df.columns:
            raise ValueError(f"Coluna obrigatória ausente: {col}")

def selecionar_arquivo_csv():
    global caminho_csv
    caminho = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if caminho:
        caminho_csv = caminho
        carregar_anos_do_csv()
        btn_executar.configure(state='normal')
        btn_exportar.configure(state='normal')

def carregar_anos_do_csv():
    try:
        df = pd.read_csv(caminho_csv, sep=';', dtype=str, encoding='latin1')
        validar_csv(df)
        df['DT_FIM_ADMIN'] = pd.to_datetime(df['DT_FIM_ADMIN'], dayfirst=True, errors='coerce')
        anos = sorted(df['DT_FIM_ADMIN'].dropna().dt.year.unique().tolist())
        anos = [str(ano) for ano in anos]
        anos.insert(0, "Todos")
        combo_ano.configure(values=anos)
        combo_ano.set("Todos")
    except Exception as e:
        combo_ano.configure(values=["Erro"])
        messagebox.showerror("Erro", f"Erro ao carregar o arquivo:\n{e}")

# ========== Automação ==========
def rodar_automacao():
    global df_final_resultado
    ano_selecionado = ano_var.get()

    if ano_selecionado != "Todos":
        try:
            int(ano_selecionado)
        except:
            messagebox.showerror("Erro", "Ano inválido!")
            return

    try:
        df = pd.read_csv(caminho_csv, sep=';', dtype=str, encoding='latin1')
        validar_csv(df)
        df['DT_INI_ADMIN'] = pd.to_datetime(df['DT_INI_ADMIN'], dayfirst=True, errors='coerce')
        df['DT_FIM_ADMIN'] = pd.to_datetime(df['DT_FIM_ADMIN'], dayfirst=True, errors='coerce')
        df = df.sort_values(by=['CNPJ_FUNDO', 'DT_INI_ADMIN'])

        df['CNPJ_ADMIN_NOVO'] = df.groupby('CNPJ_FUNDO')['CNPJ_ADMIN'].shift(-1)
        df['ADMIN_NOVO'] = df.groupby('CNPJ_FUNDO')['ADMIN'].shift(-1)
        df['DT_INI_ADMIN_NOVO'] = df.groupby('CNPJ_FUNDO')['DT_INI_ADMIN'].shift(-1)

        df_trocas = df[df['DT_FIM_ADMIN'].notna()].copy()

        df_resultado = df_trocas[[
            'CNPJ_FUNDO', 'CNPJ_ADMIN', 'ADMIN', 'DT_INI_ADMIN', 'DT_FIM_ADMIN',
            'CNPJ_ADMIN_NOVO', 'ADMIN_NOVO', 'DT_INI_ADMIN_NOVO'
        ]]

        df_resultado = df_resultado.sort_values(by=['CNPJ_FUNDO', 'DT_FIM_ADMIN'])
        df_resultado = df_resultado.groupby('CNPJ_FUNDO', as_index=False).tail(1)

        df_resultado = df_resultado[df_resultado['ADMIN_NOVO'].notna()]

        if ano_selecionado != "Todos":
            df_resultado = df_resultado[df_resultado['DT_FIM_ADMIN'].dt.year == int(ano_selecionado)]

        df_final_resultado = df_resultado.copy()

        for row in tabela.get_children():
            tabela.delete(row)

        if df_final_resultado.empty:
            messagebox.showinfo("Resultado", "Nenhuma troca de administrador encontrada.")
            estatisticas_var.set("Total de fundos com troca: 0")
        else:
            for _, row in df_final_resultado.iterrows():
                tabela.insert("", "end", values=[
                    row['CNPJ_FUNDO'], row['CNPJ_ADMIN'], row['ADMIN'],
                    row['CNPJ_ADMIN_NOVO'], row['ADMIN_NOVO']
                ])
            total_fundos = df_final_resultado['CNPJ_FUNDO'].nunique()
            estatisticas_var.set(f"Total de fundos com troca: {total_fundos}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")

def exportar_csv():
    if df_final_resultado.empty:
        messagebox.showinfo("Exportar", "Nada para exportar.")
        return
    caminho = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if caminho:
        df_final_resultado.to_csv(caminho, sep=';', index=False, encoding='utf-8-sig')
        messagebox.showinfo("Exportar", "Arquivo exportado com sucesso!")

def limpar_dados():
    global caminho_csv, df_final_resultado
    caminho_csv = None
    df_final_resultado = pd.DataFrame()
    for row in tabela.get_children():
        tabela.delete(row)
    combo_ano.set("Selecione")
    combo_ano.configure(values=["Selecione"])
    estatisticas_var.set("")
    btn_executar.configure(state='disabled')
    btn_exportar.configure(state='disabled')
    messagebox.showinfo("Limpar", "Interface resetada.")

# ========== Interface Gráfica ==========
janela = ctk.CTk()
janela.title("Verificação de Troca de Administrador")
janela.geometry("1300x700")

try:
    janela.iconbitmap("itau-logo.ico")
except:
    pass

frame_topo = ctk.CTkFrame(janela)
frame_topo.pack(pady=10)

btn_arquivo = ctk.CTkButton(
    frame_topo, text="Selecionar Arquivo CSV", command=selecionar_arquivo_csv,
    fg_color="#3498db", hover_color="#2980b9", text_color="white"
)
btn_arquivo.grid(row=0, column=0, padx=10)

ctk.CTkLabel(frame_topo, text="Filtrar por ano:").grid(row=0, column=1, padx=5)

ano_var = ctk.StringVar()
combo_ano = ctk.CTkComboBox(frame_topo, variable=ano_var, values=["Selecione"])
combo_ano.grid(row=0, column=2, padx=5)

btn_executar = ctk.CTkButton(
    frame_topo, text="Executar Automação", command=rodar_automacao, state='disabled',
    fg_color="#27ae60", hover_color="#1e8449", text_color="white"
)
btn_executar.grid(row=0, column=3, padx=10)

btn_exportar = ctk.CTkButton(
    frame_topo, text="Exportar Resultado", command=exportar_csv, state='disabled',
    fg_color="#f39c12", hover_color="#d68910", text_color="black"
)
btn_exportar.grid(row=0, column=4, padx=10)

btn_limpar = ctk.CTkButton(
    frame_topo, text="Limpar/Resetar", command=limpar_dados,
    fg_color="#e74c3c", hover_color="#c0392b", text_color="white"
)
btn_limpar.grid(row=0, column=5, padx=10)

estatisticas_var = ctk.StringVar()
ctk.CTkLabel(janela, textvariable=estatisticas_var, font=("Arial", 14)).pack(pady=5)

# ========== Tabela ==========
frame_tabela = ctk.CTkFrame(janela)
frame_tabela.pack(padx=10, pady=10, fill="both", expand=True)

colunas = ['CNPJ_FUNDO', 'CNPJ_ADMIN', 'ADMIN', 'CNPJ_ADMIN_NOVO', 'ADMIN_NOVO']

tabela = ttk.Treeview(frame_tabela, columns=colunas, show='headings')
for col in colunas:
    tabela.heading(col, text=col)
    tabela.column(col, width=200, anchor='center')

scroll_y = ttk.Scrollbar(frame_tabela, orient="vertical", command=tabela.yview)
scroll_x = ttk.Scrollbar(frame_tabela, orient="horizontal", command=tabela.xview)
tabela.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

tabela.grid(row=0, column=0, sticky='nsew')
scroll_y.grid(row=0, column=1, sticky='ns')
scroll_x.grid(row=1, column=0, sticky='ew')

frame_tabela.grid_rowconfigure(0, weight=1)
frame_tabela.grid_columnconfigure(0, weight=1)

janela.mainloop()