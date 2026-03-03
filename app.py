# app.py

import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import re
import threading
import sys
import ctypes

# Configuração do Tema Visual
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class ModernBalanceteApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        myappid = 'mycompany.myproduct.subproduct.version'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

        # configurações da janela
        self.title("Conversor de Balancetes - Fortes/Accountfy")
        self.geometry("800x650")

        # define o ícone da janela (título e barra de tarefas)
        try:
            self.iconbitmap(resource_path("icon.ico"))
        except:
            pass
        
        # Variáveis
        self.files_to_process = []
        self.output_folder = ""
        self.var_export_xlsx = ctk.BooleanVar(value=True)
        self.var_export_csv = ctk.BooleanVar(value=False)

        self._create_layout()

    def _create_layout(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # --- Cabeçalho ---
        self.header_frame = ctk.CTkFrame(self, corner_radius=0)
        self.header_frame.grid(row=0, column=0, sticky="ew")
        
        self.title_label = ctk.CTkLabel(self.header_frame, text="Automação Contábil - Formatador Excel", font=ctk.CTkFont(size=20, weight="bold"))
        self.title_label.pack(pady=15, padx=20, anchor="w")

        # --- Controles ---
        self.controls_frame = ctk.CTkFrame(self)
        self.controls_frame.grid(row=1, column=0, padx=20, pady=20, sticky="ew")
        self.controls_frame.grid_columnconfigure(1, weight=1)

        # 1. Input
        self.btn_files = ctk.CTkButton(self.controls_frame, text="Selecionar Arquivos", command=self.select_files, width=150)
        self.btn_files.grid(row=0, column=0, padx=10, pady=10)
        
        self.lbl_files = ctk.CTkEntry(self.controls_frame, placeholder_text="Nenhum arquivo selecionado")
        self.lbl_files.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.lbl_files.configure(state="disabled")

        # 2. Output
        self.btn_folder = ctk.CTkButton(self.controls_frame, text="Pasta de Saída", command=self.select_folder, width=150, fg_color="gray")
        self.btn_folder.grid(row=1, column=0, padx=10, pady=10)
        
        self.lbl_folder = ctk.CTkEntry(self.controls_frame, placeholder_text="Selecione a pasta de destino...")
        self.lbl_folder.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        self.lbl_folder.configure(state="disabled")

        # 3. opções
        self.chk_xlsx = ctk.CTkCheckBox(self.controls_frame, text="Exportar XLSX", variable=self.var_export_xlsx)
        self.chk_xlsx.grid(row=2, column=0, padx=10, pady=(10, 0), sticky="w")
        self.chk_csv = ctk.CTkCheckBox(self.controls_frame, text="Exportar CSV (Accountfy)", variable=self.var_export_csv)
        self.chk_csv.grid(row=2, column=1, padx=10, pady=(10, 0), sticky="w")

        # 4. Processar
        self.btn_process = ctk.CTkButton(self.controls_frame, text="PROCESSAR E FORMATAR", command=self.start_processing_thread, height=50, font=ctk.CTkFont(size=14, weight="bold"))
        self.btn_process.grid(row=3, column=0, columnspan=2, padx=10, pady=(20, 10), sticky="ew")
        
        # --- Log ---
        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="nsew")
        self.log_frame.grid_rowconfigure(0, weight=1)
        self.log_frame.grid_columnconfigure(0, weight=1)

        self.log_area = ctk.CTkTextbox(self.log_frame, font=ctk.CTkFont(family="Consolas", size=12))
        self.log_area.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.log_area.insert("0.0", "Aguardando arquivos...\n")
        self.log_area.configure(state="disabled")

    def log(self, message):
        self.log_area.configure(state="normal")
        self.log_area.insert("end", f">> {message}\n")
        self.log_area.see("end")
        self.log_area.configure(state="disabled")

    def select_files(self):
        filetypes = (('Excel/CSV', '*.xls *.xlsx *.csv'), ('Todos', '*.*'))
        files = filedialog.askopenfilenames(title='Selecione os Balancetes', filetypes=filetypes)
        if files:
            self.files_to_process = files
            self.lbl_files.configure(state="normal")
            self.lbl_files.delete(0, "end")
            self.lbl_files.insert(0, f"{len(files)} arquivos selecionados")
            self.lbl_files.configure(state="disabled")
            self.log(f"Selecionados: {len(files)} arquivos.")

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder = folder
            self.lbl_folder.configure(state="normal")
            self.lbl_folder.delete(0, "end")
            self.lbl_folder.insert(0, folder)
            self.lbl_folder.configure(state="disabled")
            self.log(f"Pasta de saída: {folder}")

    def start_processing_thread(self):
        threading.Thread(target=self.run_processing).start()

    def run_processing(self):
        if not self.files_to_process or not self.output_folder:
            messagebox.showwarning("Atenção", "Selecione arquivos e pasta de saída.")
            return

        self.log("--- INICIANDO ---")
        self.btn_process.configure(state="disabled")
        
        sucessos, erros = 0, 0

        for filepath in self.files_to_process:
            try:
                self.process_file(filepath)
                sucessos += 1
            except Exception as e:
                # LOGA ERRO NO CONSOLE E PREPARA PARA PERSISTÊNCIA SE NECESSÁRIO
                self.log(f"falha critica em {os.path.basename(filepath)}: {str(e)}")
                erros += 1

        self.btn_process.configure(state="normal")
        self.log(f"--- FIM: {sucessos} sucessos, {erros} erros ---")
        messagebox.showinfo("Concluído", f"processamento finalizado\nsucesso: {sucessos}\nerro: {erros}")

    def process_file(self, filepath):
        filename = os.path.basename(filepath)
        nome_base = os.path.splitext(filename)[0]
        self.log(f"lendo: {filename}")

        # 1. Carregar Arquivo
        try:
            if filepath.lower().endswith(('.xls', '.xlsx')):
                df_raw = pd.read_excel(filepath, header=None, dtype=str)
            else:
                try:
                    df_raw = pd.read_csv(filepath, header=None, sep=None, engine='python', encoding='latin1', dtype=str)
                except:
                    df_raw = pd.read_csv(filepath, header=None, sep=';', encoding='utf-8', dtype=str)
        except Exception as e:
            raise Exception(f"Falha ao abrir: {e}")

        # 2. Localizar Cabeçalho
        header_idx = -1
        col_indices = {}

        for i, row in df_raw.head(30).iterrows():
            row_vals = [str(v).strip().lower() for v in row.values]
            if 'conta' in row_vals and any('saldo' in v for v in row_vals):
                header_idx = i
                for idx, val in enumerate(row_vals):
                    if val == 'conta': col_indices['Conta'] = idx
                    elif 'descri' in val or 'hist' in val: col_indices['Descrição'] = idx
                    elif 'saldo atual' in val: col_indices['Saldo Atual'] = idx
                    elif 'saldo' in val and 'anterior' not in val and 'Saldo Atual' not in col_indices:
                         col_indices['Saldo Atual'] = idx
                break
        
        if header_idx == -1 or 'Conta' not in col_indices or 'Saldo Atual' not in col_indices:
            raise Exception("Layout não reconhecido.")

        # 3. Extrair Dados
        df_data = df_raw.iloc[header_idx+1:].copy()
        
        df_clean = pd.DataFrame()
        df_clean['Conta'] = df_data.iloc[:, col_indices['Conta']]
        df_clean['Descrição'] = df_data.iloc[:, col_indices.get('Descrição', 1)]
        
        saldo_col_idx = col_indices['Saldo Atual']
        df_clean['Saldo Atual'] = df_data.iloc[:, saldo_col_idx]

        # 4. Natureza
        natureza_col_idx = saldo_col_idx + 1
        if natureza_col_idx < df_data.shape[1]:
            df_clean['Natureza'] = df_data.iloc[:, natureza_col_idx]
        else:
            df_clean['Natureza'] = ""

        # Correção de Deslocamento
        amostra_saldo = df_clean['Saldo Atual'].dropna().head(10).astype(str).values
        if any(val.strip().upper() in ['D', 'C'] for val in amostra_saldo if len(val.strip()) == 1):
             real_natureza_idx = saldo_col_idx
             real_saldo_idx = saldo_col_idx - 1
             df_clean['Saldo Atual'] = df_data.iloc[:, real_saldo_idx]
             df_clean['Natureza'] = df_data.iloc[:, real_natureza_idx]
             self.log(" -> Ajustando colunas Saldo/Natureza.")

        # 5. Limpeza e Filtros
        df_clean = df_clean.dropna(subset=['Conta'])
        df_clean['Conta'] = df_clean['Conta'].astype(str).str.strip()
        
        # Filtro 1
        df_clean = df_clean[~df_clean['Conta'].str.contains(r'E\+', case=False, regex=True, na=False)]
        
        # Filtro 2
        df_clean = df_clean[df_clean['Conta'].str.match(r'^\d[\d\.]*')]

        # Limpeza de sufixo
        def limpar_conta(val):
            val = str(val).strip()
            if val.endswith('.0'):
                return val[:-2]
            return val
        df_clean['Conta'] = df_clean['Conta'].apply(limpar_conta)

        # 6. Formatação do Saldo
        def converter_para_br(val):
            val = str(val).strip()
            if not val or val.lower() == 'nan': return ""
            try:
                if '.' in val and ',' not in val:
                    numero = float(val)
                    return f"{numero:.2f}".replace('.', ',')
                elif ',' in val:
                     return val
                else: # Inteiro
                     return f"{float(val):.2f}".replace('.', ',')
            except:
                return val

        df_clean['Saldo Atual'] = df_clean['Saldo Atual'].apply(converter_para_br)
        df_clean['Natureza'] = df_clean['Natureza'].fillna('').astype(str).str.strip()

        # 7. exportação condicional
        if not self.var_export_xlsx.get() and not self.var_export_csv.get():
            self.log(" [aviso] nenhum formato de saída selecionado.")
            return

        # gera o excel se marcado
        if self.var_export_xlsx.get():
            out_path = os.path.join(self.output_folder, f"{nome_base}.xlsx")
            writer = pd.ExcelWriter(out_path, engine='xlsxwriter')
            df_clean.to_excel(writer, index=False, sheet_name='Plan1')

            workbook  = writer.book
            worksheet = writer.sheets['Plan1']
            fmt_texto_esq = workbook.add_format({'align': 'left'})
            fmt_texto_dir = workbook.add_format({'align': 'right'})
            
            worksheet.set_column('A:A', 20, fmt_texto_esq)
            worksheet.set_column('B:B', 60, fmt_texto_esq)
            worksheet.set_column('C:C', 20, fmt_texto_dir)
            worksheet.set_column('D:D', 10, fmt_texto_esq)

            writer.close()
            self.log(f" -> excel gerado: {nome_base}.xlsx")

        # gera o csv se marcado
        if self.var_export_csv.get():
            try:
                df_csv = pd.DataFrame()
                df_csv['CONTA_CONTABIL'] = df_clean['Conta']
                df_csv['NOME_CONTA'] = df_clean['Descrição']
                
                def formata_saldo_accountfy(row):
                    val = str(row['Saldo Atual']).replace('.', '').replace(',', '.')
                    try: val_float = float(val)
                    except: val_float = 0.0
                    nat = str(row['Natureza']).strip().upper()
                    if nat == 'C': val_float = -abs(val_float)
                    else: val_float = abs(val_float)
                    return f"{val_float:.2f}".replace('.', ',')
                
                df_csv['SALDO'] = df_clean.apply(formata_saldo_accountfy, axis=1)
                out_csv_path = os.path.join(self.output_folder, f"{nome_base}.csv")
                df_csv.to_csv(out_csv_path, index=False, sep=';', encoding='utf-8-sig')
                self.log(f" -> csv gerado: {nome_base}.csv")
            except Exception as err:
                self.log(f" [erro] falha no csv: {str(err)}")

if __name__ == "__main__":
    app = ModernBalanceteApp()
    app.mainloop()