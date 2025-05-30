import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pdfplumber
import re
import os
import subprocess
import sys
import logging

# Configurar logging para capturar avisos do pdfplumber (opcional, mas útil para depuração)
# Se não quiser ver os avisos de CropBox, pode comentar ou aumentar o nível de logging
# logging.basicConfig(level=logging.WARNING) # Mostra WARNINGS e acima
# logging.getLogger("pdfminer").setLevel(logging.WARNING) # Especificamente para pdfminer, usado por pdfplumber

# --- Funções de Extração e Processamento ---

def parse_value(value_str):
    """Converte uma string de valor monetário para float."""
    if not value_str or not isinstance(value_str, str):
        return 0.0
    cleaned_str = value_str.replace('.', '').replace(',', '.')
    try:
        return float(cleaned_str)
    except ValueError:
        return 0.0

def extract_uc_from_block(text_block):
    """Extrai o número da Unidade Consumidora (UC) do bloco de texto."""
    match = re.search(r"(?:UC:|Unidade Consumidora:)\s*(\d+)", text_block)
    if match:
        return match.group(1)
    return None

def extract_valor_total_fatura_from_block(text_block):
    """Extrai o valor total da fatura do bloco de texto."""
    match = re.search(r"Grupo / Subgrupo Tensão:.*?Valor:\s*R\$\s*([\d\.,]+)", text_block, re.DOTALL | re.IGNORECASE)
    if match:
        return parse_value(match.group(1))
    match_fallback = re.search(r"Valor:\s*R\$\s*([\d\.,]+)", text_block, re.IGNORECASE) # Fallback mais genérico
    if match_fallback:
        return parse_value(match_fallback.group(1))
    return 0.0

def extract_item_value_from_block(text_block, item_name_pattern):
    """Extrai o valor de um item específico do bloco de texto da fatura."""
    # Padrão: NOME_ITEM (pode ser multilinhas se o nome for longo) seguido por 3 colunas numéricas
    # O valor (R$) é a terceira coluna numérica.
    # Ex: Consumo TE       2.645      0,38190      1.010,13
    #     Tributo Retido IRPJ    0      0,00000     -24,78
    # A regex procura o nome do item seguido por três grupos numéricos.
    # (?s) permite que . corresponda a novas linhas, útil se o nome do item for quebrado.
    # O \s+ entre os grupos numéricos é importante.
    # Usamos ^ para ancorar no início da linha para maior precisão.
    full_regex_pattern = rf"^(?:{item_name_pattern})\s+([\d\.,-]+)\s+([\d\.,-]+)\s+([\d\.,-]+)"

    match = re.search(full_regex_pattern, text_block, re.MULTILINE | re.IGNORECASE)
    if match:
        return parse_value(match.group(3)) # O valor do item é o terceiro grupo numérico
    return 0.0

def extract_fatura_data_from_text_block(text_block, df_base, pdf_filename_for_error_logging):
    """
    Extrai todos os dados de uma fatura a partir de um bloco de texto.
    Retorna um dicionário com os dados ou um dicionário de erro.
    """
    uc_number = extract_uc_from_block(text_block)
    if not uc_number:
        # Se não há UC, este bloco provavelmente não é uma fatura ou está mal formatado.
        return None # Sinaliza que este bloco não gerou dados válidos

    base_info = df_base[df_base['UC'].astype(str) == uc_number]
    if base_info.empty:
        return {"error": f"UC {uc_number} (de {pdf_filename_for_error_logging}) não encontrada na planilha base."}

    cod_reg = base_info['Cod de Reg'].iloc[0]
    nome_base = base_info['Nome'].iloc[0]

    valor_total_fatura = extract_valor_total_fatura_from_block(text_block)
    # if valor_total_fatura == 0.0:
    #     print(f"AVISO: Valor total da fatura não encontrado ou zerado para UC {uc_number} em {pdf_filename_for_error_logging}")

    positive_items_patterns = {
        "Consumo TE": r"Consumo TE",
        "Consumo TUSD": r"Consumo TUSD",
        "COSIP Municipal": r"COSIP Municipal.*?(?=\s+[\d\.,-]+\s+[\d\.,-]+\s+[\d\.,-]+)"
    }
    negative_items_patterns = {
        "Tributo Retido IRPJ": r"Tributo Retido IRPJ",
        "Tributo Retido PIS": r"Tributo Retido PIS",
        "Tributo Retido COFINS": r"Tributo Retido COFINS",
        "Tributo Retido CSLL": r"Tributo Retido CSLL"
    }

    soma_positivos = 0.0
    for item_desc, pattern_str in positive_items_patterns.items():
        val = extract_item_value_from_block(text_block, pattern_str)
        soma_positivos += val

    soma_negativos = 0.0
    for item_desc, pattern_str in negative_items_patterns.items():
        val = extract_item_value_from_block(text_block, pattern_str)
        soma_negativos += val # Valores já vêm com sinal negativo

    valor_bruto_calculado = soma_positivos
    valor_liquido_calculado = valor_bruto_calculado + soma_negativos

    return {
        "UC": uc_number,
        "Cod de Reg": cod_reg,
        "Nome": nome_base, # Usa o nome da planilha base
        "Valor Total Fatura (R$)": valor_total_fatura,
        "Valor Bruto Calculado (R$)": valor_bruto_calculado,
        "Valor Líquido Calculado (R$)": valor_liquido_calculado,
        "pdf_filename": pdf_filename_for_error_logging # Para referência
    }

def process_pdf_file(pdf_path, df_base):
    """
    Processa um único arquivo PDF, que pode conter múltiplos registros por página.
    Retorna uma lista de dicionários (dados da fatura ou erros).
    """
    results_for_this_pdf = []
    pdf_filename = os.path.basename(pdf_path)

    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                results_for_this_pdf.append({"error": f"PDF sem páginas: {pdf_filename}"})
                return results_for_this_pdf

            for page_num, page in enumerate(pdf.pages):
                page_text = page.extract_text()
                if not page_text or not page_text.strip():
                    # print(f"Página {page_num + 1} de {pdf_filename} não contém texto extraível.")
                    continue

                # Encontrar todas as ocorrências de "UC: XXXXX" ou "Unidade Consumidora: XXXXX" na página
                uc_pattern = r"(?:UC:|Unidade Consumidora:)\s*\d+"
                matches = list(re.finditer(uc_pattern, page_text))

                if not matches:
                    # Se nenhuma UC for encontrada na página, tentar processar a página inteira como um único bloco
                    # Isso é um fallback, pode não ser ideal se houver múltiplos registros sem o padrão UC claro
                    # print(f"Nenhuma UC explícita na página {page_num+1} de {pdf_filename}. Tentando processar a página inteira.")
                    fatura_data = extract_fatura_data_from_text_block(page_text, df_base, pdf_filename)
                    if fatura_data: # Pode ser None se o bloco não for válido, ou um dict de erro
                        results_for_this_pdf.append(fatura_data)
                    continue

                for i, match in enumerate(matches):
                    start_block = match.start() # Início da string "UC: XXXX" atual

                    # O final do bloco é o início da PRÓXIMA UC na mesma página, ou o final do texto da página
                    if i + 1 < len(matches):
                        end_block = matches[i+1].start()
                    else:
                        end_block = len(page_text)
                    
                    current_text_block = page_text[start_block:end_block]
                    # print(f"--- Processando bloco da UC (Pág {page_num+1}, Bloco {i+1}) em {pdf_filename} ---")
                    # print(current_text_block[:300] + "...") # Print início do bloco para debug

                    fatura_data = extract_fatura_data_from_text_block(current_text_block, df_base, pdf_filename)
                    if fatura_data: # Pode ser None ou um dict de erro
                        results_for_this_pdf.append(fatura_data)
            
            if not results_for_this_pdf:
                 results_for_this_pdf.append({"error": f"Nenhum dado de fatura encontrado em {pdf_filename} após processar todas as páginas."})


    except Exception as e:
        # Captura exceções ao abrir/ler o PDF ou erros inesperados no loop
        results_for_this_pdf.append({"error": f"Erro crítico ao processar {pdf_filename}: {e}"})
    
    return results_for_this_pdf


# --- Funções da Interface Gráfica (maioria inalterada, exceto start_processing) ---
class AppCelescReporter:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Relatório Celesc")
        self.center_window(650, 450) # Aumentei um pouco para melhor visualização

        self.base_sheet_path = os.path.join(os.path.dirname(sys.argv[0]), "base", "ucs.sub.xlsx")
        self.df_base = None
        self.pdf_files = []
        self.output_dir = os.path.join(os.path.expanduser("~"), "Desktop") # Padrão para Desktop

        # --- Estilo ---
        style = ttk.Style(self.root)
        style.theme_use('clam') 

        # --- Frames ---
        main_frame = ttk.Frame(self.root, padding="10 10 10 10")
        main_frame.pack(expand=True, fill=tk.BOTH)

        # --- Widgets ---
        base_frame = ttk.LabelFrame(main_frame, text="Planilha Base de UCs", padding="10")
        base_frame.pack(fill=tk.X, pady=5)
        self.base_path_label = ttk.Label(base_frame, text=f"Caminho: {self.base_sheet_path}", wraplength=600)
        self.base_path_label.pack(fill=tk.X)
        self.base_status_label = ttk.Label(base_frame, text="Status: Não carregada")
        self.base_status_label.pack(fill=tk.X)
        self.load_base_sheet() 

        pdf_frame = ttk.LabelFrame(main_frame, text="Arquivos PDF das Faturas", padding="10")
        pdf_frame.pack(fill=tk.X, pady=5)
        self.pdf_button = ttk.Button(pdf_frame, text="Selecionar PDFs da Celesc", command=self.select_pdfs)
        self.pdf_button.pack(side=tk.LEFT, padx=(0,10))
        self.pdf_label = ttk.Label(pdf_frame, text="Nenhum PDF selecionado")
        self.pdf_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        output_frame = ttk.LabelFrame(main_frame, text="Pasta de Saída do Relatório", padding="10")
        output_frame.pack(fill=tk.X, pady=5)
        self.output_dir_button = ttk.Button(output_frame, text="Definir Pasta de Saída", command=self.select_output_dir)
        self.output_dir_button.pack(side=tk.LEFT, padx=(0,10))
        self.output_label = ttk.Label(output_frame, text=f"Padrão: {self.output_dir}")
        self.output_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        action_frame = ttk.Frame(main_frame, padding="10")
        action_frame.pack(fill=tk.X, pady=10)
        
        self.progress_bar = ttk.Progressbar(action_frame, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.pack(pady=5, fill=tk.X)

        self.process_button = ttk.Button(action_frame, text="Iniciar Processamento de Relatório", command=self.start_processing)
        self.process_button.pack(pady=5)
        
        self.status_label = ttk.Label(action_frame, text="Aguardando configuração...")
        self.status_label.pack(fill=tk.X, pady=5)

    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)
        self.root.geometry(f'{width}x{height}+{int(x)}+{int(y)}')

    def load_base_sheet(self):
        try:
            if not os.path.exists(self.base_sheet_path):
                self.base_status_label.config(text=f"Status: ERRO - Arquivo não encontrado em {self.base_sheet_path}", foreground="red")
                self.df_base = None
                return
            
            self.df_base = pd.read_excel(self.base_sheet_path, dtype={'UC': str, 'Cod de Reg': str, 'Nome': str})
            required_cols = ['UC', 'Cod de Reg', 'Nome']
            if not all(col in self.df_base.columns for col in required_cols):
                missing_cols = [col for col in required_cols if col not in self.df_base.columns]
                self.base_status_label.config(text=f"Status: ERRO - Colunas faltando: {', '.join(missing_cols)}", foreground="red")
                self.df_base = None
                return

            self.df_base.dropna(subset=['UC'], inplace=True)
            self.df_base['UC'] = self.df_base['UC'].astype(str).str.strip() # Garante string e remove espaços
            
            num_ucs = len(self.df_base)
            if num_ucs == 0:
                self.base_status_label.config(text="Status: Planilha base carregada, mas sem UCs válidas.", foreground="orange")
            else:
                self.base_status_label.config(text=f"Status: Planilha base carregada. {num_ucs} UCs encontradas.", foreground="green")
        except Exception as e:
            self.base_status_label.config(text=f"Status: ERRO ao carregar planilha base - {e}", foreground="red")
            self.df_base = None

    def select_pdfs(self):
        files = filedialog.askopenfilenames(
            title="Selecione os arquivos PDF da Celesc",
            filetypes=(("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*"))
        )
        if files:
            self.pdf_files = list(files)
            self.pdf_label.config(text=f"{len(self.pdf_files)} PDF(s) selecionado(s)")
        else:
            self.pdf_label.config(text="Nenhum PDF selecionado")
            self.pdf_files = []


    def select_output_dir(self):
        directory = filedialog.askdirectory(title="Selecione a pasta para salvar o relatório")
        if directory:
            self.output_dir = directory
            self.output_label.config(text=self.output_dir)
        # else:
            # Mantém o diretório padrão ou o último selecionado

    def start_processing(self):
        self.load_base_sheet() 

        if self.df_base is None or self.df_base.empty:
            messagebox.showerror("Erro de Configuração", "Planilha base de UCs não carregada, inválida ou vazia. Verifique o arquivo 'base/ucs.sub.xlsx'.")
            return
        if not self.pdf_files:
            messagebox.showerror("Erro de Configuração", "Nenhum arquivo PDF foi selecionado para processamento.")
            return
        if not self.output_dir or not os.path.isdir(self.output_dir):
            messagebox.showerror("Erro de Configuração", "Pasta de saída inválida ou não definida.")
            return

        self.status_label.config(text="Processando... Por favor, aguarde.")
        self.process_button.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = len(self.pdf_files)
        self.root.update_idletasks()

        all_processed_data = []
        all_errors = []

        for i, pdf_path in enumerate(self.pdf_files):
            self.status_label.config(text=f"Processando PDF {i+1}/{len(self.pdf_files)}: {os.path.basename(pdf_path)}")
            self.progress_bar["value"] = i + 1
            self.root.update_idletasks()
            
            # process_pdf_file retorna uma LISTA de resultados/erros para o PDF atual
            results_from_pdf = process_pdf_file(pdf_path, self.df_base) 
            
            for item in results_from_pdf:
                if isinstance(item, dict) and "error" in item:
                    all_errors.append(item["error"])
                elif isinstance(item, dict): # Sucesso na extração do bloco
                    all_processed_data.append(item)
                # Se item for None (de extract_fatura_data_from_text_block), ignoramos silenciosamente
        
        self.process_button.config(state=tk.NORMAL)

        if not all_processed_data:
            error_message = "Nenhum dado de fatura válido foi extraído."
            if all_errors:
                error_message += "\n\nErros encontrados:\n- " + "\n- ".join(list(set(all_errors))[:10]) # Mostra até 10 erros únicos
            messagebox.showwarning("Processamento Concluído", error_message)
            self.status_label.config(text="Concluído. Nenhum dado válido encontrado.")
            return

        df_report = pd.DataFrame(all_processed_data)
        
        final_columns = [
            "UC", "Cod de Reg", "Nome", 
            "Valor Total Fatura (R$)", "Valor Bruto Calculado (R$)", "Valor Líquido Calculado (R$)",
            "pdf_filename"
        ]
        for col in final_columns: # Garantir que todas as colunas existam
            if col not in df_report.columns:
                df_report[col] = None 
        df_report = df_report[final_columns] # Reordenar

        # Formatando colunas numéricas
        for col in ["Valor Total Fatura (R$)", "Valor Bruto Calculado (R$)", "Valor Líquido Calculado (R$)"]:
            if col in df_report.columns:
                df_report[col] = df_report[col].apply(lambda x: f"{x:.2f}" if pd.notnull(x) else None)


        output_file_path = os.path.join(self.output_dir, "Relatorio_Celesc.xlsx")

        try:
            df_report.to_excel(output_file_path, index=False)
            
            summary_message = f"Processamento concluído!\n{len(all_processed_data)} registros extraídos.\nRelatório salvo em:\n{output_file_path}"
            log_file_path = ""

            if all_errors:
                unique_errors = list(set(all_errors)) # Para não repetir muitos erros iguais
                summary_message += f"\n\nForam encontrados {len(all_errors)} problemas (com {len(unique_errors)} tipos de erro distintos)."
                
                try:
                    log_file_path = os.path.join(self.output_dir, "log_erros_celesc.txt")
                    with open(log_file_path, "w", encoding="utf-8") as f_log:
                        f_log.write("--- Log de Erros do Processamento Celesc ---\n\n")
                        for err_idx, err in enumerate(all_errors):
                            f_log.write(f"{err_idx+1}. {err}\n")
                    summary_message += f"\nUm log detalhado dos erros foi salvo em:\n{log_file_path}"
                except Exception as log_e:
                    summary_message += f"\n(Não foi possível salvar o log de erros: {log_e})"
                
                print("\n--- Erros Detalhados ---")
                for err in unique_errors: # Printar os únicos no console
                    print(err)
                print(f"--- Fim dos Erros (total: {len(all_errors)}, únicos: {len(unique_errors)}) ---")
                messagebox.showwarning("Processamento Concluído com Alertas", summary_message)
            else:
                messagebox.showinfo("Processamento Concluído", summary_message)
            
            self.status_label.config(text="Concluído. Relatório gerado.")
            
            if os.path.exists(output_file_path):
                if sys.platform == "win32":
                    os.startfile(output_file_path)
                elif sys.platform == "darwin":
                    subprocess.call(("open", output_file_path))
                else:
                    subprocess.call(("xdg-open", output_file_path))
            if log_file_path and os.path.exists(log_file_path): # Abrir log se foi gerado
                 if sys.platform == "win32":
                    os.startfile(log_file_path)


        except Exception as e:
            messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o relatório Excel: {e}")
            self.status_label.config(text="Erro ao salvar relatório.")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppCelescReporter(root)
    root.mainloop()