import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import pdfplumber
import re
import os
import subprocess
import sys
import logging
import threading  # Importa a biblioteca threading

# Tentar importar openpyxl e seus componentes necessários
try:
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Alignment
except ImportError:
    messagebox.showerror("Dependência Faltando",
                         "A biblioteca 'openpyxl' é necessária para formatação avançada do Excel. "
                         "Por favor, instale-a com 'pip install openpyxl' e tente novamente.")
    sys.exit(1)


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
    """Extrai o valor total da fatura (que será o Valor Líquido) do bloco de texto."""
    # Tenta o padrão mais específico primeiro (Valor no resumo da fatura individual)
    # Busca pela linha que começa com "Valor:" ou "Valor:" seguida por R$ e o valor
    # re.DOTALL | re.IGNORECASE permite que . casem com quebras de linha e ignora maiúsculas/minúsculas
    match = re.search(r"Valor:\s*R\$\s*([\d\.,]+)", text_block, re.DOTALL | re.IGNORECASE)
    if match:
        return parse_value(match.group(1))

    # Fallback para o padrão "TOTAL A PAGAR" (comum na primeira página ou rodapé)
    match_fallback_total = re.search(r"TOTAL A PAGAR\s*R\$\s*([\d\.,]+)", text_block, re.IGNORECASE)
    if match_fallback_total:
        return parse_value(match_fallback_total.group(1))

    return 0.0

def extract_item_value_from_block(text_block, item_name_pattern):
    """
    Extrai o valor da coluna 'Valor (R$)' para um item específico da seção 'Itens da Fatura'.
    Modificado para pegar o valor na 3ª coluna numérica após o nome do item,
    para lidar com o layout específico dos 'Tributos Retidos'.
    """
    if not text_block or not isinstance(text_block, str):
        return 0.0

    # Normaliza múltiplos espaços para um único espaço e remove espaços no início/fim das linhas
    cleaned_text_block = "\n".join(line.strip() for line in text_block.splitlines() if line.strip())
    cleaned_text_block = re.sub(r'[ \t]+', ' ', cleaned_text_block)


    # --- Nova Regex para itens de Tributos Retidos com colunas extras ---
    # Procura pelo item, pula 2 sequências de valores numéricos, e captura a 3ª.
    # Ex: Nome do Item | Quantidade | Preço Unitário | Valor (R$) | Outras Colunas
    # Re.escape() para tratar caracteres especiais no nome do item.
    # .*? para casar qualquer coisa não gananciosa (como espaços)
    # \s+[\d\.,]+ para casar um ou mais espaços e uma sequência numérica (ponto/vírgula inclusos)
    # Repetimos \s+[\d\.,]+ para pular as duas primeiras colunas numéricas.
    # (-?[\d\.,]+) para capturar o valor que pode ser negativo.
    pattern = rf"{item_name_pattern}.*?\s+[\d\.,]+.*?\s+[\d\.,]+.*?\s+(-?[\d\.,]+)"


    match = re.search(pattern, cleaned_text_block, re.MULTILINE | re.IGNORECASE | re.DOTALL) # Adicionado re.DOTALL para que o '.' case com quebras de linha


    if match:
        return parse_value(match.group(1))

    return 0.0


def extract_fatura_data_from_text_block(text_block, df_base, pdf_filename_for_error_logging, logger_func, page_num=None): # Adicionado page_num
    """
    Extrai todos os dados de uma fatura a partir de um bloco de texto.
    Retorna um dicionário com os dados ou um dicionário de erro.
    """
    uc_number = extract_uc_from_block(text_block)
    if not uc_number:
        return None

    base_info = df_base[df_base['UC'].astype(str) == uc_number]
    if base_info.empty:
        error_msg = f"UC {uc_number} (de {pdf_filename_for_error_logging}) não encontrada na planilha base."
        if logger_func:
            logger_func(error_msg, "ERROR")
        return {"error": error_msg, "UC": uc_number, "Numero da Pagina": pdf_filename_for_error_logging} # Chave alterada para erro

    cod_reg = base_info['Cod de Reg'].iloc[0]
    nome_base = base_info['Nome'].iloc[0]

    valor_liquido_fatura = extract_valor_total_fatura_from_block(text_block)
    if valor_liquido_fatura == 0.0 and logger_func:
         logger_func(f"AVISO: Valor Líquido da fatura (Valor Total da Fatura) não encontrado ou zerado para UC {uc_number} em {pdf_filename_for_error_logging}. Verifique o PDF ou o padrão de extração.", "WARNING")

    # --- Extração de Tributos Retidos (RETENÇÃO) ---
    tributos_retidos_patterns = {
        "IRPJ": r"Tributo Retido IRPJ",
        "PIS": r"Tributo Retido PIS",
        "COFINS": r"Tributo Retido COFINS",
        "CSLL": r"Tributo Retido CSLL"
    }

    soma_valores_negativos_tributos = 0.0
    found_any_tax_value_non_zero = False

    for nome_tributo, pattern_str in tributos_retidos_patterns.items():
        # Usamos re.escape(pattern_str) aqui para lidar com caracteres especiais nos nomes dos tributos
        valor_tributo = extract_item_value_from_block(text_block, re.escape(pattern_str))
        soma_valores_negativos_tributos += valor_tributo
        if valor_tributo != 0.0:
            found_any_tax_value_non_zero = True

    retencao_tributos = abs(soma_valores_negativos_tributos) # Nome da variável atualizado para RETENÇÃO
    
    if retencao_tributos == 0.0 and not found_any_tax_value_non_zero and logger_func:
         logger_func(f"INFO: Nenhum item de tributo retido ('Tributo Retido IRPJ/PIS/COFINS/CSLL') encontrado ou extraído com valor não zero para UC {uc_number} em {pdf_filename_for_error_logging}. 'RETENÇÃO (R$)' será 0.00.", "INFO")

    # --- Extração do COSIP ---
    # O padrão para COSIP é ligeiramente diferente, pois o município varia.
    # Vamos usar um padrão mais genérico que casa "COSIP Municipal" seguido de qualquer coisa
    # até o valor. Reutilizamos extract_item_value_from_block.
    # Usamos re.escape para "COSIP Municipal", mas .*? para o nome do município e o valor.
    # A regex é ajustada para procurar "COSIP Municipal", seguido por qualquer caractere
    # (não ganancioso), até o valor numérico.
    # A função `extract_item_value_from_block` já espera um padrão para o nome do item e busca o valor.
    # Então, para o COSIP, o "item_name_pattern" será "COSIP Municipal" e a função interna dela
    # buscará o valor.
    cosip_item_name_pattern = r"COSIP Municipal" # O padrão é só o nome do item
    valor_cosip = extract_item_value_from_block(text_block, cosip_item_name_pattern)

    if valor_cosip == 0.0 and logger_func:
        logger_func(f"INFO: COSIP (ou 'COSIP Municipal') não encontrado ou extraído com valor zero para UC {uc_number} em {pdf_filename_for_error_logging}. 'COSIP (R$)' será 0.00.", "INFO")

    valor_bruto_fatura_calculado = valor_liquido_fatura + retencao_tributos # Valor bruto antes do COSIP
    
    # Novo cálculo: Energia = Valor Bruto - COSIP
    valor_energia_calculado = valor_bruto_fatura_calculado - valor_cosip

    # Constrói o campo 'Numero da Pagina'
    numero_pagina_display = f"{pdf_filename_for_error_logging} (Pág. {page_num + 1})" if page_num is not None else pdf_filename_for_error_logging


    return {
        "UC": uc_number,
        "Cod de Reg": cod_reg,
        "Nome": nome_base,
        "ENERGIA (R$)": valor_energia_calculado, # Nova coluna
        "COSIP (R$)": valor_cosip, # Nova coluna
        "Valor Bruto (R$)": valor_bruto_fatura_calculado,
        "RETENÇÃO (R$)": retencao_tributos, # Nome da coluna renomeado
        "LÍQUIDO (R$)": valor_liquido_fatura, # Nome da coluna renomeado
        "Numero da Pagina": numero_pagina_display # Chave alterada
    }

def process_pdf_file(pdf_path, df_base, logger_func):
    """
    Processa um único arquivo PDF.
    Retorna uma lista de dicionários (dados da fatura ou erros).
    """
    results_for_this_pdf = []
    pdf_filename = os.path.basename(pdf_path)

    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                error_msg = f"PDF sem páginas: {pdf_filename}"
                logger_func(error_msg, "ERROR")
                results_for_this_pdf.append({"error": error_msg, "Numero da Pagina": pdf_filename}) # Chave alterada para erro
                return results_for_this_pdf

            for page_num, page in enumerate(pdf.pages): # page_num está disponível aqui
                page_text = page.extract_text(x_tolerance=2, y_tolerance=3)
                if not page_text or not page_text.strip():
                    logger_func(f"Página {page_num + 1} de {pdf_filename} não contém texto extraível.", "INFO")
                    continue

                uc_pattern = r"(?:UC:|Unidade Consumidora:)\s*\d+"
                matches = list(re.finditer(uc_pattern, page_text))

                if not matches:
                    if page_num == 0:
                         logger_func(f"Nenhuma UC explícita na página {page_num+1} (provável sumário) de {pdf_filename}. Pulando página.", "INFO")
                         continue
                    else:
                        logger_func(f"Nenhuma UC explícita na página {page_num+1} de {pdf_filename}. Tentando processar a página inteira como um bloco único (pode ser uma fatura de página inteira ou página sem dados).", "INFO")
                        # Passa page_num para a função de extração
                        fatura_data = extract_fatura_data_from_text_block(page_text, df_base, pdf_filename, logger_func, page_num=page_num)
                        if fatura_data:
                            results_for_this_pdf.append(fatura_data)
                        continue

                for i, match in enumerate(matches):
                    start_block = match.start()
                    end_block = matches[i+1].start() if i + 1 < len(matches) else len(page_text)
                    current_text_block = page_text[start_block:end_block]

                    # Passa page_num para a função de extração
                    fatura_data = extract_fatura_data_from_text_block(current_text_block, df_base, pdf_filename, logger_func, page_num=page_num)
                    if fatura_data:
                        results_for_this_pdf.append(fatura_data)

            if not results_for_this_pdf:
                 no_data_msg = f"Nenhum dado de fatura (com UC identificável) ou erro relevante encontrado em {pdf_filename} após processar todas as páginas com texto extraível."
                 logger_func(no_data_msg, "WARNING")

    except Exception as e:
        critical_error_msg = f"Erro crítico ao processar {pdf_filename}: {e}"
        logger_func(critical_error_msg, "CRITICAL_ERROR")
        results_for_this_pdf.append({"error": critical_error_msg, "Numero da Pagina": pdf_filename, "UC": "N/A"}) # Chave alterada para erro

    return results_for_this_pdf


# --- Classe da Interface Gráfica ---
class AppCelescReporter:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("Gerador de Relatório Celesc")
        self.center_window(700, 650)

        if getattr(sys, 'frozen', False):
             basedir = os.path.dirname(sys.executable)
        else:
             basedir = os.path.dirname(__file__)

        self.base_sheet_path = os.path.join(basedir, "base", "ucs.sub.xlsx")

        self.df_base = None
        self.pdf_files = []
        self.output_dir = os.path.join(os.path.expanduser("~"), "Desktop")

        style = ttk.Style(self.root)
        style.theme_use('clam')

        main_frame = ttk.Frame(self.root, padding="10 10 10 10")
        main_frame.pack(expand=True, fill=tk.BOTH)

        base_frame = ttk.LabelFrame(main_frame, text="Planilha Base de UCs", padding="10")
        base_frame.pack(fill=tk.X, pady=5)
        self.base_path_label = ttk.Label(base_frame, text=f"Caminho: {self.base_sheet_path}", wraplength=650)
        self.base_path_label.pack(fill=tk.X)
        self.base_status_label = ttk.Label(base_frame, text="Status: Não carregada")
        self.base_status_label.pack(fill=tk.X)
        
        # --- Seção Log em Tempo Real (Inicializada mais cedo para evitar AttributeError) ---
        log_frame = ttk.LabelFrame(main_frame, text="Log de Processamento", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5) # Empacotado antes de ser referenciado
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.tag_config("INFO", foreground="black")
        self.log_text.tag_config("WARNING", foreground="orange")
        self.log_text.tag_config("ERROR", foreground="red")
        self.log_text.tag_config("CRITICAL_ERROR", foreground="red", font=('TkDefaultFont', 9, 'bold'))
        self.log_text.tag_config("SUCCESS", foreground="green")
        self.log_text.tag_config("DEBUG", foreground="gray")

        self.load_base_sheet() # Agora pode ser chamado sem erro, pois self.log_text existe
        
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

    def log_message(self, message, level="INFO"):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{level}] {message}\n", level.upper())
        self.log_text.config(state=tk.DISABLED)
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)
        self.root.geometry(f'{width}x{height}+{int(x)}+{int(y)}')

    def load_base_sheet(self):
        """Carrega a planilha base de UCs e atualiza o status na GUI."""
        self.log_message("Tentando carregar planilha base...", "INFO")
        try:
            if not os.path.exists(self.base_sheet_path):
                msg = f"Status: ERRO - Arquivo base não encontrado em {self.base_sheet_path}"
                self.base_status_label.config(text=msg, foreground="red")
                self.log_message(msg, "ERROR")
                self.df_base = None
                return

            self.df_base = pd.read_excel(self.base_sheet_path, engine='openpyxl', dtype={'UC': str, 'Cod de Reg': str, 'Nome': str})
            required_cols = ['UC', 'Cod de Reg', 'Nome']
            if not all(col in self.df_base.columns for col in required_cols):
                missing_cols = [col for col in required_cols if col not in self.df_base.columns]
                msg = f"Status: ERRO - Colunas faltando na planilha base: {', '.join(missing_cols)}. Necessárias: {', '.join(required_cols)}"
                self.base_status_label.config(text=msg, foreground="red")
                self.log_message(msg, "ERROR")
                self.df_base = None
                return

            self.df_base.dropna(subset=['UC'], inplace=True)
            self.df_base['UC'] = self.df_base['UC'].astype(str).str.strip()

            num_ucs = len(self.df_base)
            if num_ucs == 0:
                msg = "Status: Planilha base carregada, mas sem UCs válidas após limpeza."
                self.base_status_label.config(text=msg, foreground="orange")
                self.log_message(msg, "WARNING")
            else:
                msg = f"Status: Planilha base carregada. {num_ucs} UCs encontradas."
                self.base_status_label.config(text=msg, foreground="green")
                self.log_message(msg, "INFO")
        except Exception as e:
            msg = f"Status: ERRO ao carregar planilha base - {e}"
            self.base_status_label.config(text=msg, foreground="red")
            self.log_message(msg, "CRITICAL_ERROR")
            self.df_base = None

    def select_pdfs(self):
        """Permite ao usuário selecionar múltiplos arquivos PDF."""
        files = filedialog.askopenfilenames(
            title="Selecione os arquivos PDF da Celesc",
            filetypes=(("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*"))
        )
        if files:
            self.pdf_files = list(files)
            self.pdf_label.config(text=f"{len(self.pdf_files)} PDF(s) selecionado(s)")
            self.log_message(f"{len(self.pdf_files)} PDF(s) selecionado(s).", "INFO")
        else:
            self.pdf_label.config(text="Nenhum PDF selecionado")
            self.log_message("Nenhum PDF selecionado.", "INFO")
            self.pdf_files = []

    def select_output_dir(self):
        """Permite ao usuário selecionar a pasta de saída para o relatório."""
        directory = filedialog.askdirectory(title="Selecione a pasta para salvar o relatório")
        if directory:
            self.output_dir = directory
            self.output_label.config(text=self.output_dir)
            self.log_message(f"Pasta de saída definida para: {self.output_dir}", "INFO")

    def start_processing(self):
        """Inicia o processo de extração e geração do relatório em uma nova thread."""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete('1.0', tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.log_message("Iniciando processo de verificação...", "INFO")

        self.load_base_sheet()
        if self.df_base is None or self.df_base.empty:
            msg = "Planilha base de UCs não carregada, inválida ou vazia. Verifique o arquivo 'base/ucs.sub.xlsx'."
            self.log_message(msg, "ERROR")
            messagebox.showerror("Erro de Configuração", msg)
            return
        if not self.pdf_files:
            msg = "Nenhum arquivo PDF foi selecionado para processamento."
            self.log_message(msg, "ERROR")
            messagebox.showerror("Erro de Configuração", msg)
            return
        if not self.output_dir or not os.path.isdir(self.output_dir):
            msg = "Pasta de saída inválida ou não definida."
            self.log_message(msg, "ERROR")
            messagebox.showerror("Erro de Configuração", msg)
            return

        self.status_label.config(text="Processando... Por favor, aguarde.")
        self.process_button.config(state=tk.DISABLED) # Desabilita o botão enquanto processa
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = len(self.pdf_files)
        self.root.update_idletasks() # Garante que a GUI seja atualizada antes da thread iniciar

        # Inicia o processamento real em uma thread separada
        processing_thread = threading.Thread(target=self._actual_processing_task)
        processing_thread.start()
    
    def _actual_processing_task(self):
        """Contém o loop principal de processamento de PDF, executa em uma thread separada."""
        all_extracted_data = []
        error_items = []
        total_pdfs = len(self.pdf_files)

        # Atualizações iniciais da GUI via root.after para garantir que rodem na thread principal
        self.root.after(0, lambda: self.progress_bar.config(value=0, maximum=total_pdfs))
        self.root.after(0, lambda: self.status_label.config(text=f"Processando {total_pdfs} PDFs..."))
        self.log_message(f"Iniciando processamento de {total_pdfs} PDFs...", "INFO")


        for i, pdf_path in enumerate(self.pdf_files):
            pdf_name = os.path.basename(pdf_path)
            # Atualiza elementos da GUI periodicamente da thread principal via root.after
            self.root.after(0, lambda p=pdf_name, idx=i: self.status_label.config(text=f"Processando PDF {idx+1}/{total_pdfs}: {p}"))
            self.root.after(0, lambda val=i+1: self.progress_bar.config(value=val))
            self.log_message(f"Processando PDF {i+1}/{total_pdfs}: {pdf_name}", "INFO")

            results_from_pdf = process_pdf_file(pdf_path, self.df_base, self.log_message)

            for item in results_from_pdf:
                if isinstance(item, dict):
                    if "error" in item:
                        # Adapta a estrutura de erro para as novas colunas
                        error_item = {
                            "UC": item.get("UC", "N/A"),
                            "Cod de Reg": "ERRO",
                            "Nome": "ERRO",
                            "ENERGIA (R$)": 0.0, # Nova coluna
                            "COSIP (R$)": 0.0,   # Nova coluna
                            "Valor Bruto (R$)": 0.0,
                            "RETENÇÃO (R$)": 0.0, # Renomeado
                            "LÍQUIDO (R$)": 0.0, # Renomeado
                            "Numero da Pagina": item.get("Numero da Pagina", os.path.basename(pdf_path)),
                            "Observação": item["error"]
                        }
                        error_items.append(error_item)
                    else:
                        all_extracted_data.append(item)
        
        # Após o loop, agenda a finalização na thread principal
        self.root.after(0, lambda: self._processing_complete(all_extracted_data, error_items))

    def _processing_complete(self, all_extracted_data, error_items):
        """Finaliza o processamento, cria o relatório Excel e atualiza a GUI."""
        
        # Define a ordem desejada das colunas para o relatório final (DADOS)
        final_columns_order_data = [
            "UC", "Cod de Reg", "Nome",
            "ENERGIA (R$)", # Nova coluna
            "COSIP (R$)",   # Nova coluna
            "Valor Bruto (R$)",
            "RETENÇÃO (R$)", # Renomeado
            "LÍQUIDO (R$)",  # Renomeado
            "Numero da Pagina"
        ]
        # Define a ordem desejada das colunas para o relatório de erros
        final_columns_order_errors = [
            "UC", "Cod de Reg", "Nome",
            "ENERGIA (R$)", # Nova coluna
            "COSIP (R$)",   # Nova coluna
            "Valor Bruto (R$)",
            "RETENÇÃO (R$)", # Renomeado
            "LÍQUIDO (R$)",  # Renomeado
            "Numero da Pagina",
            "Observação"
        ]

        # Lista de colunas que devem ser formatadas como moeda no Excel
        currency_cols_names_for_excel_fmt = [
            "ENERGIA (R$)",
            "COSIP (R$)",
            "Valor Bruto (R$)",
            "RETENÇÃO (R$)",
            "LÍQUIDO (R$)"
        ]

        df_report_data = pd.DataFrame(all_extracted_data)

        if not df_report_data.empty:
            # Garante que todas as colunas desejadas existam, adicionando-as se estiverem faltando
            for col in final_columns_order_data:
                if col not in df_report_data.columns:
                     df_report_data[col] = pd.NA
            df_report_data = df_report_data[final_columns_order_data] # Reordena as colunas

            # Converte colunas de moeda para numérico e preenche NaNs com 0.0
            for col_name in currency_cols_names_for_excel_fmt:
                if col_name in df_report_data.columns: # Verifica se a coluna existe no DataFrame
                    df_report_data[col_name] = pd.to_numeric(df_report_data[col_name], errors='coerce')
                    df_report_data[col_name] = df_report_data[col_name].fillna(0.0)
        else:
             df_report_data = pd.DataFrame(columns=final_columns_order_data)

        df_errors = pd.DataFrame()
        if error_items:
            df_errors = pd.DataFrame(error_items)
            # Garante que todas as colunas desejadas existam, adicionando-as se estiverem faltando
            for col in final_columns_order_errors:
                 if col not in df_errors.columns:
                    default_value = "" if col == "Observação" else (0.0 if col in currency_cols_names_for_excel_fmt else "N/A")
                    df_errors[col] = default_value
            df_errors = df_errors[final_columns_order_errors] # Reordena as colunas
        else:
             df_errors = pd.DataFrame(columns=final_columns_order_errors)

        output_file_path = os.path.join(self.output_dir, "Relatorio_Celesc.xlsx")

        try:
            self.log_message(f"Salvando relatório em: {output_file_path}", "INFO")
            with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                if not df_report_data.empty:
                    df_report_data.to_excel(writer, index=False, sheet_name='Relatorio_Dados_Extraidos')
                    workbook = writer.book
                    worksheet = writer.sheets['Relatorio_Dados_Extraidos']

                    # Aplica formato de moeda
                    for col_name_df in currency_cols_names_for_excel_fmt:
                        if col_name_df in df_report_data.columns: # Verifica se a coluna existe no DataFrame
                            # Encontra o índice da coluna Excel com base em final_columns_order_data
                            excel_col_idx = final_columns_order_data.index(col_name_df) + 1
                            col_letter = get_column_letter(excel_col_idx)
                            for row_num in range(2, worksheet.max_row + 1):
                                cell = worksheet[f'{col_letter}{row_num}']
                                if cell.value is not None and isinstance(cell.value, (int, float)):
                                    cell.number_format = 'R$ #,##0.00'
                    
                    # Ajusta a largura das colunas para Relatorio_Dados_Extraidos
                    for col_idx_df, col_name_df in enumerate(final_columns_order_data): # Usa final_columns_order_data para indexação correta
                        excel_col_idx = col_idx_df + 1
                        column_letter_val = get_column_letter(excel_col_idx)
                        max_len = 0
                        header_cell_val = worksheet[f'{column_letter_val}1'].value
                        if header_cell_val:
                             max_len = len(str(header_cell_val))
                        for row_num in range(2, worksheet.max_row + 1):
                            cell = worksheet[f'{column_letter_val}{row_num}']
                            if cell.value is not None:
                                cell_str_val = ""
                                if col_name_df in currency_cols_names_for_excel_fmt and isinstance(cell.value, (int, float)):
                                    # Formata para cálculo de comprimento, correspondendo ao formato R$ do Excel
                                    formatted_value_for_len = f"R$ {cell.value:_.2f}".replace('.',',').replace('_','.')
                                    if cell.value < 0:
                                         formatted_value_for_len = f"-R$ {abs(cell.value):_.2f}".replace('.',',').replace('_','.')
                                    cell_str_val = formatted_value_for_len
                                else:
                                    cell_str_val = str(cell.value)
                                max_len = max(max_len, len(cell_str_val))
                        adjusted_width = (max_len + 2) if max_len > 0 else 12
                        worksheet.column_dimensions[column_letter_val].width = adjusted_width
                
                if not df_errors.empty:
                    df_errors.to_excel(writer, index=False, sheet_name='Relatorio_Erros')
                    worksheet_errors = writer.sheets['Relatorio_Erros']
                    # Ajusta a largura das colunas para Relatorio_Erros
                    for col_idx_df, col_name_df in enumerate(final_columns_order_errors): # Usa final_columns_order_errors
                        excel_col_idx = col_idx_df + 1
                        column_letter_val = get_column_letter(excel_col_idx)
                        max_len = len(str(worksheet_errors[f'{column_letter_val}1'].value))
                        for row_num in range(2, worksheet_errors.max_row + 1):
                             cell = worksheet_errors[f'{column_letter_val}{row_num}']
                             if cell.value is not None:
                                max_len = max(max_len, len(str(cell.value)))
                        adjusted_width = (max_len + 2) if max_len > 0 else 12
                        if col_name_df == "Observação":
                            adjusted_width = min(adjusted_width, 80) # Limita a largura para 'Observação'
                        worksheet_errors.column_dimensions[column_letter_val].width = adjusted_width


            num_registros_extraidos = len(df_report_data)
            num_erros_reportados = len(df_errors)
            
            summary_message = f"Processamento concluído!\n"
            if num_registros_extraidos > 0:
                summary_message += f"{num_registros_extraidos} registros de fatura extraídos com sucesso na aba 'Relatorio_Dados_Extraidos'.\n"
            else:
                 summary_message += "Nenhum registro de fatura válido foi extraído.\n"

            if num_erros_reportados > 0:
                summary_message += f"{num_erros_reportados} problemas/erros encontrados durante o processamento, listados na aba 'Relatorio_Erros'."
                self.log_message(f"Processamento concluído com {num_erros_reportados} problemas/erros.", "WARNING")
                messagebox.showwarning("Processamento Concluído com Alertas", summary_message + f"\nRelatório salvo em:\n{output_file_path}")
            elif num_registros_extraidos == 0:
                summary_message = "Nenhum dado foi processado. Verifique se selecionou PDFs e se eles contêm faturas individuais com UCs identificáveis (além da página de sumário)."
                self.log_message(summary_message, "INFO")
                messagebox.showinfo("Processamento Concluído", summary_message)
            else:
                self.log_message("Processamento concluído com sucesso!", "SUCCESS")
                messagebox.showinfo("Processamento Concluído", summary_message + f"\nRelatório salvo em:\n{output_file_path}")

            self.status_label.config(text="Concluído. Relatório gerado.")

            if os.path.exists(output_file_path):
                try:
                    if sys.platform == "win32":
                        os.startfile(output_file_path)
                    elif sys.platform == "darwin":
                        subprocess.call(("open", output_file_path))
                    else:
                        subprocess.call(("xdg-open", output_file_path))
                except Exception as open_e:
                    self.log_message(f"Não foi possível abrir o relatório automaticamente: {open_e}", "WARNING")

        except Exception as e:
            self.log_message(f"Erro CRÍTICO ao salvar o relatório Excel: {e}", "CRITICAL_ERROR")
            messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o relatório Excel: {e}")
            self.status_label.config(text="Erro ao salvar relatório.")
        finally:
            self.process_button.config(state=tk.NORMAL) # Reabilita o botão em caso de erro ao salvar
            self.progress_bar.config(value=0) # Reseta a barra de progresso

if __name__ == "__main__":
    root = tk.Tk()
    app = AppCelescReporter(root)
    root.mainloop()