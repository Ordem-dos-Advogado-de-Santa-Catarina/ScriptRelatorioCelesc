import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import pdfplumber
import re
import os
import subprocess
import sys
import logging # Mantido, embora não configurado para output no exemplo

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
    # Tenta o padrão mais específico primeiro
    match = re.search(r"Grupo / Subgrupo Tensão:.*?Valor:\s*R\$\s*([\d\.,]+)", text_block, re.DOTALL | re.IGNORECASE)
    if match:
        return parse_value(match.group(1))
    # Fallback para um padrão mais genérico se o específico não for encontrado
    match_fallback = re.search(r"TOTAL A PAGAR\s*R\$\s*([\d\.,]+)", text_block, re.IGNORECASE)
    if match_fallback:
        return parse_value(match_fallback.group(1))
    # Outro fallback comum
    match_fallback_2 = re.search(r"Valor:\s*R\$\s*([\d\.,]+)", text_block, re.IGNORECASE)
    if match_fallback_2:
        return parse_value(match_fallback_2.group(1))
    return 0.0

def extract_item_value_from_block(text_block, item_name_pattern):
    """
    Extrai o valor da coluna 'Valor (R$)' para um item específico da seção 'Itens da Fatura'.
    Funciona mesmo que tenha espaçamentos irregulares ou colunas faltantes antes do valor.
    """
    # Normaliza múltiplos espaços para um único espaço, facilitando a regex.
    # Isso ajuda a lidar com desalinhamentos e OCRs imperfeitos.
    cleaned_text_block = re.sub(r'\s+', ' ', text_block)

    # A regex busca o 'item_name_pattern', seguido por quaisquer caracteres (.*?) não gananciosos,
    # e então captura um número que pode ser negativo (o sinal '-' é opcional)
    # e pode conter pontos ou vírgulas, que está no final da linha ($).
    # re.escape() é usado para garantir que 'item_name_pattern' seja tratado como uma string literal,
    # caso contenha caracteres especiais de regex.
    pattern = rf"{re.escape(item_name_pattern)}.*?(-?[\d\.,]+)\s*$" # Adicionado \s*$ para lidar com espaços no final da linha

    match = re.search(pattern, cleaned_text_block, re.MULTILINE | re.IGNORECASE)

    if match:
        return parse_value(match.group(1))

    return 0.0


def extract_fatura_data_from_text_block(text_block, df_base, pdf_filename_for_error_logging, logger_func=None):
    """
    Extrai todos os dados de uma fatura a partir de um bloco de texto.
    Retorna um dicionário com os dados ou um dicionário de erro.
    """
    uc_number = extract_uc_from_block(text_block)
    if not uc_number:
        # Se não há UC no bloco, não podemos prosseguir com este bloco específico.
        # Isso pode acontecer se o bloco de texto for parte de uma página sem uma fatura completa.
        return None 

    base_info = df_base[df_base['UC'].astype(str) == uc_number]
    if base_info.empty:
        error_msg = f"UC {uc_number} (de {pdf_filename_for_error_logging}) não encontrada na planilha base."
        if logger_func:
            logger_func(error_msg, "ERROR")
        return {"error": error_msg, "UC": uc_number, "pdf_filename": pdf_filename_for_error_logging}

    cod_reg = base_info['Cod de Reg'].iloc[0]
    nome_base = base_info['Nome'].iloc[0]

    # Este é o Valor Líquido da Fatura, conforme extraído do cabeçalho/rodapé da fatura.
    valor_liquido_fatura = extract_valor_total_fatura_from_block(text_block)
    if valor_liquido_fatura == 0.0 and logger_func:
        logger_func(f"AVISO: Valor Líquido da fatura (Valor Total da Fatura) não encontrado ou zerado para UC {uc_number} em {pdf_filename_for_error_logging}. Verifique o PDF.", "WARNING")

    # Padrões para os tributos retidos na seção "Itens da Fatura"
    tributos_retidos_patterns = {
        "IRPJ": r"Tributo Retido IRPJ",
        "PIS": r"Tributo Retido PIS",
        "COFINS": r"Tributo Retido COFINS",
        "CSLL": r"Tributo Retido CSLL"
    }

    soma_valores_negativos_tributos = 0.0
    # Dicionário para armazenar valores individuais dos tributos (opcional, para debug ou uso futuro)
    # valores_tributos_extraidos = {}

    for nome_tributo, pattern_str in tributos_retidos_patterns.items():
        # extract_item_value_from_block deve retornar o valor da coluna "Valor (R$)"
        # para esses itens, que é esperado ser negativo.
        valor_tributo = extract_item_value_from_block(text_block, pattern_str)
        # valores_tributos_extraidos[nome_tributo] = valor_tributo
        soma_valores_negativos_tributos += valor_tributo # Acumula os valores (negativos)

    # O "Desconto de Tributos Retidos (R$)" é a soma dos valores absolutos dos tributos,
    # ou seja, o oposto da soma dos valores negativos.
    desconto_total_tributos_retidos = abs(soma_valores_negativos_tributos)

    # Valor Bruto (R$) = Valor Líquido (R$) + Desconto de Tributos Retidos (R$)
    valor_bruto_fatura_calculado = valor_liquido_fatura + desconto_total_tributos_retidos
    
    # Log se o desconto for zero, pode indicar que os itens não foram encontrados
    if desconto_total_tributos_retidos == 0.0 and soma_valores_negativos_tributos == 0.0 and logger_func:
        # Verifica se algum valor de tributo foi encontrado como não zero, caso contrário pode ser um falso aviso
        found_any_tax_value_non_zero = any(extract_item_value_from_block(text_block, p) != 0.0 for p in tributos_retidos_patterns.values())
        if not found_any_tax_value_non_zero:
             logger_func(f"INFO: Nenhum valor de tributo retido encontrado para UC {uc_number} em {pdf_filename_for_error_logging}. 'Desconto de Tributos Retidos' será 0.00.", "INFO")


    return {
        "UC": uc_number,
        "Cod de Reg": cod_reg,
        "Nome": nome_base,
        "Valor Líquido (R$)": valor_liquido_fatura,
        "Desconto de Tributos Retidos (R$)": desconto_total_tributos_retidos,
        "Valor Bruto (R$)": valor_bruto_fatura_calculado,
        "pdf_filename": pdf_filename_for_error_logging
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
                results_for_this_pdf.append({"error": error_msg, "pdf_filename": pdf_filename})
                return results_for_this_pdf

            for page_num, page in enumerate(pdf.pages):
                page_text = page.extract_text(x_tolerance=1, y_tolerance=3) # Tolerâncias podem ajudar na extração
                if not page_text or not page_text.strip():
                    logger_func(f"Página {page_num + 1} de {pdf_filename} não contém texto extraível.", "INFO")
                    continue

                # Tenta encontrar blocos de fatura baseados na presença de "UC:" ou "Unidade Consumidora:"
                # Isso ajuda a processar PDFs com múltiplas faturas por página.
                uc_pattern = r"(?:UC:|Unidade Consumidora:)\s*\d+"
                matches = list(re.finditer(uc_pattern, page_text))

                if not matches:
                    # Se nenhuma UC explícita for encontrada na página, pode ser uma fatura que ocupa a página inteira
                    # ou uma página sem dados de fatura. Tenta processar a página inteira como um bloco.
                    logger_func(f"Nenhuma UC explícita na página {page_num+1} de {pdf_filename}. Tentando processar a página inteira como um bloco único.", "INFO")
                    fatura_data = extract_fatura_data_from_text_block(page_text, df_base, pdf_filename, logger_func)
                    if fatura_data and "error" not in fatura_data : # Adiciona apenas se dados válidos forem extraídos
                        results_for_this_pdf.append(fatura_data)
                    elif fatura_data and "error" in fatura_data: # Adiciona se for um erro conhecido
                        results_for_this_pdf.append(fatura_data)
                    # Se fatura_data for None (sem UC no bloco), não adiciona nada e continua.
                    continue

                # Processa cada bloco de fatura identificado na página
                for i, match in enumerate(matches):
                    start_block = match.start()
                    # Determina o fim do bloco atual: é o início do próximo match de UC, ou o fim do texto da página.
                    end_block = matches[i+1].start() if i + 1 < len(matches) else len(page_text)
                    current_text_block = page_text[start_block:end_block]

                    fatura_data = extract_fatura_data_from_text_block(current_text_block, df_base, pdf_filename, logger_func)
                    if fatura_data and "error" not in fatura_data :
                        results_for_this_pdf.append(fatura_data)
                    elif fatura_data and "error" in fatura_data:
                         results_for_this_pdf.append(fatura_data)


            if not results_for_this_pdf and not any("error" in r for r in results_for_this_pdf if isinstance(r,dict)):
                 # Se após processar todas as páginas, nenhum dado de fatura VÁLIDO foi adicionado
                 # E não houve erros de UC não encontrada ou PDF sem páginas já logados
                 no_data_msg = f"Nenhum dado de fatura (com UC identificável) encontrado em {pdf_filename} após processar todas as páginas."
                 logger_func(no_data_msg, "WARNING")
                 # Adiciona um item de erro genérico para este PDF se nenhum outro erro específico foi registrado
                 if not any(isinstance(r, dict) and r.get("pdf_filename") == pdf_filename for r in results_for_this_pdf):
                    results_for_this_pdf.append({"error": no_data_msg, "pdf_filename": pdf_filename, "UC": "N/A"})


    except Exception as e:
        critical_error_msg = f"Erro crítico ao processar {pdf_filename}: {e}"
        logger_func(critical_error_msg, "CRITICAL_ERROR")
        results_for_this_pdf.append({"error": critical_error_msg, "pdf_filename": pdf_filename, "UC": "N/A"})

    return results_for_this_pdf


# --- Classe da Interface Gráfica ---
class AppCelescReporter:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("Gerador de Relatório Celesc")
        self.center_window(700, 650)

        self.base_sheet_path = os.path.join(os.path.dirname(sys.argv[0]), "base", "ucs.sub.xlsx")
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

        log_frame = ttk.LabelFrame(main_frame, text="Log de Processamento", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.tag_config("INFO", foreground="black")
        self.log_text.tag_config("WARNING", foreground="orange")
        self.log_text.tag_config("ERROR", foreground="red")
        self.log_text.tag_config("CRITICAL_ERROR", foreground="red", font=('TkDefaultFont', 9, 'bold'))
        self.log_text.tag_config("SUCCESS", foreground="green")
        self.log_text.tag_config("DEBUG", foreground="gray")


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
        try:
            if not os.path.exists(self.base_sheet_path):
                msg = f"Status: ERRO - Arquivo base não encontrado em {self.base_sheet_path}"
                self.base_status_label.config(text=msg, foreground="red")
                if hasattr(self, 'log_text'): self.log_message(msg, "ERROR")
                self.df_base = None
                return

            self.df_base = pd.read_excel(self.base_sheet_path, dtype={'UC': str, 'Cod de Reg': str, 'Nome': str})
            required_cols = ['UC', 'Cod de Reg', 'Nome']
            if not all(col in self.df_base.columns for col in required_cols):
                missing_cols = [col for col in required_cols if col not in self.df_base.columns]
                msg = f"Status: ERRO - Colunas faltando na planilha base: {', '.join(missing_cols)}"
                self.base_status_label.config(text=msg, foreground="red")
                if hasattr(self, 'log_text'): self.log_message(msg, "ERROR")
                self.df_base = None
                return

            self.df_base.dropna(subset=['UC'], inplace=True)
            self.df_base['UC'] = self.df_base['UC'].astype(str).str.strip()

            num_ucs = len(self.df_base)
            if num_ucs == 0:
                msg = "Status: Planilha base carregada, mas sem UCs válidas."
                self.base_status_label.config(text=msg, foreground="orange")
                if hasattr(self, 'log_text'): self.log_message(msg, "WARNING")
            else:
                msg = f"Status: Planilha base carregada. {num_ucs} UCs encontradas."
                self.base_status_label.config(text=msg, foreground="green")
                if hasattr(self, 'log_text'): self.log_message(msg, "INFO")
        except Exception as e:
            msg = f"Status: ERRO ao carregar planilha base - {e}"
            self.base_status_label.config(text=msg, foreground="red")
            if hasattr(self, 'log_text'): self.log_message(msg, "CRITICAL_ERROR")
            self.df_base = None

    def select_pdfs(self):
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
        directory = filedialog.askdirectory(title="Selecione a pasta para salvar o relatório")
        if directory:
            self.output_dir = directory
            self.output_label.config(text=self.output_dir)
            self.log_message(f"Pasta de saída definida para: {self.output_dir}", "INFO")

    def start_processing(self):
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
        self.log_message("Iniciando processamento dos PDFs...", "INFO")
        self.process_button.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = len(self.pdf_files)
        self.root.update_idletasks()

        all_extracted_data = [] # Renomeado de all_processed_data para clareza
        error_items = [] # Lista para armazenar itens de erro separadamente

        for i, pdf_path in enumerate(self.pdf_files):
            pdf_name = os.path.basename(pdf_path)
            self.status_label.config(text=f"Processando PDF {i+1}/{len(self.pdf_files)}: {pdf_name}")
            self.log_message(f"Processando PDF {i+1}/{len(self.pdf_files)}: {pdf_name}", "INFO")
            self.progress_bar["value"] = i + 1
            self.root.update_idletasks()

            # results_from_pdf é uma lista de dicionários (dados ou erros)
            results_from_pdf = process_pdf_file(pdf_path, self.df_base, self.log_message)

            for item in results_from_pdf:
                if isinstance(item, dict) and "error" in item:
                    # Se o item é um erro, adiciona à lista de erros.
                    # Certifique-se de que 'UC' e 'pdf_filename' estão presentes para o relatório de erros.
                    error_item = {
                        "UC": item.get("UC", "N/A"), # Pega UC do item de erro, ou N/A
                        "Cod de Reg": "ERRO",
                        "Nome": "ERRO",
                        "Valor Líquido (R$)": 0.0,
                        "Desconto de Tributos Retidos (R$)": 0.0,
                        "Valor Bruto (R$)": 0.0,
                        "pdf_filename": item.get("pdf_filename", pdf_name),
                        "Observação": item["error"]
                    }
                    error_items.append(error_item)
                elif isinstance(item, dict): # Item de dados válido
                    all_extracted_data.append(item)
        
        self.process_button.config(state=tk.NORMAL)

        if not all_extracted_data and not error_items: # Caso nenhum PDF produziu dados ou erros (improvável com a lógica atual)
            final_msg = "Nenhum dado de fatura ou erro foi retornado do processamento dos PDFs."
            self.log_message(final_msg, "WARNING")
            messagebox.showwarning("Processamento Concluído", final_msg)
            self.status_label.config(text="Concluído. Nenhum dado.")
            return
        
        if not all_extracted_data and error_items: # Apenas erros foram encontrados
            final_msg = "Nenhum dado de fatura válido foi extraído. Apenas erros foram reportados."
            self.log_message(final_msg, "WARNING")
            # Prossegue para gerar o relatório de erros.
        
        # Define as colunas finais do relatório
        final_columns_order = [
            "UC", "Cod de Reg", "Nome",
            "Valor Líquido (R$)",
            "Desconto de Tributos Retidos (R$)",
            "Valor Bruto (R$)",
            "pdf_filename"
        ]
        # Colunas para o relatório de erros
        error_columns_order = final_columns_order + ["Observação"]


        df_report = pd.DataFrame(all_extracted_data)
        
        # Garante que todas as colunas esperadas existam no DataFrame de dados válidos, preenchendo com NA se faltar
        for col in final_columns_order:
            if col not in df_report.columns and not df_report.empty: # Adiciona coluna apenas se o df não estiver vazio
                df_report[col] = pd.NA
        if not df_report.empty:
            df_report = df_report[final_columns_order] # Reordena/seleciona colunas

        # Colunas que devem ser formatadas como moeda
        currency_cols_names = [
            "Valor Líquido (R$)",
            "Desconto de Tributos Retidos (R$)",
            "Valor Bruto (R$)"
        ]
        if not df_report.empty:
            for col_name in currency_cols_names:
                if col_name in df_report.columns:
                    df_report[col_name] = pd.to_numeric(df_report[col_name], errors='coerce')

        # Cria DataFrame para erros, se houver
        df_errors = pd.DataFrame()
        if error_items:
            df_errors = pd.DataFrame(error_items)
            # Garante colunas e ordem para o df_errors
            for col in error_columns_order:
                 if col not in df_errors.columns:
                    df_errors[col] = "" if col == "Observação" else (0.0 if col in currency_cols_names else "N/A")
            df_errors = df_errors[error_columns_order]


        output_file_path = os.path.join(self.output_dir, "Relatorio_Celesc.xlsx")

        try:
            self.log_message(f"Salvando relatório em: {output_file_path}", "INFO")
            with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                if not df_report.empty:
                    df_report.to_excel(writer, index=False, sheet_name='Relatorio_Dados_Extraidos')
                    workbook = writer.book
                    worksheet = writer.sheets['Relatorio_Dados_Extraidos']

                    df_col_to_excel_col_idx = {col_name: idx + 1 for idx, col_name in enumerate(df_report.columns)}

                    for col_name_df in currency_cols_names:
                        if col_name_df in df_col_to_excel_col_idx:
                            excel_col_idx = df_col_to_excel_col_idx[col_name_df]
                            col_letter = get_column_letter(excel_col_idx)
                            for row_num in range(2, worksheet.max_row + 1):
                                cell = worksheet[f'{col_letter}{row_num}']
                                if cell.value is not None and isinstance(cell.value, (int, float)):
                                    cell.number_format = 'R$ #,##0.00'
                                    cell.alignment = Alignment(horizontal='right')
                    
                    for col_idx_df, col_name_df in enumerate(df_report.columns):
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
                                if col_name_df in currency_cols_names and isinstance(cell.value, (int, float)):
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
                    # Autoajuste para a planilha de erros
                    for col_idx_df, col_name_df in enumerate(df_errors.columns):
                        excel_col_idx = col_idx_df + 1
                        column_letter_val = get_column_letter(excel_col_idx)
                        max_len = len(str(worksheet_errors[f'{column_letter_val}1'].value)) # Cabeçalho
                        for row_num in range(2, worksheet_errors.max_row + 1):
                             cell = worksheet_errors[f'{column_letter_val}{row_num}']
                             if cell.value is not None:
                                max_len = max(max_len, len(str(cell.value)))
                        adjusted_width = (max_len + 2) if max_len > 0 else 12
                        worksheet_errors.column_dimensions[column_letter_val].width = adjusted_width


            num_registros_extraidos = len(all_extracted_data)
            num_erros_reportados = len(error_items)
            
            summary_message = f"Processamento concluído!\n"
            if num_registros_extraidos > 0:
                summary_message += f"{num_registros_extraidos} registros de fatura extraídos com sucesso.\n"
            else:
                 summary_message += "Nenhum registro de fatura válido foi extraído.\n"

            if num_erros_reportados > 0:
                summary_message += f"{num_erros_reportados} problemas/erros encontrados durante o processamento.\n"
                summary_message += "Verifique a aba 'Relatorio_Erros' no arquivo Excel e o log na janela para detalhes."
                self.log_message(f"Processamento concluído com {num_erros_reportados} problemas/erros.", "WARNING")
                messagebox.showwarning("Processamento Concluído com Alertas", summary_message + f"\nRelatório salvo em:\n{output_file_path}")
            elif num_registros_extraidos == 0 and num_erros_reportados == 0: # Nenhum PDF, ou PDFs vazios sem erros
                summary_message = "Nenhum dado foi processado. Verifique se selecionou PDFs e se eles contêm faturas."
                self.log_message(summary_message, "INFO")
                messagebox.showinfo("Processamento Concluído", summary_message)
            else: # Apenas sucessos
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

if __name__ == "__main__":
    root = tk.Tk()
    app = AppCelescReporter(root)
    root.mainloop()