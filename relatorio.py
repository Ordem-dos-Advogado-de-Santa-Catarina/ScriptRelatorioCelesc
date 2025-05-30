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
    # Workbook não é diretamente instanciado, mas writer.book é um objeto Workbook
    # from openpyxl import Workbook # Removido pois não é instanciado diretamente
    from openpyxl.utils import get_column_letter # Usado para autoajuste de colunas.
    from openpyxl.styles import Alignment         # Usado para alinhar as células de moeda.
    # Font e PatternFill não foram usados
    # FORMAT_CURRENCY_BRL foi substituído pela string direta 'R$ #,##0.00'
except ImportError:
    messagebox.showerror("Dependência Faltando",
                         "A biblioteca 'openpyxl' é necessária para formatação avançada do Excel. "
                         "Por favor, instale-a com 'pip install openpyxl' e tente novamente.")
    sys.exit(1)


# Configurar logging (opcional, mas útil para depuração)
# logging.basicConfig(level=logging.WARNING)
# logging.getLogger("pdfminer").setLevel(logging.WARNING)

# --- Funções de Extração e Processamento (sem grandes alterações na lógica central) ---

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
    full_regex_pattern = rf"^(?:{item_name_pattern})\s+([\d\.,-]+)\s+([\d\.,-]+)\s+([\d\.,-]+)"
    match = re.search(full_regex_pattern, text_block, re.MULTILINE | re.IGNORECASE)
    if match:
        return parse_value(match.group(3))
    return 0.0

def extract_fatura_data_from_text_block(text_block, df_base, pdf_filename_for_error_logging, logger_func=None):
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
        return {"error": error_msg, "UC": uc_number, "pdf_filename": pdf_filename_for_error_logging}


    cod_reg = base_info['Cod de Reg'].iloc[0]
    nome_base = base_info['Nome'].iloc[0]

    valor_total_fatura = extract_valor_total_fatura_from_block(text_block)
    if valor_total_fatura == 0.0 and logger_func:
        logger_func(f"AVISO: Valor total da fatura não encontrado ou zerado para UC {uc_number} em {pdf_filename_for_error_logging}", "WARNING")


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
        soma_negativos += val

    valor_bruto_calculado = soma_positivos
    valor_liquido_calculado = valor_bruto_calculado + soma_negativos

    return {
        "UC": uc_number,
        "Cod de Reg": cod_reg,
        "Nome": nome_base,
        "Valor Total Fatura (R$)": valor_total_fatura,
        "Valor Bruto Calculado (R$)": valor_bruto_calculado,
        "Valor Líquido Calculado (R$)": valor_liquido_calculado,
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
                page_text = page.extract_text()
                if not page_text or not page_text.strip():
                    logger_func(f"Página {page_num + 1} de {pdf_filename} não contém texto extraível.", "INFO")
                    continue

                uc_pattern = r"(?:UC:|Unidade Consumidora:)\s*\d+"
                matches = list(re.finditer(uc_pattern, page_text))

                if not matches:
                    logger_func(f"Nenhuma UC explícita na página {page_num+1} de {pdf_filename}. Tentando processar a página inteira.", "INFO")
                    fatura_data = extract_fatura_data_from_text_block(page_text, df_base, pdf_filename, logger_func)
                    if fatura_data:
                        results_for_this_pdf.append(fatura_data)
                    continue

                for i, match in enumerate(matches):
                    start_block = match.start()
                    end_block = matches[i+1].start() if i + 1 < len(matches) else len(page_text)
                    current_text_block = page_text[start_block:end_block]

                    fatura_data = extract_fatura_data_from_text_block(current_text_block, df_base, pdf_filename, logger_func)
                    if fatura_data:
                        results_for_this_pdf.append(fatura_data)

            if not results_for_this_pdf:
                 no_data_msg = f"Nenhum dado de fatura encontrado em {pdf_filename} após processar todas as páginas."
                 logger_func(no_data_msg, "WARNING")
                 results_for_this_pdf.append({"error": no_data_msg, "pdf_filename": pdf_filename})


    except Exception as e:
        critical_error_msg = f"Erro crítico ao processar {pdf_filename}: {e}"
        logger_func(critical_error_msg, "CRITICAL_ERROR")
        results_for_this_pdf.append({"error": critical_error_msg, "pdf_filename": pdf_filename})

    return results_for_this_pdf


# --- Classe da Interface Gráfica ---
class AppCelescReporter:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("Gerador de Relatório Celesc")
        self.center_window(700, 650) # Aumentada para área de log

        self.base_sheet_path = os.path.join(os.path.dirname(sys.argv[0]), "base", "ucs.sub.xlsx")
        self.df_base = None
        self.pdf_files = []
        self.output_dir = os.path.join(os.path.expanduser("~"), "Desktop")

        style = ttk.Style(self.root)
        style.theme_use('clam')

        main_frame = ttk.Frame(self.root, padding="10 10 10 10")
        main_frame.pack(expand=True, fill=tk.BOTH)

        # --- Seção Planilha Base ---
        base_frame = ttk.LabelFrame(main_frame, text="Planilha Base de UCs", padding="10")
        base_frame.pack(fill=tk.X, pady=5)
        self.base_path_label = ttk.Label(base_frame, text=f"Caminho: {self.base_sheet_path}", wraplength=650)
        self.base_path_label.pack(fill=tk.X)
        self.base_status_label = ttk.Label(base_frame, text="Status: Não carregada")
        self.base_status_label.pack(fill=tk.X)
        self.load_base_sheet()

        # --- Seção PDFs ---
        pdf_frame = ttk.LabelFrame(main_frame, text="Arquivos PDF das Faturas", padding="10")
        pdf_frame.pack(fill=tk.X, pady=5)
        self.pdf_button = ttk.Button(pdf_frame, text="Selecionar PDFs da Celesc", command=self.select_pdfs)
        self.pdf_button.pack(side=tk.LEFT, padx=(0,10))
        self.pdf_label = ttk.Label(pdf_frame, text="Nenhum PDF selecionado")
        self.pdf_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # --- Seção Pasta de Saída ---
        output_frame = ttk.LabelFrame(main_frame, text="Pasta de Saída do Relatório", padding="10")
        output_frame.pack(fill=tk.X, pady=5)
        self.output_dir_button = ttk.Button(output_frame, text="Definir Pasta de Saída", command=self.select_output_dir)
        self.output_dir_button.pack(side=tk.LEFT, padx=(0,10))
        self.output_label = ttk.Label(output_frame, text=f"Padrão: {self.output_dir}")
        self.output_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # --- Seção Ações e Progresso ---
        action_frame = ttk.Frame(main_frame, padding="10")
        action_frame.pack(fill=tk.X, pady=10)

        self.progress_bar = ttk.Progressbar(action_frame, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.pack(pady=5, fill=tk.X)

        self.process_button = ttk.Button(action_frame, text="Iniciar Processamento de Relatório", command=self.start_processing)
        self.process_button.pack(pady=5)

        self.status_label = ttk.Label(action_frame, text="Aguardando configuração...")
        self.status_label.pack(fill=tk.X, pady=5)

        # --- Seção Log em Tempo Real ---
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
        self.root.update_idletasks() # Garante atualização da UI

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
        self.log_text.delete('1.0', tk.END) # Limpa log anterior
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

        all_processed_data = []
        error_count = 0

        for i, pdf_path in enumerate(self.pdf_files):
            pdf_name = os.path.basename(pdf_path)
            self.status_label.config(text=f"Processando PDF {i+1}/{len(self.pdf_files)}: {pdf_name}")
            self.log_message(f"Processando PDF {i+1}/{len(self.pdf_files)}: {pdf_name}", "INFO")
            self.progress_bar["value"] = i + 1
            self.root.update_idletasks()

            results_from_pdf = process_pdf_file(pdf_path, self.df_base, self.log_message)

            for item in results_from_pdf:
                if isinstance(item, dict) and "error" in item:
                    error_count +=1
                elif isinstance(item, dict):
                    all_processed_data.append(item)

        self.process_button.config(state=tk.NORMAL)

        if not all_processed_data:
            final_msg = "Nenhum dado de fatura válido foi extraído de nenhum PDF."
            self.log_message(final_msg, "WARNING")
            messagebox.showwarning("Processamento Concluído", final_msg)
            self.status_label.config(text="Concluído. Nenhum dado válido.")
            return

        df_report = pd.DataFrame(all_processed_data)

        final_columns = [
            "UC", "Cod de Reg", "Nome",
            "Valor Total Fatura (R$)", "Valor Bruto Calculado (R$)", "Valor Líquido Calculado (R$)",
            "pdf_filename"
        ]
        for col in final_columns:
            if col not in df_report.columns:
                df_report[col] = pd.NA
        df_report = df_report[final_columns]

        currency_cols_names = ["Valor Total Fatura (R$)", "Valor Bruto Calculado (R$)", "Valor Líquido Calculado (R$)"]
        for col_name in currency_cols_names:
            if col_name in df_report.columns:
                # Converte para numérico, erros viram NaN (que se torna célula vazia no Excel)
                df_report[col_name] = pd.to_numeric(df_report[col_name], errors='coerce')


        output_file_path = os.path.join(self.output_dir, "Relatorio_Celesc.xlsx")

        try:
            self.log_message(f"Salvando relatório em: {output_file_path}", "INFO")
            with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                df_report.to_excel(writer, index=False, sheet_name='Relatorio')
                workbook = writer.book
                worksheet = writer.sheets['Relatorio']

                # Formatação de Moeda e Alinhamento
                # Cria um mapeamento de nome de coluna do DataFrame para índice da coluna no Excel (1-based)
                df_col_to_excel_col_idx = {col_name: idx + 1 for idx, col_name in enumerate(df_report.columns)}

                for col_name_df in currency_cols_names:
                    if col_name_df in df_col_to_excel_col_idx:
                        excel_col_idx = df_col_to_excel_col_idx[col_name_df]
                        col_letter = get_column_letter(excel_col_idx)
                        for row_num in range(2, worksheet.max_row + 1): # Começa da linha 2 (abaixo do cabeçalho)
                            cell = worksheet[f'{col_letter}{row_num}']
                            # Aplica formatação apenas se a célula tiver um valor numérico
                            if cell.value is not None and isinstance(cell.value, (int, float)):
                                cell.number_format = 'R$ #,##0.00'
                                cell.alignment = Alignment(horizontal='right')

                # Autoajuste de largura das colunas
                for col_idx_df, col_name_df in enumerate(df_report.columns):
                    excel_col_idx = col_idx_df + 1 # openpyxl é 1-based
                    column_letter_val = get_column_letter(excel_col_idx)
                    max_len = 0

                    # Considerar o cabeçalho (primeira linha)
                    header_cell_val = worksheet[f'{column_letter_val}1'].value
                    if header_cell_val:
                         max_len = len(str(header_cell_val))

                    # Iterar sobre as células da coluna para encontrar o comprimento máximo
                    for row_num in range(2, worksheet.max_row + 1):
                        cell = worksheet[f'{column_letter_val}{row_num}']
                        if cell.value is not None:
                            cell_str_val = ""
                            # Se for uma coluna de moeda e já formatada, calcula o comprimento da string formatada
                            if col_name_df in currency_cols_names and isinstance(cell.value, (int, float)):
                                # Simula o formato brasileiro para cálculo de comprimento
                                # Ex: 1234.5 -> "R$ 1.234,50"
                                formatted_value_for_len = f"R$ {cell.value:_.2f}".replace('.',',').replace('_','.')
                                if cell.value < 0: # Pequeno ajuste para negativo, se necessário
                                     formatted_value_for_len = f"-R$ {abs(cell.value):_.2f}".replace('.',',').replace('_','.')
                                cell_str_val = formatted_value_for_len
                            else:
                                cell_str_val = str(cell.value)
                            max_len = max(max_len, len(cell_str_val))

                    adjusted_width = (max_len + 2) if max_len > 0 else 12 # Um mínimo para colunas vazias
                    worksheet.column_dimensions[column_letter_val].width = adjusted_width

            summary_message = f"Processamento concluído!\n{len(all_processed_data)} registros de fatura extraídos.\nRelatório salvo em:\n{output_file_path}"
            if error_count > 0:
                summary_message += f"\n\nForam encontrados {error_count} problemas/erros durante o processamento. Verifique o log na janela para detalhes."
                self.log_message(f"Processamento concluído com {error_count} problemas/erros.", "WARNING")
                messagebox.showwarning("Processamento Concluído com Alertas", summary_message)
            else:
                self.log_message("Processamento concluído com sucesso!", "SUCCESS")
                messagebox.showinfo("Processamento Concluído", summary_message)

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