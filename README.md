# 📑 Gerador de Relatório de Faturas Celesc

Este projeto é uma aplicação em Python com interface gráfica (GUI) que permite extrair dados de faturas da Celesc em formato PDF e gerar um relatório consolidado em Excel. É especialmente útil para empresas ou usuários que precisam auditar, conferir ou organizar informações de múltiplas contas de energia.

## 🚀 Funcionalidades

- 📄 Processa vários arquivos PDF de faturas da Celesc.
- 🔍 Faz a leitura dos dados:
  - Unidade Consumidora (UC)
  - Código de Registro
  - Nome (da base de dados)
  - Valor Líquido da Fatura
  - Desconto de Tributos Retidos (IRPJ, PIS, COFINS, CSLL)
  - Valor Bruto calculado (Líquido + Descontos)
- ⚙️ Verifica se os dados da fatura estão cadastrados na planilha base (`base/ucs.sub.xlsx`).
- 📤 Gera relatório Excel com:
  - Aba de dados extraídos (`Relatorio_Dados_Extraidos`)
  - Aba de erros/problemas encontrados (`Relatorio_Erros`)
- ✅ Interface intuitiva e fácil de usar.
- 🔔 Log em tempo real durante o processamento.

---

## 🛠️ Tecnologias Utilizadas

- Python
- Tkinter (interface gráfica)
- Pandas
- pdfplumber
- openpyxl (para manipulação avançada do Excel)
- Regex para extração de dados

---

## 📦 Instalação

### 🔗 Pré-requisitos

- Python 3.8 ou superior instalado.  
Se não tiver, [baixe aqui](https://www.python.org/downloads/).

### 🧠 Instale as dependências

Execute no terminal ou prompt de comando:

