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

# 📑 Manual de Uso - Gerador de Relatórios CELESC

## ⚙️ Configuração da Planilha Base

O programa utiliza uma planilha de referência (`database.xlsx`) que deve estar na pasta:

"/base/database.xlsx"

### 🗂️ Estrutura obrigatória da planilha:

| UC        | Cod de Reg | Nome    |
| ----------|------------|---------|
| (Número)  | (Código)   | (Cidade)|

- **UC:** Número da Unidade Consumidora (**com todos os zeros à esquerda, sem pontos, barras ou traços**).
- **Cod de Reg:** Código do centro de custo, departamento ou setor.
- **Nome:** Nome da cidade ou unidade correspondente.

> ⚠️ **Importante:** As colunas devem ter exatamente estes nomes: `UC`, `Cod de Reg`, `Nome`.  
> A falta de qualquer uma delas impedirá a execução do programa.

---

## 💻 Como Executar o Programa

Após gerar o `.exe`, siga estes passos:

### 1. Abrir o Programa
Execute o arquivo `Relatorio.exe`.

### 2. Verificar a Planilha Base
**A planilha `database.xlsx` deve estar na pasta `base`.  
O executável `Relatorio.exe` deve estar na mesma pasta que a pasta `base`.**

Verifique no campo de status se foi carregada com sucesso.

### 3. Selecionar os PDFs
- Clique em **"Selecionar PDFs da Celesc"**.
- Selecione os arquivos PDF contendo as faturas que você deseja processar.

### 4. Definir Pasta de Saída
- Clique em **"Definir Pasta de Saída"**.
- Escolha onde deseja salvar o relatório Excel gerado.

### 5. Iniciar Processamento
- Clique em **"Iniciar Processamento de Relatório"**.
- Acompanhe o progresso na barra e no log em tempo real.

### 6. Finalização
- Ao final, o arquivo **`Relatorio_Celesc.xlsx`** será salvo na pasta de saída escolhida.
- O relatório será aberto automaticamente (se possível).

---

## 🛠️ Tecnologias Utilizadas

- Python
- Tkinter (interface gráfica)
- Pandas
- pdfplumber
- openpyxl (para manipulação avançada do Excel)
- Regex para extração de dados



