# Gerador de Relatório de Faturas Celesc

Aplicação desktop desenvolvida em Python para automação de processos de auditoria de contas de energia. O sistema extrai dados de faturas da Celesc (PDF), cruza informações com uma base de dados interna e gera um relatório consolidado em Excel contendo valores líquidos, brutos e retenções tributárias.

## Funcionalidades

- **Extração em Lote:** Processamento de múltiplos arquivos PDF simultaneamente.
- **Captura de Dados:** Leitura de Unidade Consumidora (UC), valores monetários e impostos retidos (IRPJ, PIS, COFINS, CSLL).
- **Validação de Base:** Verificação automática da existência da UC na planilha de controle (`database.xlsx`).
- **Cálculo Reverso:** Geração do Valor Bruto com base no Líquido + Descontos.
- **Relatório de Erros:** Aba dedicada no Excel para apontar faturas ilegíveis ou UCs não cadastradas.
- **Interface Gráfica:** GUI com logs de processamento em tempo real.

## Estrutura de Arquivos Necessária

Para que o executável funcione corretamente, a seguinte estrutura de pastas deve ser mantida:

```text
📂 Pasta do Projeto
├── 📄 Relatorio.exe
└── 📂 base
    └── 📄 database.xlsx
```

## Configuração da Base de Dados

O arquivo `database.xlsx` (localizado dentro da pasta `base`) é obrigatório. Ele serve como referência para cruzar o número da UC com o centro de custo e o nome da unidade.

**Estrutura obrigatória das colunas:**

| UC | Cod de Reg | Nome |
| :--- | :--- | :--- |
| (Número da UC) | (Código do Centro de Custo) | (Cidade/Unidade) |

**Importante:**
1. A coluna `UC` deve conter apenas números (sem pontos ou traços).
2. Os nomes dos cabeçalhos devem ser exatamente: **UC**, **Cod de Reg**, **Nome**.

## Como Utilizar

1. Certifique-se de que o arquivo `database.xlsx` está atualizado na pasta `base`.
2. Execute o arquivo `Relatorio.exe`.
3. Na interface:
   - Clique em **Selecionar PDFs** e escolha os arquivos de fatura.
   - Clique em **Definir Pasta de Saída** para escolher onde salvar o Excel final.
   - Clique em **Iniciar Processamento**.
4. O sistema irá gerar o arquivo `Relatorio_Celesc.xlsx` contendo 2 ou 3 abas dependendo das opções marcadas:
   - `Relatorio_Dados_Extraidos`: Dados processados com sucesso.
   - `Relatorio_Erros`: Arquivos que falharam ou UCs não encontradas na base.
