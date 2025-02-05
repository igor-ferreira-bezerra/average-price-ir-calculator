# Average Price IR Calculator

A **Calculadora de Preço Médio** é uma aplicação web desenvolvida em Python com [Streamlit](https://streamlit.io/) que permite o cálculo do preço médio de ativos (ações e FIIs) a partir de um arquivo Excel contendo as movimentações realizadas na corretora. A ferramenta processa movimentações como *Transferência - Liquidação*, *Desdobro* e *Bonificação em Ativos*, além de propor atualizações de código quando identificadas movimentações de "Atualização".

## Funcionalidades

- **Leitura e Pré-processamento do Excel:**  
  Faz a leitura do arquivo Excel (aba *Movimentação*) e realiza o mapeamento e conversão dos dados para os tipos corretos (datas, inteiros e decimais).

- **Cálculo de Preço Médio:**  
  Processa as movimentações para calcular os indicadores finais, como quantidade e valor (situação) de ativos até os anos-base e anterior.

- **Tratamento de Movimentações Especiais:**  
  Lida com operações de desdobro, bonificação e transferência; em casos de "Atualização" de código, propõe a adequação dos dados com a confirmação do usuário.

- **Integração com Dados Externos:**  
  Realiza merge com arquivos CSV (`data/CNPJs_FII.csv` e `data/CNPJs_ACOES.csv`) para extrair informações adicionais dos CNPJs dos ativos.

- **Interface Intuitiva:**  
  Utiliza a interface do Streamlit para facilitar a interação do usuário, apresentando formulários para upload de arquivos, entrada do ano-base e confirmação de atualizações.

## Pré-requisitos

- **Python 3.7+**
- Bibliotecas:
  - [Streamlit](https://streamlit.io/)
  - [Pandas](https://pandas.pydata.org/)
  - [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/) (para leitura de arquivos Excel)

## Instalação

1. **Clone o repositório:**

   ```bash
   git clone <URL_DO_REPO>
   cd average-price-ir-calculator
   ```

2. **Crie e ative um ambiente virtual (opcional, mas recomendado):**

   - **No Windows:**
     ```bash
     python -m venv venv
     venv\Scripts\activate
     ```
   - **No Linux/Mac:**
     ```bash
     python3 -m venv venv
     source venv/bin/activate
     ```

3. **Instale as dependências:**

   ```bash
   pip install streamlit pandas openpyxl
   ```

## Como Usar

1. **Inicie a aplicação com Streamlit:**

   ```bash
   streamlit run calculator.py
   ```

2. **Utilize a interface web:**
   - **Upload do Arquivo Excel:**  
     Selecione o arquivo Excel que contenha a aba "Movimentação" com seus dados de movimentação.
   - **Insira o Ano Base:**  
     Defina o ano base para os cálculos (ex: ano anterior ao atual).
   - **Confirmação de Atualizações:**  
     Caso haja movimentações de "Atualização", a aplicação solicitará a confirmação via formulário. Se aprovado, os códigos serão atualizados automaticamente.
   - **Visualização dos Resultados:**  
     A calculadora exibirá duas tabelas:
     - Resultados para FIIs.
     - Resultados para Ações.

## Estrutura do Projeto