#!/usr/bin/env python
# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
from decimal import Decimal
import datetime

def safe_decimal(x):
    """
    Tenta converter o valor `x` para Decimal.
    Se x for '-' ou vazio, retorna Decimal("0.0").
    Caso a conversão falhe, exibe uma mensagem de erro e retorna Decimal("0.0").
    """
    try:
        x_str = str(x).strip()
        if x_str in ['', '-']:
            return Decimal("0.0")
        return Decimal(x_str)
    except Exception as e:
        st.error(f"Error converting value {x} to Decimal. Defaulting to 0.0")
        return Decimal("0.0")

def transform_row(row, last, current_code):
    """
    Calcula os novos valores a partir da linha corrente e dos valores acumulados (last).
    Se o código (Code) da linha for diferente do anterior, reinicia os acumuladores.
    Caso contrário, realiza os cálculos para as movimentações:
      - "Desdobro"
      - "Transferência - Liquidação"
      - "Bonificação em Ativos"
      
    Em "Desdobro" e "Bonificação em Ativos", o custo da operação é zero e o preço médio é ajustado
    dividindo o custo total (old preço médio * old quantidade) pela nova quantidade.
    
    Retorna um dicionário com os novos valores e o código atual.
    """
    if current_code != row["Code"]:
        current_code = row["Code"]
        last = {}
        new_row = {
            "Entry_Exit": row["Entry_Exit"],
            "Date": row["Date"],
            "Code": row["Code"],
            "Product": row["Product"],
            "Quantity": row["Quantity"],
            "Unit_Price": row["Unit_Price"],
            "Operation_Value": row["Operation_Value"],
            "Last_QTD_PM": 0,
            "Last_PM": Decimal("0.0"),
            "Last_OP_Value": Decimal("0.0"),
            "Current_QTD_PM": row["Quantity"],
            "Current_PM": row["Unit_Price"],
            "Current_OP_Value": row["Operation_Value"]
        }
    else:
        if row["Movement"] == "Desdobro":
            current_qty = row["Quantity"] + last["Current_QTD_PM"]
            if last["Current_QTD_PM"] != 0:
                current_pm = Decimal(last["Current_PM"]) / (Decimal(current_qty) / Decimal(last["Current_QTD_PM"]))
            else:
                current_pm = Decimal(0)
            current_op_value = current_pm * Decimal(current_qty)
            new_row = {
                "Entry_Exit": row["Entry_Exit"],
                "Date": row["Date"],
                "Code": row["Code"],
                "Product": row["Product"],
                "Quantity": row["Quantity"],
                "Unit_Price": Decimal("0"),
                "Operation_Value": Decimal("0"),
                "Last_QTD_PM": last["Current_QTD_PM"],
                "Last_PM": Decimal(last["Current_PM"]),
                "Last_OP_Value": Decimal(last["Current_OP_Value"]),
                "Current_QTD_PM": current_qty,
                "Current_PM": current_pm,
                "Current_OP_Value": current_op_value
            }
        elif row["Movement"] == "Transferência - Liquidação":
            if row["Entry_Exit"] == 1:
                current_qty = row["Quantity"] + last["Current_QTD_PM"]
                denominator = last["Current_QTD_PM"] + row["Quantity"]
                if denominator != 0:
                    current_pm = (row["Operation_Value"] + last["Current_OP_Value"]) / Decimal(denominator)
                else:
                    current_pm = Decimal(0)
                current_op_value = current_pm * Decimal(current_qty)
                new_row = {
                    "Entry_Exit": row["Entry_Exit"],
                    "Date": row["Date"],
                    "Code": row["Code"],
                    "Product": row["Product"],
                    "Quantity": row["Quantity"],
                    "Unit_Price": row["Unit_Price"],
                    "Operation_Value": row["Operation_Value"],
                    "Last_QTD_PM": last["Current_QTD_PM"],
                    "Last_PM": last["Current_PM"],
                    "Last_OP_Value": last["Current_OP_Value"],
                    "Current_QTD_PM": current_qty,
                    "Current_PM": current_pm,
                    "Current_OP_Value": current_op_value
                }
            elif row["Entry_Exit"] == -1:
                current_qty = row["Quantity"] - last["Current_QTD_PM"]
                current_pm = last["Current_PM"] if current_qty != 0 else Decimal(0)
                current_op_value = current_pm * Decimal(current_qty)
                new_row = {
                    "Entry_Exit": row["Entry_Exit"],
                    "Date": row["Date"],
                    "Code": row["Code"],
                    "Product": row["Product"],
                    "Quantity": row["Quantity"],
                    "Unit_Price": row["Unit_Price"],
                    "Operation_Value": row["Operation_Value"],
                    "Last_QTD_PM": last["Current_QTD_PM"],
                    "Last_PM": last["Current_PM"],
                    "Last_OP_Value": last["Current_OP_Value"],
                    "Current_QTD_PM": current_qty,
                    "Current_PM": Decimal(current_pm),
                    "Current_OP_Value": current_op_value
                }
            else:
                st.error(f"Movement not mapped: {row}")
                new_row = None
        elif row["Movement"] == "Bonificação em Ativos":
            current_qty = row["Quantity"] + last["Current_QTD_PM"]
            if last["Current_QTD_PM"] != 0:
                current_pm = Decimal(last["Current_PM"]) / (Decimal(current_qty) / Decimal(last["Current_QTD_PM"]))
            else:
                current_pm = Decimal(0)
            current_op_value = current_pm * Decimal(current_qty)
            new_row = {
                "Entry_Exit": row["Entry_Exit"],
                "Date": row["Date"],
                "Code": row["Code"],
                "Product": row["Product"],
                "Quantity": row["Quantity"],
                "Unit_Price": Decimal("0"),
                "Operation_Value": Decimal("0"),
                "Last_QTD_PM": last["Current_QTD_PM"],
                "Last_PM": Decimal(last["Current_PM"]),
                "Last_OP_Value": Decimal(last["Current_OP_Value"]),
                "Current_QTD_PM": current_qty,
                "Current_PM": current_pm,
                "Current_OP_Value": current_op_value
            }
        else:
            st.error(f"Movement not mapped: {row}")
            new_row = None
    return new_row, current_code

def preprocess_excel(excel_file):
    """
    Lê o arquivo Excel, realiza as renomeações necessárias,
    converte os tipos e cria a coluna 'Code'.
    """
    try:
        df_excel = pd.read_excel(excel_file, sheet_name='Movimentação')
    except Exception as e:
        st.error("Erro ao ler o arquivo Excel. Verifique se o arquivo está correto e se o sheet 'Movimentação' existe.")
        st.error(e)
        st.stop()
    
    df_excel = df_excel.rename(columns={
        "Entrada/Saída": "Entry_Exit",
        "Data": "Date",
        "Movimentação": "Movement",
        "Produto": "Product",
        "Instituição": "Institution",
        "Quantidade": "Quantity",
        "Preço unitário": "Unit_Price",
        "Valor da Operação": "Operation_Value"
    })
    
    df_excel["Date"] = pd.to_datetime(df_excel["Date"], format='%d/%m/%Y')
    df_excel["Entry_Exit"] = df_excel["Entry_Exit"].map({"Credito": 1, "Debito": -1}).fillna(0).astype(int)
    df_excel["Code"] = df_excel["Product"].str.split("-").str[0].str.strip()
    df_excel["Quantity"] = df_excel["Quantity"].astype(int)
    df_excel["Unit_Price"] = df_excel["Unit_Price"].apply(safe_decimal)
    df_excel["Operation_Value"] = df_excel["Operation_Value"].apply(safe_decimal)
    return df_excel

def check_and_update_codigo(df, ANO_BASE):
    """
    Verifica se há movimentações de "Atualização". Para cada movimentação de "Atualização",
    se existir também uma movimentação (com "Transferência") com a mesma quantidade, propõe a atualização
    do código, considerando que o código do movimento "Atualização" é o novo.
    
    Apenas os códigos para os quais o usuário responder "Sim" serão atualizados.
    Em caso de atualização de código for "Sim", atualiza também a quantidade (ação).
    Se houver propostas de atualização, o cálculo final é realizado automaticamente após a confirmação.
    """
    proposals = []
    df_atualizacao = df[df["Movement"] == "Atualização"]
    for idx, row_atual in df_atualizacao.iterrows():
        matching_transfer = df[
            (df["Movement"].str.contains("Transferência", na=False)) &
            (df["Quantity"] == row_atual["Quantity"])
        ]
        if not matching_transfer.empty:
            transfer_index = matching_transfer.index[0]
            old_code = df.loc[transfer_index, "Code"]
            new_code = row_atual["Code"]  # Código novo
            proposals.append((idx, old_code, new_code, row_atual["Quantity"], transfer_index))
    
    if proposals:
        update_form_container = st.empty()
        with update_form_container.form(key="update_form"):
            st.write("Foram encontradas atualizações de código. Por favor, confirme cada atualização abaixo:")
            responses = {}
            for idx, old_code, new_code, qty, _ in proposals:
                responses[idx] = st.radio(
                    f"Atualizar código de {old_code} para {new_code} (Quantidade: {qty})?",
                    options=["Sim", "Não"],
                    key=f"update_{idx}"
                )
            submitted = st.form_submit_button("Confirmar Atualizações")
        
        if not submitted:
            st.info("Aguardando a confirmação das atualizações...")
            st.stop()
        else:
            update_form_container.empty()
            for idx, old_code, new_code, qty, transfer_index in proposals:
                if responses[idx] == "Sim":
                    df.at[transfer_index, "Code"] = new_code
                    df.at[transfer_index, "Quantity"] = qty
                    st.success(f"Código {old_code} atualizado para {new_code} (Quantidade: {qty}).")
                else:
                    st.info(f"Código {old_code} não atualizado (Quantidade: {qty}).")
            st.info("Processamento das atualizações concluído. Calculando resultados finais...")
            calculate_results(df, ANO_BASE)
            st.stop()
    return df

def calculate_results(df, ANO_BASE):
    """
    A partir do DataFrame pré-processado (e atualizado, se aplicável), filtra as movimentações pertinentes,
    aplica as transformações e calcula os indicadores finais, exibindo os resultados para FII e Ações.
    """
    ANO_ANTERIOR = ANO_BASE - 1
    df_calc = df[df["Movement"].isin(["Transferência - Liquidação", "Desdobro", "Bonificação em Ativos"])]
    df_calc = df_calc.sort_values(by=["Code", "Date"]).reset_index(drop=True)
    
    data = []
    last = {}
    current_code = ''
    for _, row in df_calc.iterrows():
        new_row, current_code = transform_row(row, last, current_code)
        if new_row is not None:
            data.append(new_row)
            last = new_row
    df_pm = pd.DataFrame(data)
    df_pm["Year"] = df_pm["Date"].dt.year
    
    df_until_base = df_pm[df_pm["Date"] <= pd.to_datetime(f"{ANO_BASE}-12-31")]
    df_until_anterior = df_pm[df_pm["Date"] <= pd.to_datetime(f"{ANO_ANTERIOR}-12-31")]
    
    df_ano_base = df_until_base.sort_values(by="Date", ascending=False).drop_duplicates(subset=["Code"])
    df_ano_anterior = df_until_anterior.sort_values(by="Date", ascending=False).drop_duplicates(subset=["Code"])
    
    df_ano_base = df_ano_base.rename(columns={
        "Current_QTD_PM": f"Quantidade_em_{ANO_BASE}",
        "Current_OP_Value": f"Situacao_em_{ANO_BASE}"
    })
    df_ano_anterior = df_ano_anterior.rename(columns={
        "Current_QTD_PM": f"Quantidade_em_{ANO_ANTERIOR}",
        "Current_OP_Value": f"Situacao_em_{ANO_ANTERIOR}"
    })
    
    df_final = pd.merge(df_ano_anterior, df_ano_base, on="Code", how="outer")
    df_final = df_final[[
        "Code",
        f"Quantidade_em_{ANO_ANTERIOR}", f"Situacao_em_{ANO_ANTERIOR}",
        f"Quantidade_em_{ANO_BASE}", f"Situacao_em_{ANO_BASE}"
    ]]
    
    df_final[f"Quantidade_em_{ANO_ANTERIOR}"] = df_final[f"Quantidade_em_{ANO_ANTERIOR}"].fillna(0)
    df_final[f"Situacao_em_{ANO_ANTERIOR}"] = df_final[f"Situacao_em_{ANO_ANTERIOR}"].fillna(Decimal("0.0"))
    df_final[f"Quantidade_em_{ANO_BASE}"] = df_final[f"Quantidade_em_{ANO_BASE}"].fillna(0)
    df_final[f"Situacao_em_{ANO_BASE}"] = df_final[f"Situacao_em_{ANO_BASE}"].fillna(Decimal("0.0"))
    
    df_final[f"Situacao_em_{ANO_ANTERIOR}"] = df_final[f"Situacao_em_{ANO_ANTERIOR}"].apply(
        lambda x: format(x, '.2f').replace('.', ',') if isinstance(x, Decimal) else format(Decimal(x), '.2f').replace('.', ',')
    )
    df_final[f"Situacao_em_{ANO_BASE}"] = df_final[f"Situacao_em_{ANO_BASE}"].apply(
        lambda x: format(x, '.2f').replace('.', ',') if isinstance(x, Decimal) else format(Decimal(x), '.2f').replace('.', ',')
    )
    
    df_fii = pd.read_csv('./data/CNPJs_FII.csv', dtype={'FII_CNPJ': str})
    df_acoes = pd.read_csv('./data/CNPJs_ACOES.csv', dtype={'STOCK_CNPJ': str})
    
    df_fii_result = pd.merge(df_final, df_fii, left_on="Code", right_on="FII_TICKER", how="inner")
    df_fii_result = df_fii_result.rename(columns={"FII_CNPJ": "CNPJ"})
    df_fii_result["Discriminação"] = "Cotas de " + df_fii_result["Code"] + " distribuídas na corretora NuInvest, CNPJ 62.169.875/0001-79."
    df_fii_result = df_fii_result[[
        "Code",
        f"Quantidade_em_{ANO_ANTERIOR}", f"Situacao_em_{ANO_ANTERIOR}",
        f"Quantidade_em_{ANO_BASE}", f"Situacao_em_{ANO_BASE}",
        "CNPJ", "Discriminação"
    ]]
    
    df_acoes_result = pd.merge(df_final, df_acoes, left_on="Code", right_on="STOCK_TICKER", how="inner")
    df_acoes_result = df_acoes_result.rename(columns={"STOCK_CNPJ": "CNPJ"})
    df_acoes_result["Discriminação"] = (
        df_acoes_result[f"Quantidade_em_{ANO_BASE}"].astype(str) +
        " ações emitidas pela empresa " +
        df_acoes_result["STOCK_NAME"] + " " +
        df_acoes_result["CNPJ"]
    )
    df_acoes_result = df_acoes_result[[
        "Code",
        f"Quantidade_em_{ANO_ANTERIOR}", f"Situacao_em_{ANO_ANTERIOR}",
        f"Quantidade_em_{ANO_BASE}", f"Situacao_em_{ANO_BASE}",
        "CNPJ", "Discriminação"
    ]]
    

    df_fii_result['Sorting_Key'] = df_fii_result[f"Situacao_em_{ANO_BASE}"].apply(lambda x: float(x.replace(',', '.')))
    df_fii_result = df_fii_result.sort_values(by='Sorting_Key', ascending=False).drop(columns='Sorting_Key')
    
    df_acoes_result['Sorting_Key'] = df_acoes_result[f"Situacao_em_{ANO_BASE}"].apply(lambda x: float(x.replace(',', '.')))
    df_acoes_result = df_acoes_result.sort_values(by='Sorting_Key', ascending=False).drop(columns='Sorting_Key')
    
    st.subheader("Resultado para FII")
    st.dataframe(df_fii_result)
    st.subheader("Resultado para Ações")
    st.dataframe(df_acoes_result)

def main():
    st.title("Calculadora de Preço Médio")
    
    st.info("""
    Para obter o arquivo de movimentação:
    1. Acesse https://www.investidor.b3.com.br/extrato/movimentacao
    2. Faça login na sua conta
    3. Selecione o período máximo possível 
    4. Clique em "Baixar" para baixar o arquivo Excel
    5. Faça upload do arquivo baixado abaixo
    """)

    uploaded_xlsx = st.file_uploader("Faça o upload do arquivo Excel (.xlsx)", type=["xlsx"])
    ANO_BASE = st.number_input("Insira o Ano Base", min_value=1900, max_value=2100, value=datetime.datetime.now().year -1, step=1)
    
    if uploaded_xlsx is not None:
        df = preprocess_excel(uploaded_xlsx)
        df = check_and_update_codigo(df, ANO_BASE)
        calculate_results(df, ANO_BASE)

if __name__ == '__main__':
    main() 
