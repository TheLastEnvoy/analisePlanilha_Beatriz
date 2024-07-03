import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import os

# Variáveis globais para armazenar os dados das planilhas
planilha_anterior = None
planilha_atual = None

# Função para carregar as planilhas
def carregar_planilhas():
    global planilha_anterior, planilha_atual

    st.sidebar.header("Carregar Planilhas")
    upload_planilha_ontem = st.sidebar.file_uploader("Selecionar Planilha de Ontem", type=["xlsx"])
    upload_planilha_hoje = st.sidebar.file_uploader("Selecionar Planilha de Hoje", type=["xlsx"])

    if upload_planilha_ontem is not None:
        planilha_anterior = pd.read_excel(upload_planilha_ontem, engine='openpyxl')

    if upload_planilha_hoje is not None:
        planilha_atual = pd.read_excel(upload_planilha_hoje, engine='openpyxl')

# Função para executar o código principal
def executar_codigo():
    global planilha_anterior, planilha_atual

    if planilha_anterior is None or planilha_atual is None:
        st.warning("Por favor, selecione ambas as planilhas antes de executar.")
        return

    try:
        # Considerando que a coluna chave para identificar os lotes seja 'Lote'
        coluna_chave = 'Lote'

        # Encontrar os lotes novos que estão na planilha atual, mas não estavam na planilha anterior
        novos_lotes = planilha_atual[~planilha_atual[coluna_chave].isin(planilha_anterior[coluna_chave])]

        # Criar uma nova planilha Excel
        nome_arquivo_saida = 'novos_lotes_identificados.xlsx'
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'Planilha'

        # Adicionar cabeçalho da planilha
        header = list(planilha_atual.columns)
        sheet.append(header)

        # Copiar dados da planilha_atual para a nova planilha
        for r in dataframe_to_rows(planilha_atual, index=False, header=False):
            sheet.append(r)

        # Aplicar formatação às linhas dos novos itens (linhas marcadas como novos)
        gray_fill = PatternFill(start_color='C0C0C0', end_color='C0C0C0', fill_type='solid')
        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, min_col=1, max_col=sheet.max_column), start=2):
            if planilha_atual.loc[row_idx - 2, 'Lote'] in novos_lotes['Lote'].values:
                for cell in row:
                    cell.fill = gray_fill

        # Salvar a planilha
        workbook.save(filename=nome_arquivo_saida)
        workbook.close()

        st.success(f"Foram identificados {len(novos_lotes)} novos lotes.\nPlanilha gerada: {nome_arquivo_saida}")

        # Permitir download do arquivo gerado
        with open(nome_arquivo_saida, "rb") as file:
            btn = st.download_button(
                label="Download da planilha gerada",
                data=file,
                file_name=nome_arquivo_saida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Ocorreu um erro ao executar o código:\n{str(e)}")

# Título do dashboard
st.title("Identificação de Novos Lotes")

# Carregar planilhas
carregar_planilhas()

# Botão para executar o código
if st.sidebar.button("Executar Código"):
    executar_codigo()
