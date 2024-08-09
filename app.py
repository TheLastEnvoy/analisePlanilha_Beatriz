import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# Função para carregar as planilhas
def carregar_planilhas():
    st.sidebar.header("Carregar Planilhas")
    upload_planilha_ontem = st.sidebar.file_uploader("Selecionar Planilha de Ontem", type=["xlsx"])
    upload_planilha_hoje = st.sidebar.file_uploader("Selecionar Planilha de Hoje", type=["xlsx"])

    planilha_anterior = None
    planilha_atual = None

    if upload_planilha_ontem is not None:
        # Ler todas as abas e concatenar em um único DataFrame, ignorando a primeira aba
        xls_ontem = pd.read_excel(upload_planilha_ontem, sheet_name=None, engine='openpyxl')
        sheets_ontem = [df for name, df in xls_ontem.items() if name != "Document map"]
        planilha_anterior = pd.concat(sheets_ontem, ignore_index=True)

    if upload_planilha_hoje is not None:
        # Ler todas as abas e concatenar em um único DataFrame, ignorando a primeira aba
        xls_hoje = pd.read_excel(upload_planilha_hoje, sheet_name=None, engine='openpyxl')
        sheets_hoje = [df for name, df in xls_hoje.items() if name != "Document map"]
        planilha_atual = pd.concat(sheets_hoje, ignore_index=True)

    return planilha_anterior, planilha_atual

# Função para executar o código principal
def executar_codigo(planilha_anterior, planilha_atual):
    try:
        # Considerando que a coluna chave para identificar os lotes seja 'Lote'
        coluna_chave = 'Lote'

        # Verificar se a coluna 'Lote' existe em ambas as planilhas
        if coluna_chave not in planilha_anterior.columns or coluna_chave not in planilha_atual.columns:
            raise KeyError(f"A coluna '{coluna_chave}' não foi encontrada em uma ou ambas as planilhas.")

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

        # Salvar a planilha em um objeto binário
        from io import BytesIO
        output = BytesIO()
        workbook.save(output)
        workbook.close()
        output.seek(0)

        st.success(f"Foram identificados {len(novos_lotes)} novos lotes.")

        return output, nome_arquivo_saida

    except KeyError as e:
        st.error(str(e))
        return None, None
    except Exception as e:
        st.error(f"Ocorreu um erro ao executar o código:\n{str(e)}")
        return None, None

# Título do dashboard
st.title("Ferramenta para análise de novos lotes")

# Carregar planilhas
planilha_anterior, planilha_atual = carregar_planilhas()

# Botão para executar o código
if st.sidebar.button("Executar Código"):
    if planilha_anterior is not None and planilha_atual is not None:
        output, nome_arquivo_saida = executar_codigo(planilha_anterior, planilha_atual)
        if output is not None:
            st.download_button(
                label="Download da planilha gerada",
                data=output,
                file_name=nome_arquivo_saida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("Por favor, selecione ambas as planilhas antes de executar.")
