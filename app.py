import pandas as pd

def carregar_planilhas():
    st.sidebar.header("Carregar Planilhas")
    upload_planilha_ontem = st.sidebar.file_uploader("Selecionar Planilha de Ontem", type=["xlsx"])
    upload_planilha_hoje = st.sidebar.file_uploader("Selecionar Planilha de Hoje", type=["xlsx"])

    planilha_anterior = None
    planilha_atual = None

    if upload_planilha_ontem is not None:
        planilha_anterior = ler_todas_abas(upload_planilha_ontem)

    if upload_planilha_hoje is not None:
        planilha_atual = ler_todas_abas(upload_planilha_hoje)

    return planilha_anterior, planilha_atual

def ler_todas_abas(arquivo):
    # Lê todas as abas da planilha
    xls = pd.ExcelFile(arquivo)
    dfs = []
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        # Remove linhas vazias e cabeçalhos desnecessários
        df = df.dropna(how='all').reset_index(drop=True)
        # Assume que a primeira linha não vazia é o cabeçalho real
        header_row = df.first_valid_index()
        if header_row is not None:
            df.columns = df.iloc[header_row]
            df = df.iloc[header_row + 1:].reset_index(drop=True)
        dfs.append(df)

    # Concatena todos os DataFrames
    return pd.concat(dfs, ignore_index=True)
