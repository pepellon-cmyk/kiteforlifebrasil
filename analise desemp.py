import streamlit as st
import pandas as pd
import os

# Tenta importar o openpyxl, se falhar, mostra uma mensagem amigável
try:
    from openpyxl import load_workbook
    BIBLIOTECA_OK = True
except ImportError:
    BIBLIOTECA_OK = False

st.set_page_config(page_title="App Responsivo Avaliação", layout="wide")

if not BIBLIOTECA_OK:
    st.error("⚠️ Erro de Configuração: A biblioteca 'openpyxl' não foi instalada.")
    st.info("Para corrigir, crie um ficheiro chamado **requirements.txt** no seu repositório com o texto: `openpyxl`")
    st.stop()

# --- Configurações do Ficheiro ---
NOME_FICHEIRO = "kite f lifeavaliacao_de_desempenho_-_2025.2.py.xlsm"

st.title("📱 App de Avaliação de Desempenho")

# Verifica se o ficheiro Excel existe na pasta
if os.path.exists(NOME_FICHEIRO):
    # Sidebar para navegação entre as sheets detetadas no seu ficheiro
    aba = st.sidebar.selectbox("Escolha a Sheet", ["sheet1", "sheet2", "sheet3", "sheet4"])
    
    # Leitura dos dados
    df = pd.read_excel(NOME_FICHEIRO, sheet_name=aba)
    
    st.subheader(f"Dados da {aba}")
    st.dataframe(df, use_container_width=True) # Torna a tabela responsiva

    # Formulário para adicionar novos dados
    with st.expander("➕ Adicionar Nova Avaliação"):
        with st.form("meu_formulario"):
            col1, col2 = st.columns(2)
            nome = col1.text_input("Nome")
            nota = col2.number_input("Nota", 0, 10)
            
            if st.form_submit_button("Guardar"):
                # Lógica para gravar sem corromper as macros do XLSM
                wb = load_workbook(NOME_FICHEIRO, keep_vba=True)
                ws = wb[aba]
                ws.append([nome, nota])
                wb.save(NOME_FICHEIRO)
                st.success("Gravado com sucesso!")
                st.rerun()
else:
    st.warning(f"Ficheiro {NOME_FICHEIRO} não encontrado no servidor.")