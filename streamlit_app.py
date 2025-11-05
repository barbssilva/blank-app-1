import streamlit as st
import pandas as pd
import io
import tempfile
from pathlib import Path
import glob

import os
import openpyxl
import copy
from openpyxl.utils import get_column_letter
import xlwings as xw

st.title("Packing Lists - BRAVE KID")

st.write("📁 Indique o nr da fatura/s:')

if "faturas" not in st.session_state:
    st.session_state.faturas = [""]

# Mostrar inputs dinâmicos
nova_lista = []
for i, valor in enumerate(st.session_state.faturas):
    novo_valor = st.text_input(f"Fatura {i+1}", valor, key=f"fatura_{i}").strip()
    nova_lista.append(novo_valor)

# Adiciona nova caixa se a última tiver valor
if nova_lista[-1] != "":
    nova_lista.append("")

st.session_state.faturas = nova_lista

# Cria a string final (ignorando vazios)
faturas_string = "_".join([f for f in st.session_state.faturas if f])

st.write(
    "Carregue todos os ficheiros excel necessários (PL standard e summary):"
)

from functions import join_excels, join_pls, remove_pls

standard_files = st.file_uploader(
    "Carregue as PLs standard",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    key="uploader_standard"
)

summary_files = st.file_uploader(
    "Carregue as PLs summary",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    key="uploader_summary"
)

# para visualizar os ficheiros que foram carregados
col1, col2 = st.columns(2)
with col1:
    st.caption("Standard")
    st.write([f.name for f in (standard_files or [])])
with col2:
    st.caption("Summary")
    st.write([f.name for f in (summary_files or [])])


if standard_files:
    standard_temp_paths = []  # aqui guardas o caminho de cada ficheiro temporário
    for f in standard_files:
        # cria um ficheiro temporário com a mesma extensão
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_excel:
            # guarda o conteúdo do ficheiro carregado
            temp_excel.write(f.read())
            # guarda o caminho
            standard_temp_paths.append(Path(temp_excel.name))
    #obter o diretorio do ficheiro temporário:
    temp_dir_standard = standard_temp_paths[0].parent
    output_file_standard = os.path.join(temp_dir_standard,'STANDARD_PL.xlsx')
    
        
if summary_files:
    summary_temp_paths = []  # aqui guardas o caminho de cada ficheiro temporário
    for f in summary_files:
        # cria um ficheiro temporário com a mesma extensão
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_excel:
            # guarda o conteúdo do ficheiro carregado
            temp_excel.write(f.read())
            # guarda o caminho
            summary_temp_paths.append(Path(temp_excel.name))
    #obter o diretorio do ficheiro temporário:
    temp_dir_summary = summary_temp_paths[0].parent
    output_file_summary = os.path.join(temp_dir_summary,'SUMMARY_PL.xlsx')
    last_file = os.path.join(temp_dir_summary,'Standard and Summary PACKING LIST'+ faturas_string +'.xlsx')

if summary_files and standard_files:
    placeholder = st.empty()
    placeholder.info("⏳ Por favor aguarde...")
        
    standard_pl=join_excels(standard_temp_paths,'standard', output_file_standard)
    summary_pl=join_excels(summary_temp_paths,'summary', output_file_summary)
        
    join_pls(summary_pl,standard_pl,last_file)
        
    remove_pls(standard_pl,summary_pl)
        
    placeholder.empty()
    st.success("Processo terminado!")
                
    # Abrir o ficheiro Excel processado para download
    with open(last_file, "rb") as f:
        st.download_button("Descarregar Excel Processado", f, file_name=os.path.basename(last_file))
        

    
