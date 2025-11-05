import streamlit as st
import pandas as pd
import io
import tempfile
from pathlib import Path
import glob

st.title("Packing lists")
st.write(
    "Comece por carregar todos os ficheiros excel necessários (PL standard e summary):"
)

#from .... import (colocar as funções que estão no ficheiro com o código)
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

# Pequena ajuda visual
col1, col2 = st.columns(2)
with col1:
    st.caption("Standard")
    st.write([f.name for f in (standard_files or [])])
with col2:
    st.caption("Summary")
    st.write([f.name for f in (summary_files or [])])
