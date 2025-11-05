import streamlit as st
import pandas as pd
import io
import tempfile
from pathlib import Path
import glob

st.set_page_config(page_title="Packing Lists", layout="wide")

st.title("Packing lists")
st.write(
    "Comece por carregar todos os ficheiros excel necessários"
)
