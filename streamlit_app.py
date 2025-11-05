import streamlit as st
import pandas as pd
import io
import tempfile
from pathlib import Path
import glob

st.title("Packing lists")
st.write(
    "Comece por carregar todos os ficheiros excel necessários"
)
