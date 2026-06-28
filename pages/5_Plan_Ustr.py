import pandas as pd
import streamlit as st
import io
from datetime import datetime
from openpyxl.styles import Alignment, Font

def normalize_column_name(col):
    """Нормализует имя столбца или вкладки для защиты от ошибок раскладки (РУС/ENG) и пробелов"""
    if not isinstance(col, str):
        return col
    col = col.strip().upper().replace(" ", "").replace("_", "")
    replacements = {
        'K': 'К', 'M': 'М', 'P': 'Р',
