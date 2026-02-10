"""
工事細目自動判定システム - Streamlit版（パスワード保護付き）
限定公開可能なWebアプリ
"""

import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import time
from datetime import datetime

# ページ設定
st.set_page_config(
    page_title="工事細目自動判定システム",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)
