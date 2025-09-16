#datalyManager.py
import streamlit as st
import os
import sys

# 패키지 루트(= dataly_manager의 부모 폴더)를 sys.path에 추가
APP_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(APP_DIR)
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

# 절대 임포트로 통일
from dataly_manager.ui.table_to_excel_ui import render_table_to_excel
from dataly_manager.ui.photo_to_excel_ui import render_photo_to_excel
from dataly_manager.ui.home_ui import render_home_ui
from dataly_manager.ui.wsd_to_excel_ui import render_wsd_to_excel_ui

st.markdown("""
    <style>
    .main-title {font-size:2.1rem; font-weight:bold; color:#174B99; margin-bottom:0;}
    .sub-desc {font-size:1.1rem; color:#222;}
    .logo {height:60px; margin-bottom:15px;}
    .footer {color:#999; font-size:0.9rem; margin-top:40px;}
    div.stButton > button:first-child {background:#174B99; color:white; font-weight:bold; border-radius:8px;}
    .stTabs [data-baseweb="tab-list"] {background:#F6FAFD;}
    </style>
""", unsafe_allow_html=True)

col1, col2 = st.columns([0.15, 0.85])
with col1:
    st.image("https://static.streamlit.io/examples/dog.jpg", width=55)
with col2:
    st.markdown('<div class="main-title">Dataly Manager</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-desc">업무 자동화, 평가 데이터 변환, 관리, 수합 웹앱</div>', unsafe_allow_html=True)

tabs = st.tabs([
    "🏠 홈",
    "📊 표 변환 (JSON→Excel)",
    "🖼️ 사진 변환 (JSON→Excel)",
    "📄 SRL_ZA 변환 (JSON→Excel)"
])

# 홈 - home_ui.py 사용
with tabs[0]:
    render_home_ui()

# 표 변환 (JSON→Excel) — table_to_excel.py 사용
with tabs[1]:
    render_table_to_excel()

# 사진 변환 (JSON→Excel) — photo_to_excel.py 사용
with tabs[2]:
    render_photo_to_excel()

# ✅ WSD/DP/SRL/ZA → 엑셀 변환
with tabs[3]:
    render_wsd_to_excel_ui()

st.markdown("""
<hr>
<div class="footer">
문의: 검증 엔지니어 | Powered by Streamlit<br>
Copyright &copy; 2025. All rights reserved.
</div>
""", unsafe_allow_html=True)

# st.markdown("""
# <style>
# /* 2번째, 4번째 탭 버튼 숨김 */
# .stTabs [data-baseweb="tab-list"] button:nth-child(2),
# .stTabs [data-baseweb="tab-list"] button:nth-child(3) { display: none !important; }
# </style>
# """, unsafe_allow_html=True)