#datalyManager.py
import streamlit as st
import os
import sys

# (안전) 현재 디렉토리를 import 경로에 추가
APP_DIR = os.path.dirname(__file__)
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

from dataly_manager.ui.table_to_excel_ui import render_table_to_excel
from dataly_manager.ui.photo_to_excel_ui import render_photo_to_excel



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
    "🖼️ 사진 변환 (JSON→Excel)"
])

# 홈
with tabs[0]:
    st.header("🏠 홈")
    st.markdown("""
    왼쪽 탭에서 원하는 기능을 선택해 사용하세요.
    """)
    st.markdown("### 빠른 안내")
    st.markdown("""
    - **📊 표 변환**: `project_*.json` → 엑셀 변환 및 엑셀의 설명을 JSON에 반영.  
    - **🖼️ 사진 변환**: 사진용 `project_*.json` → 엑셀 변환 및 설명 반영.
    """)

    st.divider()
    st.markdown("### 빠른 점검")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.metric("Python", sys.version.split()[0])
    with col_b:
        st.metric("작업 디렉토리", os.path.basename(APP_DIR))
    with col_c:
        st.metric("캐시", "스트림릿 런타임")

    st.caption("상단 탭에서 기능을 선택해 주세요.")


# 표 변환 (JSON→Excel) — table_to_excel.py 사용
with tabs[1]:
    render_table_to_excel()

# 사진 변환 (JSON→Excel) — photo_to_excel.py 사용
with tabs[2]:
    render_photo_to_excel()

st.markdown("""
<hr>
<div class="footer">
문의: 검증 엔지니어 | Powered by Streamlit<br>
Copyright &copy; 2025. All rights reserved.
</div>
""", unsafe_allow_html=True)