import streamlit as st
import sys, os

# 패키지 루트(= dataly_manager의 부모 폴더)를 sys.path에 추가
APP_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(APP_DIR)
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

def render_home_ui():
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