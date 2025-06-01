import streamlit as st
import pandas as pd
import io
import msoffcrypto
import re

st.set_page_config(page_title="📊 OCS 진료 분석기", layout="wide")
st.title("🦷 전체과 진료 요약 + 교수별 시간대 분석")

ocs_file = st.file_uploader("1️⃣ OCS 파일 업로드", type="xlsx")
ocs_password = st.text_input("🔐 OCS 파일 비밀번호 (있을 경우 입력)", type="password")

# 고정된 doctor_list 파일 로드
doctor_file_path = "doctor_list.xlsx"
doctor_excel = pd.ExcelFile(doctor_file_path)

시간순 = [9, 10, 11, 13, 14, 15, 16]

def classify_bozon_detail(text):
    text = str(text).lower()
    if any(k in text for k in ['endo', 'rct', 'c/f', 'post', 'core']):
        return 'Endo'
    elif any(k in text for k in ['resin', 'gi', 'cr', 'crown', 'zir', 'imp', 'occ', 'class']):
        return 'Operative'
    elif any(k in text for k in ['r/c', 'pano']):
        return '기타'
    else:
        return '기타'

def get_hour_flexible(time_str):
    time_str = str(time_str)
    match = re.search(r'(\d{1,2})[:시]', time_str)
    if match:
        return int(match.group(1))
    return None

def get_am_pm(hour):
    return '오전' if hour is not None and hour < 12 else '오후'

def detect_header_row(df):
    for i in range(min(10, len(df))):
        row = df.iloc[i].astype(str).tolist()
        if any("예약" in cell for cell in row):
            return i
    return None

# 시트명을 실제 과명과 매칭해주는 함수
def match_sheet_to_dept(sheet_name, dept_doctor_map):
    for dept in dept_doctor_map.keys():
        if dept in sheet_name:
            return dept
    return None

if ocs_file:
    try:
        if ocs_password:
            office_file = msoffcrypto.OfficeFile(ocs_file)
            office_file.load_key(password=ocs_password)
            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)
            ocs_excel = pd.ExcelFile(decrypted)
        else:
            ocs_excel = pd.ExcelFile(ocs_file)

        # doctor_list 처리
        dept_doctor_map = {}
        for sheet in doctor_excel.sheet_names:
            df = doctor_excel.parse(sheet)
            fr_list = df['FR'].dropna().astype(str).str.strip().tolist() if 'FR' in df.columns else []
            p_list = df['P'].dropna().astype(str).str.strip().tolist() if 'P' in df.columns else []
            dept_doctor_map[sheet.strip()] = {'FR': fr_list, 'P': p_list}

        all_records = []

        for sheet in ocs_excel.sheet_names:
            dept = match_sheet_to_dept(sheet, dept_doctor_map)
            if not dept:
                continue
            try:
                preview = ocs_excel.parse(sheet, nrows=10, header=None)
                header_row = detect_header_row(preview)
                if header_row is None:
                    continue

                df = ocs_excel.parse(sheet, skiprows=header_row)
                if '예약의사' not in df.columns or '예약시간' not in df.columns:
                    continue

                df = df[df['예약의사'].notna()]
                df['시'] = df['예약시간'].astype(str).apply(get_hour_flexible)
                df['시간대'] = df['시'].apply(get_am_pm)
                df = df[df['시'].isin(시간순)]

                df['보존내역'] = df['진료내역'].astype(str).apply(classify_bozon_detail) if '진료내역' in df.columns else '-'
                df['예약의사'] = df['예약의사'].astype(str).str.strip()

                df['구분'] = df['예약의사'].apply(lambda x:
                    'FR' if x in dept_doctor_map[dept]['FR'] else
                    ('P' if x in dept_doctor_map[dept]['P'] else 'FR'))

                for _, row in df.iterrows():
                    all_records.append({
                        '과명': dept,
                        '시': row['시'],
                        '시간대': row['시간대'],
                        '구분': row['구분'],
                        '보존내역': row['보존내역'],
                        '예약의사': row['예약의사']
                    })
            except Exception as e:
                st.warning(f"⚠️ 시트 {sheet} 오류: {e}")

        df_all = pd.DataFrame(all_records)

        # 이후 출력/집계 코드는 이전 코드와 동일하게 유지
        # (생략 가능 — 필요시 다시 이어서 붙여드림)

    except Exception as e:
        st.error(f"❌ 분석 중 오류 발생: {e}")
