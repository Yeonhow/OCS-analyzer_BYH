import streamlit as st
import pandas as pd
import io
import msoffcrypto
import re

st.set_page_config(page_title="📊 OCS 진료 분석기", layout="wide")
st.title("🦷 전체과 진료 요약 + 교수별 시간대 분석")

ocs_file = st.file_uploader("1️⃣ OCS 파일 업로드", type="xlsx")
ocs_password = st.text_input("🔐 OCS 파일 비밀번호 (있을 경우 입력)", type="password")

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

# ✅ 과명을 시트명과 유연하게 매칭
def match_sheet_to_dept(sheet_name, dept_doctor_map):
    for dept in dept_doctor_map.keys():
        if dept in sheet_name or sheet_name in dept:
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

        st.subheader("📋 전체과 시간대별 진료 요약 (FR진료수(P진료수))")
        total_group = df_all.groupby(['시', '과명', '구분']).size().reset_index(name='진료수')
        pivot_fr = total_group[total_group['구분'] == 'FR'].pivot(index='시', columns='과명', values='진료수').fillna(0).astype(int).astype(str)
        pivot_p = total_group[total_group['구분'] == 'P'].pivot(index='시', columns='과명', values='진료수').fillna(0).astype(int).astype(str)
        pivot_fr, pivot_p = pivot_fr.align(pivot_p, join='outer', axis=1, fill_value='0')
        merged_total = pivot_fr + "(" + pivot_p + ")"

        # ✅ 가장 많은 과 표시
        styled = merged_total.copy()
        numeric_fr = pivot_fr.astype(int)
        numeric_p = pivot_p.astype(int)
        max_each_row = []
        for idx in styled.index:
            row_values = {}
            for col in styled.columns:
                fr_val = numeric_fr.loc[idx, col] if col in numeric_fr.columns else 0
                p_val = numeric_p.loc[idx, col] if col in numeric_p.columns else 0
                total_val = fr_val + p_val if col == '교정과' else fr_val
                row_values[col] = total_val
            max_col = max(row_values, key=row_values.get)
            max_each_row.append(max_col)

        for idx, max_col in zip(styled.index, max_each_row):
            styled.loc[idx, max_col] = f"✅ {styled.loc[idx, max_col]}"

        오전_fr = numeric_fr.loc[[9,10,11]].sum()
        오후_fr = numeric_fr.loc[[13,14,15,16]].sum()
        오전_p = numeric_p.loc[[9,10,11]].sum()
        오후_p = numeric_p.loc[[13,14,15,16]].sum()
        frp_summary = (오전_fr.astype(str) + "(" + 오전_p.astype(str) + ")").to_frame().T
        frp_summary = pd.concat([frp_summary,
                                 (오후_fr.astype(str) + "(" + 오후_p.astype(str) + ")").to_frame().T])
        frp_summary.index = ['오전 총합 FR(P)', '오후 총합 FR(P)']

        st.subheader("📋 전체과 오전/오후별 총진료수 (FR진료수(P진료수))")
        styled = styled.reindex(시간순).reset_index()
        st.dataframe(styled, use_container_width=True)
        st.dataframe(frp_summary, use_container_width=True)

        st.subheader("🦷 보존과 - Endo / Operative / 기타 (FR진료수(P진료수))")
        df_bozon = df_all[df_all['과명'] == '보존과']
        df_bozon = df_bozon[df_bozon['보존내역'].isin(['Endo', 'Operative', '기타'])]
        bozon_group = df_bozon.groupby(['시', '보존내역', '구분']).size().reset_index(name='진료수')
        bozon_fr = bozon_group[bozon_group['구분'] == 'FR'].pivot(index='시', columns='보존내역', values='진료수').fillna(0).astype(int).astype(str)
        bozon_p = bozon_group[bozon_group['구분'] == 'P'].pivot(index='시', columns='보존내역', values='진료수').fillna(0).astype(int).astype(str)
        bozon_fr = bozon_fr.reindex(시간순, fill_value='0')
        bozon_p = bozon_p.reindex(시간순, fill_value='0')
        bozon_fr, bozon_p = bozon_fr.align(bozon_p, join='outer', axis=1, fill_value='0')
        bozon_merged = bozon_fr + "(" + bozon_p + ")"
        bozon_merged = bozon_merged.fillna("0(0)").reset_index()
        st.dataframe(bozon_merged, use_container_width=True)

        st.subheader("👨‍⚕️ 교수별 오전/오후 진료 요약 (구강내과 · 보철과)")
        df_prof = df_all[(df_all['과명'].isin(['구강내과', '보철과'])) & (df_all['구분'] == 'P')]
        df_prof_summary = df_prof.pivot_table(
            index=['과명', '예약의사'], columns='시간대', values='구분', aggfunc='count', fill_value=0
        ).reset_index()
        st.dataframe(df_prof_summary, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            styled.to_excel(writer, index=False, sheet_name='전체과_시간대별')
            frp_summary.to_excel(writer, index=False, sheet_name='FRP_오전오후합계')
            bozon_merged.to_excel(writer, index=False, sheet_name='보존과_세부분류')
            df_prof_summary.to_excel(writer, index=False, sheet_name='P교수별_오전오후')
        output.seek(0)

        st.download_button(
            label="📥 분석 결과 엑셀 다운로드",
            data=output,
            file_name="OCS_진료분석결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ 분석 중 오류 발생: {e}")
