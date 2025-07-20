import streamlit as st
import pandas as pd
import io
import datetime

st.set_page_config(page_title="건강검진 자가설문지 정리", layout="wide")

st.markdown(
    "<h1 style='text-align: center;'>🏥 건강검진 자가설문지 정리</h1>",
    unsafe_allow_html=True
)

st.write("엑셀 파일을 업로드하면, 줄바꿈 포함 정리 엑셀 파일을 다운받을 수 있습니다.")

uploaded_file = st.file_uploader("📤 엑셀 파일 업로드 (xlsx 형식)", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    try:
        info_cols = ['이름', '생년월일', '성별', '소속기관']
        survey_start_col = df.columns.get_loc("1. 심장혈관계 [1.1 고혈압]")
        survey_end_col = df.columns.get_loc("40-1. 파견 장소")
        survey_cols = df.columns[survey_start_col:survey_end_col]
        survey_df = df[survey_cols]

        results = []
        for i, row in survey_df.iterrows():
            info = df.loc[i, info_cols].to_dict()
            answered = row.notna().sum()
            not_answered = row.isna().sum()
            yes_cols = row[row == '예'].index.tolist()
            etc_cols = row[~row.isin(['예', '아니오']) & row.notna()]
            etc_info = list(zip(etc_cols.index, etc_cols.values))
            yes_count = (row == '예').sum()
            no_count = (row == '아니오').sum()
            etc_count = len(etc_info)

            combined = {
                **info,
                '총_답변수': answered,
                '미응답수': not_answered,
                "'아니오'_응답수": no_count,
                "'예'_응답수": yes_count,
                "'예'_응답_항목": '\n'.join(yes_cols),
                "'기타'_응답수": etc_count,
                "기타_응답": '\n'.join([f"{col} → {val}" for col, val in etc_info])
            }
            results.append(combined)

        summary_df = pd.DataFrame(results)
        summary_df.insert(0, '번호', range(1, len(summary_df) + 1))
        summary_df = summary_df[[ 
            '번호', '이름', '생년월일', '성별', '소속기관',
            '총_답변수', '미응답수', "'아니오'_응답수",
            "'예'_응답수", "'예'_응답_항목",
            "'기타'_응답수", "기타_응답"
        ]]

        st.success("✅ 분석이 완료되었습니다! 아래에서 요약 데이터를 확인하고 파일을 다운로드하세요.")
        st.markdown(f"<h5>👥 총 설문자 수: <span style='color:#0066cc'>{len(summary_df)}명</span></h5>", unsafe_allow_html=True)

        st.dataframe(summary_df, use_container_width=True)

        st.markdown("## 📋 응답 요약 보기")
        for idx, row in summary_df.iterrows():
            st.markdown(f"""
                <h4 style='margin-bottom:0.2em;'>🔹 {idx+1}. <span style="color:#333;">{row['이름']} ({row['생년월일']})</span></h4>
                <p style='margin-top:0; margin-bottom:0.5em;'>소속기관: <b>{row['소속기관']}</b></p>
                """, unsafe_allow_html=True)

            if row["'예'_응답_항목"]:
                st.markdown(f"""
                    <div style='background-color:#e6f4ea; padding:10px; border-radius:8px; margin-bottom:8px;'>
                    ✅ <b>'예' 응답 항목:</b><br>{row["'예'_응답_항목"].replace(chr(10), '<br>')}
                    </div>
                """, unsafe_allow_html=True)

            if row["기타_응답"]:
                st.markdown(f"""
                    <div style='background-color:#fdf3e6; padding:10px; border-radius:8px; margin-bottom:8px;'>
                    📝 <b>기타 응답:</b><br>{row['기타_응답'].replace(chr(10), '<br>')}
                    </div>
                """, unsafe_allow_html=True)

            st.markdown("<hr style='border:1px solid #ddd;'>", unsafe_allow_html=True)

        # 엑셀 다운로드
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            summary_df.to_excel(writer, index=False, sheet_name='응답요약')
            workbook = writer.book
            worksheet = writer.sheets['응답요약']
            wrap_format = workbook.add_format({'text_wrap': True})
            for col in ["'예'_응답_항목", "기타_응답"]:
                col_idx = summary_df.columns.get_loc(col)
                worksheet.set_column(col_idx, col_idx, 60, wrap_format)

        st.download_button(
            label="📥 엑셀 파일 다운로드",
            data=output.getvalue(),
            file_name=f"설문_요약_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"⚠️ 오류가 발생했습니다: {e}")
