import streamlit as st
import pdfplumber
import pandas as pd
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from io import BytesIO

# --------------------------------------------------
# 1. 고정 설정 (Streamlit Secrets 활용)
# --------------------------------------------------
GMAIL_USER = "jeongyunduk@gmail.com"
TARGET_EMAIL = "yunduk.jeong@sk.com"

try:
    GMAIL_APP_PASSWORD = st.secrets["GMAIL_PASSWORD"]
except Exception:
    GMAIL_APP_PASSWORD = None

# 페이지 설정
st.set_page_config(page_title="Asset Converter (시트 분리형)", layout="wide")

st.title("📄 자산현황 PDF 2 EXCEL 변환기 ")
st.markdown(f"PDF 파일을 업로드하면 **각 파일별로 시트를 만든 후** 하나의 엑셀 파일로 전송합니다.")

# --------------------------------------------------
# 2. PDF 처리 로직 (시트 분리형)
# --------------------------------------------------
uploaded_files = st.file_uploader("PDF 파일을 업로드하세요. (최대 5개)", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if len(uploaded_files) > 5:
        st.warning("최대 5개까지만 처리가능합니다. 상위 5개 파일만 진행합니다.")
        uploaded_files = uploaded_files[:5]

    all_sheets_data = {}
    total_count = 0

    with st.spinner('파일별로 데이터를 처리 중입니다...'):
        try:
            for uploaded_file in uploaded_files:
                # [수정] 파일명에서 ".pdf"와 "자산내역_"을 모두 제거
                sheet_name = uploaded_file.name.replace(".pdf", "").replace("자산내역_", "")[:30]
                
                with pdfplumber.open(uploaded_file) as pdf:
                    file_rows = []
                    for page in pdf.pages:
                        table = page.extract_table()
                        if table: file_rows.extend(table)
                
                processed_data = []
                for row in file_rows:
                    clean_row = [str(item).replace('\n', ' ').strip() for item in row if item is not None and str(item).strip() != ""]
                    asset_num = next((item for item in clean_row if item.replace('.', '').isdigit() and len(item.replace('.', '')) >= 10), None)
                    
                    if asset_num:
                        dates = [item for item in clean_row if re.match(r'\d{4}[.\s]\s?\d{2}[.\s]\s?\d{2}', item)]
                        numbers = [item for item in clean_row if item != asset_num and (',' in item or (item.isdigit() and len(item) < 10))]
                        texts = [item for item in clean_row if item not in dates and item not in numbers and item != asset_num]

                        entry = {
                            "매장명": texts[1] if len(texts) > 1 else (texts[0] if len(texts) > 0 else ""),
                            "자산번호": asset_num,
                            "자산명": texts[2] if len(texts) > 2 else (texts[1] if len(texts) > 1 else "정보없음"),
                            "취득일자": dates[0] if len(dates) > 0 else "",
                            "기준일자": dates[1] if len(dates) > 1 else "",
                            "취득가액": numbers[1] if len(numbers) > 1 else "0",
                            "장부가액": numbers[2] if len(numbers) > 2 else "0",
                            "변상금액": numbers[3] if len(numbers) > 3 else "0"
                        }
                        processed_data.append(entry)

                if processed_data:
                    df = pd.DataFrame(processed_data)
                    calc_cols = ['취득가액', '장부가액', '변상금액']
                    
                    # 수치 계산을 위해 숫자형으로 변환
                    for col in calc_cols:
                        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '').str.replace(' ', ''), errors='coerce').fillna(0)
                    
                    # 합계 계산
                    totals = {col: df[col].sum() for col in calc_cols}
                    summary_row = {c: "" for c in df.columns}
                    summary_row.update(totals)
                    summary_row['자산번호'] = '합계'
                    
                    final_df = pd.concat([df, pd.DataFrame([summary_row])], ignore_index=True)
                    
                    # [수정] 천 단위 콤마 표기 적용 (계산 완료 후 변환)
                    for col in calc_cols:
                        final_df[col] = final_df[col].apply(lambda x: format(int(x), ',') if str(x).replace('.0','').isdigit() else x)
                    
                    all_sheets_data[sheet_name] = final_df
                    total_count += len(df)

            if not all_sheets_data:
                st.error("추출된 데이터가 없습니다.")
            else:
                st.success(f"총 {len(all_sheets_data)}개의 시트가 준비되었습니다.")
                
                st.subheader("🔍 시트별 미리보기")
                tabs = st.tabs(list(all_sheets_data.keys()))
                for i, tab in enumerate(tabs):
                    sheet_key = list(all_sheets_data.keys())[i]
                    tab.dataframe(all_sheets_data[sheet_key], use_container_width=True)

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for sheet_name, df in all_sheets_data.items():
                        df.to_excel(writer, index=False, sheet_name=sheet_name)
                excel_data = output.getvalue()

                st.divider()

                st.subheader("📩 엑셀 파일 이메일 전송")
                if st.button(f"🚀 {TARGET_EMAIL}으로 전송"):
                    if GMAIL_APP_PASSWORD is None:
                        st.error("Secrets 설정에서 GMAIL_PASSWORD를 확인해주세요.")
                    else:
                        try:
                            msg = MIMEMultipart()
                            msg['From'] = GMAIL_USER
                            msg['To'] = TARGET_EMAIL
                            msg['Subject'] = f"[자동발송] 자산현황 보고서 (시트 {len(all_sheets_data)}개)"
                            
                            part = MIMEBase('application', "octet-stream")
                            part.set_payload(excel_data)
                            encoders.encode_base64(part)
                            part.add_header('Content-Disposition', 'attachment; filename="multi_sheet_report.xlsx"')
                            msg.attach(part)

                            server = smtplib.SMTP('smtp.gmail.com', 587)
                            server.starttls()
                            server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
                            server.sendmail(GMAIL_USER, TARGET_EMAIL, msg.as_string())
                            server.quit()
                            
                            st.success("메일 전송이 완료되었습니다!")
                        except Exception as e:
                            st.error(f"전송 오류: {e}")
                
        except Exception as e:
            st.error(f"오류 발생: {e}")