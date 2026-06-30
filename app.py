import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

# 1. 웹페이지 설정
st.set_page_config(page_title="체인로지스 정산 변환기", layout="centered")

st.title("📊 체인로지스 정산 데이터 변환 프로그램")
st.write("원본 엑셀 파일을 업로드하고 '변환 시작' 버튼을 클릭하세요.")

# 2. 파일 업로드 섹션
uploaded_file = st.file_uploader("변환할 엑셀 파일(xlsx)을 업로드해주세요.", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.success(f"파일이 업로드되었습니다! (파일명: {uploaded_file.name})")

    # 3. 변환 시작 버튼
    if st.button("변환 시작"):
        try:
            # 로딩 바 표시
            with st.spinner('데이터를 규칙에 맞게 가공하는 중입니다...'):

                # 엑셀 읽기
                df = pd.read_excel(uploaded_file)

                # --- [데이터 가공 로직 시작] ---

                # [추가 규칙] '픽업시간' 열이 비어있는(NaN) 행 전체 삭제
                if '픽업시간' in df.columns:
                    df = df.dropna(subset=['픽업시간'])
                    # 빈 문자열('')이나 공백만 있는 경우도 제거하기 위한 안전장치
                    df = df[df['픽업시간'].astype(str).str.strip() != '']

                # 계산을 위해 datetime 형식으로 임시 변환
                if '픽업시간' in df.columns:
                    df['픽업시간_dt'] = pd.to_datetime(df['픽업시간'], errors='coerce')
                if '완료시간' in df.columns:
                    df['완료시간_dt'] = pd.to_datetime(df['완료시간'], errors='coerce')

                # 규칙 1. '일자' 열 생성 (픽업시간 앞)
                if '픽업시간' in df.columns:
                    date_col = df['픽업시간_dt'].dt.strftime('%Y-%m-%d')
                    pickup_idx = df.columns.get_loc('픽업시간')
                    df.insert(pickup_idx, '일자', date_col)

                # 규칙 2. '소요시간' 열 생성 (완료시간 뒤)
                if '픽업시간' in df.columns and '완료시간' in df.columns:
                    duration_col = (df['완료시간_dt'] - df['픽업시간_dt']).dt.total_seconds() / 3600.0
                    duration_col = duration_col.round(2)  # 소수점 둘째자리 반올림
                    finish_idx = df.columns.get_loc('완료시간')
                    df.insert(finish_idx + 1, '소요시간', duration_col)

                # 규칙 6. '출고/회수' 열 생성 (발송인 뒤)
                if '발송인' in df.columns:
                    type_col = df['발송인'].apply(lambda x: '출고' if str(x).strip() == '신상마켓' else '회수')
                    sender_idx = df.columns.get_loc('발송인')
                    df.insert(sender_idx + 1, '출고/회수', type_col)

                # 규칙 3. '경유지수' 열 생성 (경유지개수 뒤)
                if '경유지개수' in df.columns:
                    stopover_col = pd.to_numeric(df['경유지개수'], errors='coerce').fillna(0).astype(int) + 1
                    stop_idx = df.columns.get_loc('경유지개수')
                    df.insert(stop_idx + 1, '경유지수', stopover_col)

                # 규칙 4. '대박스개수' 열 생성 (물품정보 뒤)
                if '물품정보' in df.columns:
                    def extract_box_count(text):
                        if pd.isna(text):
                            return 0
                        match = re.search(r'대박스\s*(\d+)', str(text))
                        if match:
                            return int(match.group(1))
                        return 0


                    box_col = df['물품정보'].apply(extract_box_count)
                    info_idx = df.columns.get_loc('물품정보')
                    df.insert(info_idx + 1, '대박스개수', box_col)

                # 규칙 5. '배송무게' 열 이름 변경 및 데이터 숫자화
                if '배송무게' in df.columns:
                    df['배송무게'] = df['배송무게'].astype(str).str.extract(r'(\d+\.?\d*)')
                    df['배송무게'] = pd.to_numeric(df['배송무게'], errors='coerce').fillna(0)
                    df = df.rename(columns={'배송무게': '배송무게(kg)'})

                # 임시 변수 제거
                df = df.drop(columns=['픽업시간_dt', '완료시간_dt'], errors='ignore')

                # --- [데이터 가공 로직 끝] ---

                # 4. 결과 파일을 메모리에 생성 (openpyxl 사용)
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                processed_data = output.getvalue()

                # --- [동적 파일명 생성 로직] ---
                input_filename = uploaded_file.name
                # 파일 이름에서 숫자+월 (예: 5월, 05월, 12월 등) 추출
                month_match = re.search(r'(\d+월)', input_filename)

                if month_match:
                    extracted_month = month_match.group(1)  # "5월" 추출
                    download_filename = f"체인로지스 정산_{extracted_month}.xlsx"
                else:
                    # 파일 이름에 '월'이 없는 경우 오늘 날짜를 기본값으로 사용
                    today_str = datetime.now().strftime('%Y%m%d')
                    download_filename = f"체인로지스 정산_{today_str}.xlsx"

                st.balloons()  # 성공 축하 풍선 효과
                st.success("변환 완료")

                # 5. 다운로드 버튼 표시 (동적으로 바뀐 파일명 적용)
                st.download_button(
                    label="📥 변환된 엑셀 파일 다운로드",
                    data=processed_data,
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"⚠️ 에러 발생: {e}")
            st.info("엑셀 파일의 필수 열 이름(픽업시간, 완료시간, 발송인 등)이 일치하는지 확인해주세요.")