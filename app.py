import streamlit as st
import pandas as pd
import random
from datetime import datetime
import os

# 엑셀 파일 경로
EXCEL_FILE = os.path.join(os.path.dirname(__file__), 'barcode_database.xlsx')

# 세션 상태 초기화
if 'df' not in st.session_state:
    st.session_state.df = None
if 'current_barcode' not in st.session_state:
    st.session_state.current_barcode = None
if 'is_occupied' not in st.session_state:
    st.session_state.is_occupied = False

# 엑셀 파일 로드 또는 생성
def load_or_create_excel():
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE, dtype={'바코드': str})
            st.write("엑셀 파일이 성공적으로 로드되었습니다.")
            return df
        except Exception as e:
            st.error(f"엑셀 파일을 로드하는 중 오류가 발생했습니다: {e}")

    st.write("엑셀 파일을 찾을 수 없거나 로드할 수 없습니다. 새 파일을 생성합니다.")
    df = pd.DataFrame(columns=['바코드', '품번코드', '상품이름', '색상', '사이즈', '기재일', '수정일'])
    df['바코드'] = df['바코드'].astype(str)
    save_excel(df)
    return df

# 엑셀 파일 저장
def save_excel(df):
    try:
        df.to_excel(EXCEL_FILE, index=False)
        st.success("엑셀 파일이 성공적으로 저장되었습니다.")
    except Exception as e:
        st.error(f"엑셀 파일을 저장하는 중 오류가 발생했습니다: {e}")

# 바코드 생성 함수
def generate_barcode(df):
    prefix = "88061987"
    while True:
        suffix = ''.join([str(random.randint(0, 9)) for _ in range(5)])
        barcode = prefix + suffix
        if barcode not in df['바코드'].values:
            return barcode

# 메인 앱
def main():
    st.title('바코드 관리 시스템')

    # 다른 사용자가 접속 중인지 확인
    if st.session_state.is_occupied:
        st.error('현재 다른 사용자가 접속 중입니다. 잠시 후 다시 시도해 주세요.')
        return

    st.session_state.is_occupied = True

    # 데이터프레임 로드 (세션 상태 이용)
    if st.session_state.df is None:
        st.session_state.df = load_or_create_excel()

    df = st.session_state.df

    # 종료 버튼
    if st.button("종료"):
        st.session_state.is_occupied = False
        st.success('접속이 종료되었습니다. 다른 사용자가 접속할 수 있습니다.')
        return

    # 바코드 발급
    st.header('바코드 발급')

    # 바코드 생성 또는 재생성
    if st.button("바코드 생성/재생성"):
        st.session_state.current_barcode = generate_barcode(df)

    if st.session_state.current_barcode:
        st.write(f"현재 생성된 바코드: {st.session_state.current_barcode}")

    with st.form("barcode_issue"):
        품번코드 = st.text_input('품번코드')
        상품이름 = st.text_input('상품이름')
        색상 = st.text_input('색상')
        사이즈 = st.text_input('사이즈')
        현재날짜 = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        if st.form_submit_button("발급"):
            if st.session_state.current_barcode:
                new_data = pd.DataFrame({
                    '바코드': [st.session_state.current_barcode],
                    '품번코드': [품번코드],
                    '상품이름': [상품이름],
                    '색상': [색상],
                    '사이즈': [사이즈],
                    '기재일': [현재날짜],
                    '수정일': [현재날짜]
                })
                df = pd.concat([df, new_data], ignore_index=True)
                save_excel(df)
                st.session_state.df = df  # 세션 상태 업데이트
                st.success('바코드가 발급되었습니다.')
                st.session_state.current_barcode = None  # 바코드 초기화
            else:
                st.error("먼저 바코드를 생성해주세요.")

    # 바코드 수정
    st.header('바코드 수정')
    update_barcode = st.selectbox('수정할 바코드 선택', df['바코드'].tolist())

    # 선택된 바코드의 기존 정보를 불러오기
    selected_row = df[df['바코드'] == update_barcode]

    if not selected_row.empty:
        # 각 입력 필드에 선택된 바코드의 정보 자동 채우기
        update_품번코드 = st.text_input('새 품번코드', selected_row['품번코드'].values[0])
        update_상품이름 = st.text_input('새 상품이름', selected_row['상품이름'].values[0])
        update_색상 = st.text_input('새 색상', selected_row['색상'].values[0])
        update_사이즈 = st.text_input('새 사이즈', selected_row['사이즈'].values[0])

        # 수정 버튼
        if st.button("수정"):
            idx = df[df['바코드'] == update_barcode].index[0]
            df.loc[idx, '품번코드'] = update_품번코드 if update_품번코드 else df.loc[idx, '품번코드']
            df.loc[idx, '상품이름'] = update_상품이름 if update_상품이름 else df.loc[idx, '상품이름']
            df.loc[idx, '색상'] = update_색상 if update_색상 else df.loc[idx, '색상']
            df.loc[idx, '사이즈'] = update_사이즈 if update_사이즈 else df.loc[idx, '사이즈']
            df.loc[idx, '수정일'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            save_excel(df)
            st.session_state.df = df  # 세션 상태 업데이트
            st.success('바코드 정보가 수정되었습니다.')

    # 바코드 삭제
    st.header('바코드 삭제')
    with st.form("barcode_delete"):
        delete_barcode = st.selectbox('삭제할 바코드 선택', df['바코드'].tolist())

        if st.form_submit_button("삭제"):
            df = df[df['바코드'] != delete_barcode]
            save_excel(df)
            st.session_state.df = df  # 세션 상태 업데이트
            st.success('바코드가 삭제되었습니다.')

    # 데이터베이스 조회
    st.header('데이터베이스 조회')
    st.dataframe(df)

    # 엑셀 파일 새로고침
    if st.button('엑셀 파일 새로고침'):
        st.session_state.df = load_or_create_excel()
        st.success('엑셀 파일이 새로고침되었습니다.')

if __name__ == "__main__":
    try:
        main()
    finally:
        st.session_state.is_occupied = False
