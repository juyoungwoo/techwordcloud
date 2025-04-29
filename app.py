import streamlit as st
import pandas as pd
from konlpy.tag import Okt
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from collections import Counter
import io

# 기본 설정
st.set_page_config(page_title="워드클라우드 생성기", layout="wide")

# 타이틀
st.title("📄 특허 정보 기반 워드클라우드 생성기")

# 파일 업로드
uploaded_file = st.file_uploader("특허 정보가 기재된 엑셀 파일을 업로드하세요(예, 특허명, 특허명세서, 특허 요약 등)", type=["xlsx"])

# 제거할 단어 입력
remove_words_input = st.text_input(
    "생성된 워드클라우드에서 추가적으로 제거할 단어를 입력하세요 (여러개 입력할 때는 쉼표로 구분)", ""
)

# 버튼 누를 때까지 기다리기
if uploaded_file:
    # 폰트 경로 설정
    try:
        font_path = './NanumGothic-Bold.ttf'  # 같은 폴더에 있을 경우
        fontprop = fm.FontProperties(fname=font_path)
        font_name = fontprop.get_name()
    except:
        st.warning("NanumGothic 폰트를 찾을 수 없습니다. 기본 폰트를 사용합니다.")
        font_name = 'sans-serif'
    
    plt.rcParams['font.family'] = font_name

    # 엑셀 읽기
    df = pd.read_excel(uploaded_file)

    # 숫자 열 제거
    text_columns = []
    for col in df.columns:
        sample = df[col].dropna().astype(str)
        if not all(sample.str.replace('.', '', 1).str.isnumeric()):
            text_columns.append(col)

    if not text_columns:
        st.error("텍스트가 포함된 열이 없습니다. 엑셀 파일을 확인해주세요.")
        st.stop()

    text_data = df[text_columns].fillna('').apply(lambda row: ' '.join(row), axis=1)

    # 텍스트 합치기
    full_text = " ".join(text_data.tolist())

    # 형태소 분석
    okt = Okt()
    tokens = okt.pos(full_text)

    # 기본 불용어
    basic_stopwords = {'상기', '포함', '발명', '장치', '여기', '주요', '방법', '기초', '정보', '활용', '개선', '효율', '시스템'}
    partial_stopwords = [
        '특징', '하나', '단계', '이용', '로부터', '사용', '따라서', '위해', '적용',
        '방법 제공', '통해', '대한', '복수', '각각', '구성', '향상', '기존', '통합', '해당', '실시', '다수'
    ]

    # 제거할 추가 단어 처리
    custom_stopwords = set(map(str.strip, remove_words_input.split(","))) if remove_words_input else set()

    def is_valid(word):
        return (
            len(word) > 1 and
            word not in basic_stopwords and
            word not in custom_stopwords and
            not any(stop in word for stop in partial_stopwords)
        )

    # 명사구 추출
    phrases = []
    current_phrase = []

    for word, tag in tokens:
        if tag == 'Noun' and is_valid(word):
        c    urrent_phrase.append(word)
        else:
            if len(current_phrase) >= 2:
                phrases.append(' '.join(current_phrase))
            current_phrase = []
    if len(current_phrase) >= 2:
        phrases.append(' '.join(current_phrase))
    
    # --- 여기 수정!!
    # 명사구 전체를 custom_stopwords와 비교해서 제거
    phrases = [p for p in phrases if p not in custom_stopwords]
    
    # 빈도수 계산
    counter = Counter(phrases)
    filtered_counter = {k: v for k, v in counter.items() if v >= 6}

    # 워드클라우드 생성
    wordcloud = WordCloud(
        font_path=font_path if 'Nanum' in font_name else None,
        width=1600,
        height=800,
        background_color='white',
        prefer_horizontal=1.0,
        relative_scaling=0.3,
        scale=2,
        margin=3,
        colormap='CMRmap'
    ).generate_from_frequencies(filtered_counter)

    # 워드클라우드 출력
    st.subheader("워드클라우드 결과")
    fig, ax = plt.subplots(figsize=(15, 8))
    ax.imshow(wordcloud, interpolation='bilinear')
    ax.axis('off')
    st.pyplot(fig)

    # 빈도수 테이블 출력
    freq_df = pd.DataFrame(filtered_counter.items(), columns=["단어", "빈도수"]).sort_values(by="빈도수", ascending=False)
    
    st.subheader("단어 빈도수 테이블")
    st.dataframe(freq_df)

    # 엑셀 다운로드
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        freq_df.to_excel(writer, index=False, sheet_name='Word Frequencies')
    output.seek(0)

    st.download_button(
        label="📥 빈도수 엑셀 다운로드",
        data=output,
        file_name='word_frequencies.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
