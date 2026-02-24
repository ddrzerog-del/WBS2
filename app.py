import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
import re

# --- 공통 로직: 텍스트에서 레벨 판단 ---
def get_level_from_text(text):
    # '1.1.1' 또는 '1-2-1' 같은 숫자 패턴을 찾아 점(.)의 개수로 레벨 판단
    match = re.match(r'^([\d\.\-]+)', str(text).strip())
    if match:
        code = match.group(1).strip('.')
        return code.count('.')
    return 0

# --- 엑셀 파일 처리 ---
def parse_excel(file):
    df = pd.read_excel(file)
    data = []
    # 첫 번째 컬럼에 WBS 번호가 있다고 가정하거나, 전체를 훑음
    for _, row in df.iterrows():
        text = str(row.iloc[0]) # 첫 번째 칸 기준
        if text.strip():
            level = get_level_from_text(text)
            data.append({'level': level, 'text': text})
    return data

# --- PPT 파일 처리 ---
def parse_ppt(file):
    prs = Presentation(file)
    data = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                level = get_level_from_text(shape.text)
                data.append({'level': level, 'text': shape.text})
    return data

# --- PPT WBS 생성 로직 (좌우 자동 정렬) ---
def create_wbs_ppt(wbs_data):
    prs = Presentation()
    prs.slide_width = Inches(13.33) # 16:9 비율
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    if not wbs_data: return prs

    # 레벨별로 그룹화
    levels_map = {}
    for item in wbs_data:
        lvl = item['level']
        if lvl not in levels_map: levels_map[lvl] = []
        levels_map[lvl].append(item)

    margin_x = Inches(0.5)
    content_width = prs.slide_width - (margin_x * 2)
    box_height = Inches(0.7)
    v_gap = Inches(0.4)

    # 레벨별 색상 테마
    colors = [RGBColor(44, 62, 80), RGBColor(52, 152, 219), RGBColor(46, 204, 113), RGBColor(155, 89, 182)]

    for lvl, items in levels_map.items():
        count = len(items)
        box_width = (content_width / count) - Inches(0.1)
        
        for i, item in enumerate(items):
            left = margin_x + (i * (content_width / count))
            top = Inches(1) + (lvl * (box_height + v_gap))
            
            shape = slide.shapes.add_shape(1, left, top, box_width, box_height)
            shape.fill.solid()
            shape.fill.fore_color.rgb = colors[lvl % len(colors)]
            shape.line.color.rgb = RGBColor(255, 255, 255)
            
            tf = shape.text_frame
            tf.text = item['text']
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            p = tf.paragraphs[0]
            p.font.size = Pt(10)
            p.font.color.rgb = RGBColor(255, 255, 255)
            p.font.bold = True

    return prs

# --- Streamlit UI ---
st.set_page_config(page_title="WBS 자동 정렬기", layout="wide")
st.title("📂 WBS 자동 생성 및 정렬 프로그램")
st.info("엑셀이나 PPT를 올리면 숫자 체계(1, 1.1 등)를 인식하여 깔끔한 WBS 슬라이드를 만들어줍니다.")

uploaded_file = st.file_uploader("파일 업로드 (xlsx, pptx)", type=["xlsx", "pptx"])

if uploaded_file:
    file_type = uploaded_file.name.split('.')[-1]
    wbs_items = []

    if file_type == "xlsx":
        wbs_items = parse_excel(uploaded_file)
    elif file_type == "pptx":
        wbs_items = parse_ppt(uploaded_file)

    if wbs_items:
        st.success(f"총 {len(wbs_items)}개의 항목을 인식했습니다.")
        
        # 데이터 미리보기
        with st.expander("인식된 데이터 보기"):
            st.table(pd.DataFrame(wbs_items))

        if st.button("🚀 PPT로 이쁘게 정렬하기"):
            out_prs = create_wbs_ppt(wbs_items)
            ppt_io = io.BytesIO()
            out_prs.save(ppt_io)
            ppt_io.seek(0)

            st.download_button(
                label="🎁 완성된 PPT 다운로드",
                data=ppt_io,
                file_name="Formatted_WBS.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
    else:
        st.error("항목을 인식하지 못했습니다. 숫자 체계(예: 1.1, 1.2)가 포함되어 있는지 확인하세요.")