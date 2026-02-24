import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
import re
import matplotlib.pyplot as plt
import matplotlib.patches as patches

# --- HEX 색상을 RGB로 변환하는 함수 ---
def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return tuple(int(hex_str[i:i+2], 16) for i in (0, 2, 4))

# --- 1. 데이터 파싱 및 트리 구조화 ---
def parse_data(df):
    structured_data = []
    for _, row in df.iterrows():
        code = str(row[0]).strip() if pd.notnull(row[0]) else ""
        name = str(row[1]).strip() if pd.notnull(row[1]) else ""
        if not code: continue

        match = re.match(r'^([\d\.]+)', code)
        if match:
            clean_code = match.group(1).rstrip('.')
            level = clean_code.count('.') + 1
            structured_data.append({
                'id_code': clean_code,
                'code_text': code,
                'name_text': name,
                'level': level
            })
    return structured_data

def build_tree(data):
    def split_code(c):
        return [int(i) for i in c['id_code'].split('.')]
    data.sort(key=split_code)
    
    nodes = {}
    root_nodes = []
    for item in data:
        code = item['id_code']
        node = {
            'code': code, 'code_text': item['code_text'], 
            'name_text': item['name_text'], 'level': item['level'], 'children': []
        }
        nodes[code] = node
        parts = code.split('.')
        if len(parts) > 1:
            parent_code = ".".join(parts[:-1])
            if parent_code in nodes:
                nodes[parent_code]['children'].append(node)
            else:
                if item['level'] == 1: root_nodes.append(node)
        else:
            root_nodes.append(node)
    return root_nodes

# --- 2. 좌표 및 레이아웃 계산 (고정 너비/높이 사용) ---
def calculate_layout(root_nodes, config):
    layout_data = []
    # 전체 너비와 높이는 고정 (31cm, 16cm)
    wbs_w, wbs_h = 31.0, 16.0
    l1_gap_x = config['l1_gap_x']
    l2_gap_x = config['l2_gap_x']
    v_gap_a = config['v_gap_a']
    extra_gaps = {3: config['extra_l3'], 4: config['extra_l4']}

    start_x = (33.8 - wbs_w) / 2
    start_y = (19.05 - wbs_h) / 2

    l1_count = len(root_nodes)
    if l1_count == 0: return []
    l1_width = (wbs_w - (l1_gap_x * (l1_count - 1))) / l1_count

    for i, l1 in enumerate(root_nodes):
        x_l1 = start_x + (i * (l1_width + l1_gap_x))
        layout_data.append({'node': l1, 'x': x_l1, 'y': start_y, 'w': l1_width, 'h': 1.0, 'level': 1})

        if l1['children']:
            l2_count = len(l1['children'])
            l2_width = (l1_width - (l2_gap_x * (l2_count - 1))) / l2_count
            current_y_for_l2 = start_y + 1.0 + v_gap_a

            for j, l2 in enumerate(l1['children']):
                x_l2 = x_l1 + (j * (l2_width + l2_gap_x))
                layout_data.append({'node': l2, 'x': x_l2, 'y': current_y_for_l2, 'w': l2_width, 'h': 1.0, 'level': 2})

                def stack_recursive(parent_node, px, py, pw, ph):
                    nonlocal layout_data
                    last_y = py + ph
                    for idx, child in enumerate(parent_node['children']):
                        lvl = child['level']
                        gap = v_gap_a + (extra_gaps.get(lvl, 0) if idx > 0 else 0)
                        target_y = last_y + gap
                        reduction = 0.3 * (lvl - 2)
                        c_w = max(l2_width - reduction, 4.0) 
                        c_x = (px + pw) - c_w
                        c_h = 0.8
                        layout_data.append({'node': child, 'x': c_x, 'y': target_y, 'w': c_w, 'h': c_h, 'level': lvl})
                        last_y = stack_recursive(child, c_x, target_y, c_w, c_h)
                    return last_y
                stack_recursive(l2, x_l2, current_y_for_l2, l2_width, 1.0)
    return layout_data

# --- 3. 실시간 미리보기 (Matplotlib) ---
def draw_preview(layout_data, code_width_cm, level_colors):
    fig, ax = plt.subplots(figsize=(12, 6.75))
    ax.set_xlim(0, 33.8)
    ax.set_ylim(0, 19.05)
    ax.invert_yaxis()
    ax.add_patch(patches.Rectangle((0, 0), 33.8, 19.05, linewidth=1, edgecolor='#cccccc', facecolor='#ffffff'))

    for item in layout_data:
        x, y, w, h = item['x'], item['y'], item['w'], item['h']
        lvl = item['level']
        node = item['node']
        
        # 선택된 색상 적용
        lvl_key = f"L{lvl}" if lvl <= 4 else "L5+"
        main_color = level_colors[lvl_key]
        text_c = 'white' if lvl <= 2 else 'black'

        # 1. 메인 박스
        ax.add_patch(patches.Rectangle((x, y), w, h, linewidth=0.5, edgecolor='#aaaaaa', facecolor=main_color))
        # 2. 코드 박스 (약간 더 어둡게 표시)
        ax.add_patch(patches.Rectangle((x, y), code_width_cm, h, linewidth=0.5, edgecolor='#666666', facecolor='#000000', alpha=0.1))
        # 3. 텍스트
        ax.text(x + code_width_cm/2, y + h/2, node['code_text'], color=text_c, fontsize=6, ha='center', va='center', fontweight='bold')
        display_name = (node['name_text'][:15] + '..') if len(node['name_text']) > 15 else node['name_text']
        ax.text(x + code_width_cm + 0.2, y + h/2, display_name, color=text_c, fontsize=6, ha='left', va='center')

    ax.set_axis_off()
    st.pyplot(fig)

# --- 4. PPT 생성 (그림자 제거 및 커스텀 색상) ---
def generate_ppt(layout_data, code_width_cm, level_colors):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Cm(33.8), Cm(19.05)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    for item in layout_data:
        node = item['node']
        lvl = item['level']
        x, y, w, h = item['x'], item['y'], item['w'], item['h']
        
        lvl_key = f"L{lvl}" if lvl <= 4 else "L5+"
        rgb = hex_to_rgb(level_colors[lvl_key])

        # 1. 메인 박스
        main_shp = slide.shapes.add_shape(1, Cm(x), Cm(y), Cm(w), Cm(h))
        main_shp.line.color.rgb = RGBColor(180, 180, 180)
        main_shp.fill.solid()
        main_shp.fill.fore_color.rgb = RGBColor(*rgb)
        
        # 그림자 명시적 제거
        main_shp.shadow.inherit = False 

        # 명칭 텍스트
        tf = main_shp.text_frame
        tf.text = node['name_text']
        tf.margin_left = Cm(code_width_cm + 0.2)
        tf.vertical_anchor = 1
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        p.font.size = Pt(9 if lvl > 2 else 11)
        p.font.color.rgb = RGBColor(255, 255, 255) if lvl <= 2 else RGBColor(0, 0, 0)

        # 2. 코드 박스
        code_shp = slide.shapes.add_shape(1, Cm(x), Cm(y), Cm(code_width_cm), Cm(h))
        code_shp.line.color.rgb = RGBColor(180, 180, 180)
        code_shp.fill.solid()
        # 코드박스는 메인박스 색상을 그대로 쓰거나 반투명하게 겹침 (여기선 구분 위해 살짝 어둡게 처리)
        code_shp.fill.fore_color.rgb = RGBColor(*rgb) 
        code_shp.shadow.inherit = False

        code_tf = code_shp.text_frame
        code_tf.text = node['code_text']
        cp = code_tf.paragraphs[0]
        cp.alignment = PP_ALIGN.CENTER
        cp.font.size = Pt(8 if lvl > 2 else 10)
        cp.font.bold = True
        cp.font.color.rgb = RGBColor(255, 255, 255) if lvl <= 2 else RGBColor(0, 0, 0)

    return prs

# --- 5. Streamlit UI ---
st.set_page_config(page_title="WBS Designer Pro", layout="wide")

st.sidebar.title("🎨 레벨별 색상 설정")
c1 = st.sidebar.color_picker("Level 1 (Top)", "#1F497D")
c2 = st.sidebar.color_picker("Level 2", "#365F91")
c3 = st.sidebar.color_picker("Level 3", "#D9E1F2")
c4 = st.sidebar.color_picker("Level 4", "#F2F2F2")
c5 = st.sidebar.color_picker("Level 5+", "#FFFFFF")

level_colors = {"L1": c1, "L2": c2, "L3": c3, "L4": c4, "L5+": c5}

st.sidebar.title("📏 간격 조정")
code_w_input = st.sidebar.slider("코드 박스 너비 (cm)", 1.0, 5.0, 2.2, 0.1)
v_gap_a = st.sidebar.slider("기본 수직 간격", 0.0, 2.0, 0.4, 0.05)
l1_gap_x = st.sidebar.slider("L1 좌우 간격", 0.0, 5.0, 1.2, 0.1)

with st.sidebar.expander("↕️ 그룹간 추가 여백"):
    extra_l3 = st.number_input("L3 그룹 시작 시", 0.0, 2.0, 0.4)
    extra_l4 = st.number_input("L4 그룹 시작 시", 0.0, 2.0, 0.2)

config = {
    'l1_gap_x': l1_gap_x, 'l2_gap_x': 0.4, 'v_gap_a': v_gap_a, 
    'extra_l3': extra_l3, 'extra_l4': extra_l4
}

st.title("📊 WBS 프로 디자이너")
uploaded_file = st.file_uploader("엑셀 파일(.xlsx) 업로드", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    raw_data = parse_data(df)
    
    if raw_data:
        tree = build_tree(raw_data)
        layout_data = calculate_layout(tree, config)
        
        st.subheader("🖼️ 실시간 디자인 미리보기")
        draw_preview(layout_data, code_w_input, level_colors)
        
        if st.button("🚀 PPT 생성 및 다운로드", use_container_width=True):
            final_ppt = generate_ppt(layout_data, code_w_input, level_colors)
            ppt_io = io.BytesIO()
            final_ppt.save(ppt_io)
            ppt_io.seek(0)
            st.download_button("🎁 PPT 파일 다운로드", ppt_io, "Custom_WBS.pptx")
