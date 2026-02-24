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
        node = {'code': code, 'code_text': item['code_text'], 'name_text': item['name_text'], 'level': item['level'], 'children': []}
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

# --- 2. 좌표 및 레이아웃 계산 ---
def calculate_layout(root_nodes, config):
    layout_data = []
    wbs_w = config['wbs_w']
    wbs_h = config['wbs_h']
    l1_gap_x = config['l1_gap_x']
    l2_gap_x = config['l2_gap_x']
    v_gap_a = config['v_gap_a']
    extra_gaps = {3: config['extra_l3'], 4: config['extra_l4'], 5: config['extra_l5']}
    start_x = (33.8 - wbs_w) / 2
    start_y = (19.05 - wbs_h) / 2
    l1_count = len(root_nodes)
    if l1_count == 0: return []
    l1_width = (wbs_w - (l1_gap_x * (l1_count - 1))) / l1_count
    for i, l1 in enumerate(root_nodes):
        x_l1 = start_x + (i * (l1_width + l1_gap_x))
        y_l1 = start_y
        h_l1 = 1.0
        layout_data.append({'node': l1, 'x': x_l1, 'y': y_l1, 'w': l1_width, 'h': h_l1, 'level': 1})
        if l1['children']:
            l2_count = len(l1['children'])
            l2_width = (l1_width - (l2_gap_x * (l2_count - 1))) / l2_count
            current_y_for_l2 = y_l1 + h_l1 + v_gap_a
            for j, l2 in enumerate(l1['children']):
                x_l2 = x_l1 + (j * (l2_width + l2_gap_x))
                y_l2 = current_y_for_l2
                h_l2 = 1.0
                layout_data.append({'node': l2, 'x': x_l2, 'y': y_l2, 'w': l2_width, 'h': h_l2, 'level': 2})
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
                stack_recursive(l2, x_l2, y_l2, l2_width, h_l2)
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
        
        lvl_key = f"L{lvl}" if lvl <= 4 else "L5+"
        color = level_colors[lvl_key]
        text_c = 'white' if lvl <= 2 else 'black'

        # 1. 메인 박스
        ax.add_patch(patches.Rectangle((x, y), w, h, linewidth=0.5, edgecolor='#aaaaaa', facecolor=color))
        
        # 2. 코드 박스 (살짝 어둡게 오버레이)
        ax.add_patch(patches.Rectangle((x, y), code_width_cm, h, linewidth=0.5, edgecolor='#666666', facecolor='#000000', alpha=0.05))
        
        # [신규] 3. 가장 왼쪽 0.2cm 포인트 띠 (Accent Strip)
        ax.add_patch(patches.Rectangle((x, y), 0.2, h, linewidth=0, facecolor='#000000', alpha=0.3))
        
        ax.text(x + code_width_cm/2, y + h/2, node['code_text'], color=text_c, fontsize=6, ha='center', va='center', fontweight='bold')
        display_name = (node['name_text'][:20] + '..') if len(node['name_text']) > 20 else node['name_text']
        ax.text(x + code_width_cm + 0.2, y + h/2, display_name, color=text_c, fontsize=6, ha='left', va='center')
    ax.set_axis_off()
    st.pyplot(fig)

# --- 4. PPT 생성 (그림자 제거 + 포인트 띠 추가) ---
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
        main_shp.line.color.rgb = RGBColor(150, 150, 150)
        main_shp.fill.solid()
        main_shp.fill.fore_color.rgb = RGBColor(*rgb)
        main_shp.shadow.inherit = False 

        tf = main_shp.text_frame
        tf.text = node['name_text']
        tf.margin_left = Cm(code_width_cm + 0.2)
        tf.vertical_anchor = 1
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        p.font.size = Pt(9 if lvl > 2 else 11)
        p.font.color.rgb = RGBColor(255, 255, 255) if lvl <= 2 else RGBColor(0, 0, 0)

        # 2. 코드 박스 (위에 덮음)
        code_shp = slide.shapes.add_shape(1, Cm(x), Cm(y), Cm(code_width_cm), Cm(h))
        code_shp.fill.solid()
        code_shp.fill.fore_color.rgb = RGBColor(*rgb)
        code_shp.line.color.rgb = RGBColor(150, 150, 150)
        code_shp.shadow.inherit = False 
        
        code_tf = code_shp.text_frame
        code_tf.text = node['code_text']
        code_p = code_tf.paragraphs[0]
        code_p.alignment = PP_ALIGN.CENTER
        code_p.font.size = Pt(8 if lvl > 2 else 10)
        code_p.font.bold = True
        code_p.font.color.rgb = RGBColor(255, 255, 255) if lvl <= 2 else RGBColor(0, 0, 0)

        # [신규 추가] 3. 가장 왼쪽 0.2cm 포인트 띠 (Accent Strip)
        accent_shp = slide.shapes.add_shape(1, Cm(x), Cm(y), Cm(0.2), Cm(h))
        accent_shp.fill.solid()
        # 원본 색상보다 아주 약간 어둡게 하거나, 레벨이 높으면 검정계열로 강조
        if lvl <= 2:
            accent_shp.fill.fore_color.rgb = RGBColor(0, 0, 0) # L1, L2는 검정 포인트
        else:
            accent_shp.fill.fore_color.rgb = RGBColor(100, 100, 100) # 그 외는 회색 포인트
        accent_shp.line.fill.background() # 테두리 없음
        accent_shp.shadow.inherit = False

    return prs

# --- 5. Streamlit UI (기존과 동일) ---
st.set_page_config(page_title="WBS Master Pro", layout="wide")
st.sidebar.title("🎨 디자인 및 간격 설정")

with st.sidebar.expander("🎨 레벨별 색상 설정", expanded=True):
    c1 = st.color_picker("Level 1", "#1F497D")
    c2 = st.color_picker("Level 2", "#365F91")
    c3 = st.color_picker("Level 3", "#F2F2F2")
    c4 = st.color_picker("Level 4", "#F2F2F2")
    c5 = st.color_picker("Level 5+", "#F2F2F2")
    level_colors = {"L1": c1, "L2": c2, "L3": c3, "L4": c4, "L5+": c5}

code_w_input = st.sidebar.slider("코드 박스 너비 (cm)", 1.0, 5.0, 2.5, 0.1)

with st.sidebar.expander("📏 캔버스 및 기본 간격", expanded=True):
    wbs_w = st.number_input("
