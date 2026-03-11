# -*- coding: utf-8 -*-
from powerpoint_utils import *

def render_3_cards(slide, data):
    """
    3개 카드 레이아웃 (동적 높이 계산)

    개선사항:
    - 모든 카드의 최대 본문 줄 수 계산
    - 각 줄당 0.28인치 (11pt 폰트 + 라인 간격 1.1 + 여백)
    - 3줄 이상의 본문 텍스트 완벽 지원
    """
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    # 카드는 항상 가용 높이 전체를 채움 (auto_size로 오버플로우 방지)
    card_h = bh
    gap = Inches(0.3); w_card = (bw - (gap * 2)) / 3
    for i, key in enumerate(['card_1', 'card_2', 'card_3']):
        item = content.get(key, {})
        x = bx + i * (w_card + gap)
        create_content_box(slide, x, by, w_card, card_h, "", "", "white")

        # 아이콘+텍스트 그룹을 카드 내에서 수직 중앙 정렬
        icon_size = Inches(0.8)
        icon_text_gap = Inches(0.15)
        text_area_height = card_h * 0.65

        total_content_height = icon_size + icon_text_gap + text_area_height
        top_margin = (card_h - total_content_height) / 2

        # 아이콘 배치
        icon_y = by + top_margin
        draw_icon_search(slide, x + w_card/2 - icon_size/2, icon_y, icon_size, item.get('search_q'))

        # 텍스트박스 배치 (auto_size로 오버플로우 방지)
        text_y = icon_y + icon_size + icon_text_gap
        text_height = text_area_height
        tb = slide.shapes.add_textbox(x, text_y, w_card, text_height)
        tb.text_frame.word_wrap = True
        tb.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tb.text_frame.margin_left = Inches(0.2)
        tb.text_frame.margin_right = Inches(0.2)
        tb.text_frame.margin_top = Inches(0.1)
        tb.text_frame.margin_bottom = Inches(0.1)
        p = tb.text_frame.paragraphs[0]; p.text = item.get('title',''); p.font.bold=True; p.font.size=Pt(17); p.alignment=PP_ALIGN.CENTER; p.font.color.rgb=COLORS["PRIMARY"]; p.font.name=FONTS["BODY_TITLE"]
        body_lines = [l for l in item.get('body', '').split('\n') if l.strip()] or ['']
        is_list = len(body_lines) > 1
        import re as _re
        for li, body_line in enumerate(body_lines):
            p2 = tb.text_frame.add_paragraph()
            stripped = body_line.strip()
            is_numbered = bool(_re.match(r'^\d+[.):\s]\s', stripped))
            if is_list and stripped and not stripped.startswith('•') and not is_numbered:
                p2.text = "• " + body_line
            else:
                p2.text = body_line
            p2.font.size=Pt(13); p2.alignment=PP_ALIGN.CENTER; p2.font.color.rgb=COLORS["BLACK"]; p2.font.name=FONTS["BODY_TEXT"]
            if li == 0: p2.space_before = Pt(8)

# 1. Bento Grid
def render_bento_grid(slide, data):
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    gap = Inches(0.2); w_main = (bw - gap) / 2
    # main도 터미널이면 compact 적용
    main_compact = content.get('main',{}).get('terminal', False)
    create_content_box(slide, bx, by, w_main, bh, content.get('main',{}).get('title'), content.get('main',{}).get('body'), "gray", content.get('main',{}).get('search_q'), compact=main_compact, terminal=content.get('main',{}).get('terminal', False))
    h_sub = (bh - gap) / 2
    create_content_box(slide, bx+w_main+gap, by, w_main, h_sub, content.get('sub1',{}).get('title'), content.get('sub1',{}).get('body'), "white", content.get('sub1',{}).get('search_q'), compact=True, terminal=content.get('sub1',{}).get('terminal', False))
    create_content_box(slide, bx+w_main+gap, by+h_sub+gap, w_main, h_sub, content.get('sub2',{}).get('title'), content.get('sub2',{}).get('body'), "white", content.get('sub2',{}).get('search_q'), compact=True, terminal=content.get('sub2',{}).get('terminal', False))

# 3. Grid 2x2
def render_grid_2x2(slide, data):
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    gap = Inches(0.2); w_half = (bw - gap) / 2; h_half = (bh - gap) / 2
    coords = [(0,0), (1,0), (0,1), (1,1)]
    for i, key in enumerate(['item1', 'item2', 'item3', 'item4']):
        item = content.get(key, {})
        c, r = coords[i]
        create_content_box(slide, bx + c*(w_half+gap), by + r*(h_half+gap), w_half, h_half, item.get('title'), item.get('body'), "white", item.get('search_q'), compact=True, terminal=item.get('terminal', False))

# 4. Quad Matrix (Alias)
def render_quad_matrix(slide, data): render_grid_2x2(slide, data)

# 5. Challenge & Solution
def render_challenge_solution(slide, data):
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    # challenge와 solution은 wrapper 레벨에 있음
    content = wrapper

    # challenge/solution은 dict 또는 string일 수 있음
    ch = content.get('challenge', {})
    sol = content.get('solution', {})
    ch_title = ch.get('title', 'CHALLENGE') if isinstance(ch, dict) else 'CHALLENGE'
    ch_body = ch.get('body', '') if isinstance(ch, dict) else str(ch)
    sol_title = sol.get('title', 'SOLUTION') if isinstance(sol, dict) else 'SOLUTION'
    sol_body = sol.get('body', '') if isinstance(sol, dict) else str(sol)

    gap = Inches(0.6); w_half = (bw - gap) / 2
    create_content_box(slide, bx, by, w_half, bh, ch_title, ch_body, "gray")
    create_content_box(slide, bx+w_half+gap, by, w_half, bh, sol_title, sol_body, "white")
    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, bx+w_half-Inches(0.5)+(gap/2), by+(bh/2)-Inches(0.5), Inches(1.0), Inches(1.0))
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = COLORS["PRIMARY"]

# 6. Timeline Steps
def render_timeline_steps(slide, data):
    """카드 형태 타임라인 (가시성 최적화)"""
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    steps = content.get('steps', [])
    if not steps: return

    # 카드 간 간격
    arrow_gap = Inches(0.4)
    card_width = (bw - (arrow_gap * (len(steps) - 1))) / len(steps)

    for i, step in enumerate(steps):
        x = bx + i * (card_width + arrow_gap)

        # 카드 박스 (명확한 배경)
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, by, card_width, bh)
        card.fill.solid()
        card.fill.fore_color.rgb = COLORS["BG_BOX"]
        card.line.color.rgb = COLORS["PRIMARY"]
        card.line.width = Pt(2.0)

        # 숫자 배지 (큰 원형)
        badge_size = Inches(0.8)
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, x + (card_width/2) - (badge_size/2), by + Inches(0.3), badge_size, badge_size)
        badge.fill.solid()
        badge.fill.fore_color.rgb = COLORS["PRIMARY"]
        badge.line.color.rgb = COLORS["PRIMARY"]

        # 배지 숫자
        tf_badge = badge.text_frame
        tf_badge.clear()
        tf_badge.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_badge = tf_badge.paragraphs[0]
        p_badge.text = str(i + 1)
        p_badge.font.name = FONTS["BODY_TITLE"]
        p_badge.font.bold = True
        p_badge.font.size = Pt(28)
        p_badge.font.color.rgb = COLORS["BG_WHITE"]
        p_badge.alignment = PP_ALIGN.CENTER

        # 텍스트 영역 (아이콘 제거 → 배지 바로 아래에서 시작)
        text_y = by + Inches(1.3)
        text_h = bh - Inches(1.3) - Inches(0.3)
        tb = slide.shapes.add_textbox(x + Inches(0.2), text_y, card_width - Inches(0.4), text_h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.15)
        tf.margin_right = Inches(0.15)

        # 날짜/기간
        p = tf.paragraphs[0]
        p.text = step.get('date','')
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(16)
        p.font.color.rgb = COLORS["PRIMARY"]
        p.font.name = FONTS["BODY_TITLE"]
        p.space_after = Pt(10)

        # 설명
        p2 = tf.add_paragraph()
        p2.text = step.get('desc','')
        p2.font.size = Pt(14)
        p2.alignment = PP_ALIGN.CENTER
        p2.font.color.rgb = COLORS["BLACK"]
        p2.font.name = FONTS["BODY_TEXT"]
        p2.line_spacing = 1.3

        # 단계 간 화살표 (마지막 단계 제외)
        if i < len(steps) - 1:
            arrow_x = x + card_width + Inches(0.05)
            arrow_y = by + (bh / 2) - Inches(0.3)
            arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, arrow_x, arrow_y, arrow_gap - Inches(0.1), Inches(0.6))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = COLORS["PRIMARY"]
            arrow.line.color.rgb = COLORS["PRIMARY"]

# 7. Process Arrow
def render_process_arrow(slide, data):
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    steps = content.get('steps', [])
    if not steps: return
    gap = Inches(0.3); w_step = (bw - (gap * (len(steps)-1))) / len(steps)
    for i, step in enumerate(steps):
        x = bx + i*(w_step+gap)
        shp = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, x, by, w_step, Inches(0.8))
        shp.fill.solid(); shp.fill.fore_color.rgb = COLORS["PRIMARY"]
        p = shp.text_frame.paragraphs[0]; p.text = step.get('title',''); p.font.color.rgb = COLORS["BG_WHITE"]; p.alignment=PP_ALIGN.CENTER; p.font.size=Pt(14); p.font.bold=True
        create_content_box(slide, x, by + Inches(1.0), w_step, bh - Inches(1.0), "", step.get('body',''), "white", step.get('search_q'), terminal=step.get('terminal', False))

# 7-2. Phased Columns (단계별 컬럼 + 의미 기반 색상)
def render_phased_columns(slide, data):
    """단계별 컬럼 레이아웃 (의미 기반 색상)

    N개 세로 컬럼 나란히 배치, 의미 기반 고유 색상.
    각 컬럼: 색상 헤더 스트립 + 본문 내용 + 아이콘

    data.data.data.steps: [
        {"title": "1. 현황분석", "body": "...", "search_q": "..."},
        ...
    ]
    """
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    steps = content.get('steps', [])
    if not steps:
        return

    n = len(steps)
    gap = Inches(0.15)
    col_w = (bw - gap * (n - 1)) / n
    header_h = Inches(0.7)

    # 의미 기반 색상 팔레트 (각 단계별 고유 색상)
    _phase_colors = [
        COLORS["PRIMARY"],          # 파랑
        RGBColor(4, 120, 87),       # 초록
        RGBColor(194, 65, 12),      # 주황
        RGBColor(185, 28, 28),      # 빨강
        RGBColor(30, 58, 138),      # 진파랑
        RGBColor(120, 53, 15),      # 갈색
        RGBColor(88, 28, 135),      # 보라
    ]
    colors = [_phase_colors[i % len(_phase_colors)] for i in range(n)]

    for i, step in enumerate(steps):
        x = bx + i * (col_w + gap)

        # 헤더 스트립 (색상)
        header = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, by, col_w, header_h)
        header.fill.solid()
        header.fill.fore_color.rgb = colors[i]
        header.line.color.rgb = colors[i]
        header.line.width = Pt(0.5)

        tf = header.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.1)
        tf.margin_right = Inches(0.1)
        p = tf.paragraphs[0]
        p.text = step.get('title', '')
        p.font.name = FONTS["BODY_TITLE"]
        p.font.bold = True
        p.font.size = Pt(13)
        p.font.color.rgb = COLORS["BG_WHITE"]
        p.alignment = PP_ALIGN.CENTER

        # 본문 박스 (헤더 아래)
        body_y = by + header_h + Inches(0.1)
        body_h = bh - header_h - Inches(0.1)

        body_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, body_y, col_w, body_h)
        body_box.fill.solid()
        body_box.fill.fore_color.rgb = COLORS["BG_BOX"]
        body_box.line.color.rgb = COLORS["BORDER"]
        body_box.line.width = Pt(1.0)

        tf_body = body_box.text_frame
        tf_body.clear()
        tf_body.word_wrap = True
        tf_body.vertical_anchor = MSO_ANCHOR.TOP
        tf_body.margin_left = Inches(0.15)
        tf_body.margin_right = Inches(0.15)
        tf_body.margin_top = Inches(0.15)
        tf_body.margin_bottom = Inches(0.15)

        body_text = step.get('body', '')
        lines = str(body_text).split('\n')
        for j, line in enumerate(lines):
            p = tf_body.paragraphs[0] if j == 0 else tf_body.add_paragraph()
            p.text = line
            p.font.name = FONTS["BODY_TEXT"]
            p.font.size = Pt(12)
            p.font.color.rgb = COLORS["BLACK"]
            p.alignment = PP_ALIGN.LEFT
            p.space_after = Pt(4)
            p.level = 0  # 글머리 기호 레벨 설정

        # 아이콘 (본문 박스 우하단)
        if step.get('search_q'):
            icon_size = Inches(0.5)
            icon_x = x + col_w - icon_size - Inches(0.1)
            icon_y = body_y + body_h - icon_size - Inches(0.1)
            draw_icon_search(slide, icon_x, icon_y, icon_size, step['search_q'])

# 8. Architecture Wide
def render_architecture_wide(slide, data):
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    h_diag = bh * 0.45
    # 다이어그램 영역: 로컬 이미지 또는 아이콘+화살표 폴백
    diagram_src = content.get('diagram_path', '')
    diag_loaded = False
    if diagram_src and os.path.exists(diagram_src):
        try:
            slide.shapes.add_picture(diagram_src, bx, by, width=bw, height=h_diag)
            diag_loaded = True
        except: pass
    if not diag_loaded:
        # 아이콘+화살표 폴백: 컬럼 아이콘들을 가로로 배치
        cols_data = [content.get(f'col{i+1}', {}) for i in range(3)]
        icon_keys = [c.get('search_q', '') for c in cols_data if isinstance(c, dict)]
        if icon_keys:
            icon_n = len(icon_keys); icon_size = Inches(1.0)
            icon_gap = (bw - icon_size * icon_n) / max(icon_n + 1, 1)
            for idx, sq in enumerate(icon_keys):
                ix = bx + icon_gap * (idx + 1) + icon_size * idx
                iy = by + (h_diag - icon_size) / 2
                draw_icon_search(slide, ix, iy, icon_size, sq)
                if idx < icon_n - 1:
                    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, ix + icon_size + Inches(0.1), by + h_diag / 2 - Inches(0.15), icon_gap - Inches(0.2), Inches(0.3))
                    arrow.fill.solid(); arrow.fill.fore_color.rgb = COLORS["PRIMARY"]; arrow.line.color.rgb = COLORS["PRIMARY"]
        else:
            ph = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, bx, by, bw, h_diag)
            ph.fill.solid(); ph.fill.fore_color.rgb = RGBColor(230,230,230); ph.text_frame.text = "Diagram Area"

    y_desc = by + h_diag + Inches(0.2); h_desc = bh - h_diag - Inches(0.2); gap = Inches(0.15); w_col = (bw - (gap*2)) / 3
    for i, k in enumerate(['col1', 'col2', 'col3']):
        if k in content:
            item = content[k]
            create_content_box(slide, bx + i*(w_col+gap), y_desc, w_col, h_desc, item.get('title',''), item.get('body',''), "white", item.get('search_q'), compact=True)

# 9. Image Left
def render_image_left(slide, data):
    """좌측 이미지 + 우측 텍스트 레이아웃 (개조식)"""
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)

    content = wrapper.get('data', {})
    gap = Inches(0.25)
    w_half = (bw - gap) / 2

    # 좌측: 이미지 (image_path 또는 search_q 폴더 체인)
    image_path = content.get('image_path')
    img_loaded = False
    if image_path and os.path.exists(image_path):
        try:
            slide.shapes.add_picture(image_path, bx, by, width=w_half, height=bh)
            img_loaded = True
        except Exception as e:
            print(f"⚠️ 이미지 로드 실패: {str(e)[:50]}")
    if not img_loaded:
        sq = content.get('search_q', '')
        if sq:
            for folder in ['architecture', 'screenshots', 'icons']:
                candidate = os.path.join(folder, sq.replace(' ', '_') + '.png')
                if os.path.exists(candidate):
                    try:
                        slide.shapes.add_picture(candidate, bx, by, width=w_half, height=bh)
                        img_loaded = True
                    except: pass
                    break
    if not img_loaded:
        placeholder = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, bx, by, w_half, bh)
        placeholder.fill.solid()
        placeholder.fill.fore_color.rgb = COLORS["BG_BOX"]
        placeholder.line.color.rgb = COLORS["BORDER"]

    # 우측: 텍스트 (개조식 - 불릿 포인트)
    text_x = bx + w_half + gap
    text_box = slide.shapes.add_textbox(text_x, by, w_half, bh)
    tf = text_box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.2)
    tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.3)
    tf.margin_bottom = Inches(0.3)

    # bullets 배열 처리
    bullets = content.get('bullets', [])
    if bullets:
        for i, bullet in enumerate(bullets):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()

            p.text = f"• {bullet}"
            p.font.name = FONTS["BODY_TEXT"]
            p.font.size = Pt(16)
            p.font.color.rgb = COLORS["BLACK"]
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = 1.3
            p.space_after = Pt(12)
    else:
        # 하위 호환성: body 필드 지원
        body_text = content.get('body', '')
        if body_text:
            lines = body_text.split('\n')
            for i, line in enumerate(lines):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()

                p.text = line.strip()
                p.font.name = FONTS["BODY_TEXT"]
                p.font.size = Pt(14)
                p.font.color.rgb = COLORS["BLACK"]
                p.alignment = PP_ALIGN.LEFT
                p.line_spacing = 1.2
                p.space_after = Pt(8)

# 10. Comparison VS
def render_comparison_vs(slide, data):
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    # VS 원형을 위한 충분한 간격 확보 (더 증가)
    gap = Inches(0.8); w_half = (bw - gap) / 2
    # 아이콘 없이 텍스트만 표시 (comparison_vs는 텍스트 비교가 목적)
    create_content_box(slide, bx, by, w_half, bh, content.get('item_a_title','A'), content.get('item_a_body',''), "gray")
    create_content_box(slide, bx + w_half + gap, by, w_half, bh, content.get('item_b_title','B'), content.get('item_b_body',''), "white")

    # VS 원형 + 텍스트
    oval_x = bx + w_half - Inches(0.5) + (gap/2)
    oval_y = by + (bh/2) - Inches(0.5)
    oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, oval_x, oval_y, Inches(1.0), Inches(1.0))
    oval.fill.solid()
    oval.fill.fore_color.rgb = COLORS["PRIMARY"]
    oval.line.color.rgb = COLORS["PRIMARY"]

    # VS 텍스트 추가
    tf = oval.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = "VS"
    p.font.name = FONTS["BODY_TITLE"]
    p.font.bold = True
    p.font.size = Pt(20)
    p.font.color.rgb = COLORS["BG_WHITE"]
    p.alignment = PP_ALIGN.CENTER

# 11. Key Metric
def render_key_metric(slide, data): render_3_cards(slide, data)

# 12. Detail Image
def render_detail_image(slide, data):
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    # 텍스트 박스 높이를 35%로 증가 (기존 25% → 35%)
    h_text = bh * 0.35
    create_content_box(slide, bx, by, bw, h_text, content.get('title',''), content.get('body',''), "gray")

    # 이미지 영역 높이 조정
    img_y = by + h_text + Inches(0.2)
    img_h = bh - h_text - Inches(0.2)

    # 회색 텍스트박스로 이미지 설명 표시
    search_q = content.get('search_q', '')
    
    # 회색 박스 생성
    placeholder = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, img_y, bw, img_h)
    placeholder.fill.solid()
    placeholder.fill.fore_color.rgb = RGBColor(240, 240, 240)  # 연한 회색
    placeholder.line.color.rgb = COLORS["BORDER"]
    placeholder.line.width = Pt(1.5)
    
    # 텍스트 프레임 설정
    tf = placeholder.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.5)
    tf.margin_right = Inches(0.5)
    tf.margin_top = Inches(0.5)
    tf.margin_bottom = Inches(0.5)
    
    # 이미지 설명 텍스트 작성
    if search_q:
        # 검색어 기반 상세 설명
        image_descriptions = {
            'kubernetes': '쿠버네티스 아키텍처 다이어그램\n\n- Control Plane (Master Node)\n- Worker Nodes\n- Pod, Service, Ingress 구조\n- 컨테이너 오케스트레이션 흐름도',
            'aws_architecture': 'AWS 클라우드 아키텍처 다이어그램\n\n- VPC 네트워크 구조\n- EC2, RDS, S3 등 주요 서비스\n- Load Balancer 및 Auto Scaling\n- Multi-AZ 고가용성 구성',
            'microservices': '마이크로서비스 아키텍처 다이어그램\n\n- API Gateway\n- 서비스 간 통신 (REST/gRPC)\n- Service Mesh\n- 데이터베이스 분리 패턴',
            'cicd': 'CI/CD 파이프라인 다이어그램\n\n- Source Control (Git)\n- Build & Test\n- Deploy to Staging/Production\n- Monitoring & Rollback',
            'database': '데이터베이스 아키텍처 다이어그램\n\n- Master-Slave 복제 구조\n- Sharding 전략\n- 캐싱 레이어 (Redis)\n- 백업 및 복구 프로세스'
        }
        
        description = image_descriptions.get(search_q, 
            f'{search_q.replace("_", " ").title()} 관련 다이어그램\n\n- 시스템 구성도\n- 주요 컴포넌트\n- 데이터 흐름\n- 연동 관계')
    else:
        description = '이미지 영역\n\n다이어그램 또는 스크린샷을 삽입하세요'
    
    # 제목 추가
    p_title = tf.paragraphs[0]
    p_title.text = '[이미지 영역]'
    p_title.font.name = FONTS["BODY_TITLE"]
    p_title.font.size = Pt(18)
    p_title.font.bold = True
    p_title.font.color.rgb = COLORS["DARK_GRAY"]
    p_title.alignment = PP_ALIGN.CENTER
    p_title.space_after = Pt(12)
    
    # 설명 추가
    p_desc = tf.add_paragraph()
    p_desc.text = description
    p_desc.font.name = FONTS["BODY_TEXT"]
    p_desc.font.size = Pt(13)
    p_desc.font.color.rgb = COLORS["DARK_GRAY"]
    p_desc.alignment = PP_ALIGN.LEFT
    p_desc.line_spacing = 1.4
    
    print(f"   ✅ [이미지 영역 생성] '{search_q}' - 상세 설명 포함")

# 13. Comparison Table
def render_comparison_table(slide, data):
    """표 형태 비교 레이아웃 (3열 비교)"""
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    # 3개 열 데이터
    columns = content.get('columns', [])
    if not columns or len(columns) != 3:
        return

    gap = Inches(0.2)
    w_col = (bw - (gap * 2)) / 3

    # 헤더 행 (제목)
    header_h = Inches(0.8)
    for i, col in enumerate(columns):
        x = bx + i * (w_col + gap)
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, by, w_col, header_h)
        shp.fill.solid()
        shp.fill.fore_color.rgb = COLORS["PRIMARY"]
        shp.line.color.rgb = COLORS["PRIMARY"]
        shp.line.width = Pt(1.0)

        tf = shp.text_frame
        tf.clear()
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = col.get('title', '') if isinstance(col, dict) else str(col)
        p.font.name = FONTS["BODY_TITLE"]
        p.font.bold = True
        p.font.size = Pt(16)
        p.font.color.rgb = COLORS["BG_WHITE"]
        p.alignment = PP_ALIGN.CENTER

    # 데이터 행들
    rows = content.get('rows', [])
    row_h = (bh - header_h - Inches(0.2)) / len(rows) if rows else Inches(1.0)

    for row_idx, row in enumerate(rows):
        row_y = by + header_h + Inches(0.2) + (row_idx * row_h)
        values = row if isinstance(row, list) else row.get('values', ['', '', ''])

        for col_idx, value in enumerate(values):
            x = bx + col_idx * (w_col + gap)

            shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, row_y, w_col, row_h - Inches(0.05))
            shp.fill.solid()
            shp.fill.fore_color.rgb = COLORS["BG_WHITE"]
            shp.line.color.rgb = COLORS["BORDER"]
            shp.line.width = Pt(1.0)

            tf = shp.text_frame
            tf.clear()
            tf.margin_left = Inches(0.15)
            tf.margin_right = Inches(0.15)
            tf.margin_top = Inches(0.1)
            tf.margin_bottom = Inches(0.1)
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE

            p = tf.paragraphs[0]
            p.text = str(value)
            p.font.name = FONTS["BODY_TEXT"]
            p.font.size = Pt(14)
            p.font.color.rgb = COLORS["BLACK"]
            p.alignment = PP_ALIGN.CENTER

# 14. Detail Sections (KMS PPT 슬라이드 2~4 참조)

# ── 다이어그램 공통 색상 팔레트 ──
_SEM_BOX_STYLES = {
    'gray':    (RGBColor(248, 249, 250), RGBColor(150, 150, 150), RGBColor(33, 33, 33)),
    'red':     (RGBColor(254, 242, 242), RGBColor(185, 28, 28), RGBColor(127, 29, 29)),
    'orange':  (RGBColor(255, 247, 237), RGBColor(194, 65, 12), RGBColor(154, 52, 18)),
    'green':   (RGBColor(236, 253, 245), RGBColor(4, 120, 87), RGBColor(6, 95, 70)),
    'blue':    (RGBColor(239, 246, 255), RGBColor(30, 58, 138), RGBColor(30, 64, 175)),
    'navy':    (RGBColor(224, 231, 255), RGBColor(49, 46, 129), RGBColor(55, 48, 163)),
    'primary': (RGBColor(239, 246, 255), RGBColor(0, 67, 218), RGBColor(30, 64, 175)),
}

def _diagram_box(slide, x, y, w, h, label, color='gray', font_size=13):
    """공통 다이어그램 박스 그리기"""
    fill_c, line_c, text_c = _SEM_BOX_STYLES.get(color, _SEM_BOX_STYLES['gray'])
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.fill.solid(); shp.fill.fore_color.rgb = fill_c
    shp.line.color.rgb = line_c; shp.line.width = Pt(1.5)

    tf = shp.text_frame; tf.clear(); tf.word_wrap = True
    tf.margin_left = Inches(0.12); tf.margin_right = Inches(0.12)
    tf.margin_top = Inches(0.06); tf.margin_bottom = Inches(0.06)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    lines = label.split('\n')
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.name = FONTS["BODY_TEXT"]
        p.font.size = Pt(font_size) if i == 0 else Pt(font_size - 2)
        p.font.bold = (i == 0)
        p.font.color.rgb = text_c if i == 0 else COLORS["GRAY"]
        p.alignment = PP_ALIGN.CENTER; p.space_after = Pt(2)
    return shp

def _diagram_arrow_label(slide, x, y, w, h, label, direction='down'):
    """화살표 라벨 (방향 지원)"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.clear(); tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    prefix = {'down': '⬇', 'right': '➡', 'left': '⬅', 'up': '⬆'}.get(direction, '⬇')
    p = tf.paragraphs[0]
    p.text = f"{prefix} {label}" if label else prefix
    p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(11)
    p.font.color.rgb = COLORS["GRAY"]; p.alignment = PP_ALIGN.CENTER

def _diagram_shape_arrow(slide, x, y, w, h, direction='down'):
    """실제 화살표 shape (방향 지원)"""
    shape_type = {
        'down': MSO_SHAPE.DOWN_ARROW, 'right': MSO_SHAPE.RIGHT_ARROW,
        'left': MSO_SHAPE.LEFT_ARROW, 'up': MSO_SHAPE.UP_ARROW,
    }.get(direction, MSO_SHAPE.DOWN_ARROW)
    arrow = slide.shapes.add_shape(shape_type, x, y, w, h)
    arrow.fill.solid(); arrow.fill.fore_color.rgb = COLORS["PRIMARY"]
    arrow.line.color.rgb = COLORS["PRIMARY"]
    return arrow


def _draw_diagram_flow(slide, rx, ry, rw, rh, items):
    """type=flow: 수직 흐름도 (박스 + 화살표 라벨, 개수 자동 대응)"""
    boxes = [it for it in items if it.get('type') != 'arrow']
    arrows = [it for it in items if it.get('type') == 'arrow']

    arrow_h = Inches(0.3)
    gap = Inches(0.06)
    total_h = arrow_h * len(arrows) + gap * (len(items) - 1)
    box_h = (rh - total_h) / max(len(boxes), 1)

    pad_x = Inches(0.1)
    bw = rw - pad_x * 2; bx = rx + pad_x
    cy = ry

    for item in items:
        if item.get('type') == 'arrow':
            _diagram_arrow_label(slide, bx, cy, bw, arrow_h, item.get('label', ''))
            cy += arrow_h + gap
        else:
            _diagram_box(slide, bx, cy, bw, box_h, item.get('label', ''), item.get('color', 'gray'))
            cy += box_h + gap


def _draw_diagram_layers(slide, rx, ry, rw, rh, layers):
    """type=layers: 수평 계층 다이어그램 (아키텍처 티어, 분리된 영역)

    layers: [
        {"title": "Data Layer", "desc": "Encrypted — 변경 없음", "color": "green"},
        {"title": "Key Layer", "desc": "KMS CMK로 보호", "color": "blue", "items": ["CMK", "Data Key"]},
    ]
    """
    n = len(layers)
    gap = Inches(0.15)
    layer_h = (rh - gap * (n - 1)) / n
    pad_x = Inches(0.1)

    for i, layer in enumerate(layers):
        ly = ry + i * (layer_h + gap)
        fill_c, line_c, text_c = _SEM_BOX_STYLES.get(layer.get('color', 'gray'), _SEM_BOX_STYLES['gray'])

        # 외곽 테두리
        outer = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, rx + pad_x, ly, rw - pad_x * 2, layer_h)
        outer.fill.solid(); outer.fill.fore_color.rgb = fill_c
        outer.line.color.rgb = line_c; outer.line.width = Pt(2.0)

        # 레이어 제목
        title_h = Inches(0.3)
        tb = slide.shapes.add_textbox(rx + pad_x + Inches(0.15), ly + Inches(0.08), rw - pad_x * 2 - Inches(0.3), title_h)
        p = tb.text_frame.paragraphs[0]
        p.text = layer.get('title', '')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(12); p.font.color.rgb = line_c

        # 레이어 설명
        desc = layer.get('desc', '')
        if desc:
            tb_d = slide.shapes.add_textbox(rx + pad_x + Inches(0.15), ly + Inches(0.35), rw - pad_x * 2 - Inches(0.3), Inches(0.25))
            p_d = tb_d.text_frame.paragraphs[0]
            p_d.text = desc
            p_d.font.name = FONTS["BODY_TEXT"]; p_d.font.size = Pt(10)
            p_d.font.color.rgb = text_c

        # 내부 아이템 박스 (가로 배치)
        sub_items = layer.get('items', [])
        if sub_items:
            inner_y = ly + Inches(0.65)
            inner_h = layer_h - Inches(0.8)
            inner_gap = Inches(0.1)
            inner_w = (rw - pad_x * 2 - Inches(0.3) - inner_gap * (len(sub_items) - 1)) / len(sub_items)

            for j, sub in enumerate(sub_items):
                sx = rx + pad_x + Inches(0.15) + j * (inner_w + inner_gap)
                sub_label = sub if isinstance(sub, str) else sub.get('label', '')
                sub_color = 'gray' if isinstance(sub, str) else sub.get('color', 'gray')
                _diagram_box(slide, sx, inner_y, inner_w, inner_h, sub_label, sub_color, font_size=11)


def _draw_diagram_compare(slide, rx, ry, rw, rh, sides):
    """type=compare: 좌우 비교 다이어그램

    sides: [
        {"title": "Before", "items": [...], "color": "red"},
        {"title": "After", "items": [...], "color": "green"},
    ]
    """
    n = len(sides)
    gap = Inches(0.15)
    side_w = (rw - gap * (n - 1)) / n
    pad_x = Inches(0.05)

    for i, side in enumerate(sides):
        sx = rx + i * (side_w + gap)
        fill_c, line_c, text_c = _SEM_BOX_STYLES.get(side.get('color', 'gray'), _SEM_BOX_STYLES['gray'])

        # 헤더
        header_h = Inches(0.45)
        hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx, ry, side_w, header_h)
        hdr.fill.solid(); hdr.fill.fore_color.rgb = line_c
        hdr.line.color.rgb = line_c
        tf = hdr.text_frame; tf.clear()
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.text = side.get('title', '')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(13); p.font.color.rgb = COLORS["BG_WHITE"]
        p.alignment = PP_ALIGN.CENTER

        # 아이템 박스들 (수직 배치)
        items = side.get('items', [])
        if items:
            item_y = ry + header_h + Inches(0.1)
            item_gap = Inches(0.08)
            item_h = (rh - header_h - Inches(0.1) - item_gap * (len(items) - 1)) / len(items)

            for j, item in enumerate(items):
                iy = item_y + j * (item_h + item_gap)
                label = item if isinstance(item, str) else item.get('label', '')
                color = side.get('color', 'gray') if isinstance(item, str) else item.get('color', side.get('color', 'gray'))
                _diagram_box(slide, sx + pad_x, iy, side_w - pad_x * 2, item_h, label, color, font_size=11)


def _draw_diagram_process(slide, rx, ry, rw, rh, steps):
    """type=process: 좌→우 가로 프로세스 (쉐브론 + 설명)

    steps: [
        {"title": "Step 1", "desc": "설명", "color": "gray"},
        ...
    ]
    """
    n = len(steps)
    gap = Inches(0.08)
    step_w = (rw - gap * (n - 1)) / n
    chevron_h = Inches(0.6)

    for i, step in enumerate(steps):
        sx = rx + i * (step_w + gap)
        fill_c, line_c, text_c = _SEM_BOX_STYLES.get(step.get('color', 'gray'), _SEM_BOX_STYLES['gray'])

        # 쉐브론 헤더
        chv = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, sx, ry, step_w, chevron_h)
        chv.fill.solid(); chv.fill.fore_color.rgb = line_c
        chv.line.color.rgb = line_c
        tf = chv.text_frame; tf.clear()
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.text = step.get('title', '')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(10); p.font.color.rgb = COLORS["BG_WHITE"]
        p.alignment = PP_ALIGN.CENTER

        # 설명 박스
        desc = step.get('desc', '')
        if desc:
            desc_y = ry + chevron_h + Inches(0.08)
            desc_h = rh - chevron_h - Inches(0.08)
            _diagram_box(slide, sx, desc_y, step_w, desc_h, desc, step.get('color', 'gray'), font_size=10)


def _draw_right_diagram(slide, rx, ry, rw, rh, diagram_data):
    """우측 다이어그램 라우터 — type에 따라 다른 렌더러 호출

    지원 type:
    - flow: 수직 박스+화살표 흐름도 (기본값)
    - layers: 수평 계층 다이어그램 (아키텍처 티어)
    - compare: 좌우 비교 다이어그램
    - process: 좌→우 가로 프로세스
    """
    # dict 형태 (type 지정)
    if isinstance(diagram_data, dict):
        d_type = diagram_data.get('type', 'flow')
        items = diagram_data.get('items', diagram_data.get('steps', diagram_data.get('layers', diagram_data.get('sides', []))))

        if d_type == 'layers':
            _draw_diagram_layers(slide, rx, ry, rw, rh, items)
        elif d_type == 'compare':
            _draw_diagram_compare(slide, rx, ry, rw, rh, items)
        elif d_type == 'process':
            _draw_diagram_process(slide, rx, ry, rw, rh, items)
        else:
            _draw_diagram_flow(slide, rx, ry, rw, rh, items)

    # list 형태 (하위 호환: flow로 처리)
    elif isinstance(diagram_data, list):
        _draw_diagram_flow(slide, rx, ry, rw, rh, diagram_data)


def render_detail_sections(slide, data):
    """좌측 멀티섹션 텍스트 + 우측 다이어그램/이미지 레이아웃

    KMS PPT 참조: 개요 → 강조 박스(의미 색상) → 조건/불릿 구조
    우측: diagram 데이터 → shape 직접 그리기 / image_path → 이미지 로드
    """
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    gap = Inches(0.3)
    w_left = (bw - gap) * 0.5
    w_right = (bw - gap) * 0.5

    # ── 좌측 콘텐츠 높이 사전 계산 ──
    _ov = content.get('overview', '')
    _hl = content.get('highlight', '')
    _cn = content.get('condition', '')
    overview = _ov if isinstance(_ov, dict) else {'title': '개요', 'body': str(_ov)} if _ov else {}
    highlight = _hl if isinstance(_hl, dict) else {'title': '핵심 성과', 'body': str(_hl)} if _hl else {}
    condition = _cn if isinstance(_cn, dict) else {'title': '적용 조건', 'body': str(_cn)} if _cn else {}

    section_count = sum([1 for s in [overview, highlight, condition] if s])
    if section_count == 0:
        return

    section_gap = Inches(0.12)
    total_gap = section_gap * (section_count - 1)
    available_h = bh - total_gap

    ratios = []
    if overview: ratios.append(('overview', 0.30))
    if highlight: ratios.append(('highlight', 0.45))
    if condition: ratios.append(('condition', 0.25))

    total_ratio = sum(r[1] for r in ratios)
    section_heights = {}
    for name, ratio in ratios:
        section_heights[name] = available_h * (ratio / total_ratio)

    current_y = by

    # (1) 개요 섹션
    if overview:
        sec_h = section_heights['overview']
        tb = slide.shapes.add_textbox(bx, current_y, w_left, sec_h)
        tf = tb.text_frame; tf.word_wrap = True; tf.clear()
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tf.margin_left = Inches(0.1); tf.margin_right = Inches(0.1)
        tf.margin_top = Inches(0.05); tf.margin_bottom = Inches(0.05)
        tf.vertical_anchor = MSO_ANCHOR.TOP

        p = tf.paragraphs[0]
        p.text = overview.get('title', '개요')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(16); p.font.color.rgb = COLORS["DARK_GRAY"]
        p.space_after = Pt(6)

        for line in overview.get('body', '').split('\n'):
            p = tf.add_paragraph()
            p.text = line
            p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(13)
            p.font.color.rgb = COLORS["GRAY"]; p.space_after = Pt(3)

        current_y += sec_h + section_gap

    # (2) 강조 박스 (의미 기반 색상)
    if highlight:
        sec_h = section_heights['highlight']
        color_key = highlight.get('color', 'red')
        sem_colors = {
            'red':    ("SEM_RED", "SEM_RED_BG", "SEM_RED_TEXT"),
            'orange': ("SEM_ORANGE", "SEM_ORANGE_BG", "SEM_ORANGE_TEXT"),
            'green':  ("SEM_GREEN", "SEM_GREEN_BG", "SEM_GREEN_TEXT"),
            'blue':   ("SEM_BLUE", "SEM_BLUE_BG", "SEM_BLUE_TEXT"),
        }
        title_c, bg_c, text_c = sem_colors.get(color_key, sem_colors['red'])

        hl_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, current_y, w_left, sec_h)
        hl_box.fill.solid()
        hl_box.fill.fore_color.rgb = COLORS[bg_c]
        hl_box.line.color.rgb = COLORS[title_c]
        hl_box.line.width = Pt(1.5)

        tf = hl_box.text_frame; tf.clear(); tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tf.margin_left = Inches(0.2); tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.04); tf.margin_bottom = Inches(0.04)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        p = tf.paragraphs[0]
        p.text = highlight.get('title', '')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(12); p.font.color.rgb = COLORS[title_c]
        p.space_after = Pt(3)

        for line in highlight.get('body', '').split('\n'):
            p = tf.add_paragraph()
            p.text = line
            p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(12)
            p.font.color.rgb = COLORS[text_c]; p.space_after = Pt(2)

        current_y += sec_h + section_gap

    # (3) 조건/불릿 섹션
    if condition:
        sec_h = section_heights['condition']
        tb = slide.shapes.add_textbox(bx, current_y, w_left, sec_h)
        tf = tb.text_frame; tf.word_wrap = True; tf.clear()
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tf.margin_left = Inches(0.1); tf.margin_right = Inches(0.1)
        tf.margin_top = Inches(0.05); tf.margin_bottom = Inches(0.05)
        tf.vertical_anchor = MSO_ANCHOR.TOP

        p = tf.paragraphs[0]
        p.text = condition.get('title', '')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(12); p.font.color.rgb = COLORS["SEM_BLUE"]
        p.space_after = Pt(6)

        for bullet in condition.get('bullets', []):
            p = tf.add_paragraph()
            p.text = f"• {bullet}"
            p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(12)
            p.font.color.rgb = COLORS["GRAY"]; p.space_after = Pt(3)

    # ── 우측: diagram shape 또는 이미지 ──
    right_x = bx + w_left + gap
    rendered = False

    # 1순위: diagram 데이터 → shape로 직접 그리기
    diagram = content.get('diagram', [])
    if diagram:
        _draw_right_diagram(slide, right_x, by, w_right, bh, diagram)
        rendered = True

    # 2순위: image_path 직접 지정
    if not rendered:
        image_path = content.get('image_path', '')
        if image_path and os.path.exists(image_path):
            try:
                from PIL import Image as PILImage
                with PILImage.open(image_path) as img:
                    orig_w, orig_h = img.size
                aspect = orig_w / orig_h
                aw, ah = int(w_right), int(bh)
                if aw / aspect <= ah:
                    fw, fh = aw, int(aw / aspect)
                else:
                    fh, fw = ah, int(ah * aspect)
                cx = int(right_x) + (aw - fw) // 2
                cy = int(by) + (ah - fh) // 2
                slide.shapes.add_picture(image_path, cx, cy, width=fw, height=fh)
                rendered = True
            except ImportError:
                slide.shapes.add_picture(image_path, int(right_x), int(by), width=int(w_right), height=int(bh))
                rendered = True
            except Exception as e:
                print(f"   ⚠️ [이미지 로드 실패] {str(e)[:50]}")

    # 3순위: architecture/ 폴더 검색
    if not rendered:
        search_q = content.get('search_q', '')
        if search_q:
            img_file = os.path.join('architecture', search_q.replace(' ', '_') + '.png')
            if os.path.exists(img_file):
                try:
                    from PIL import Image as PILImage
                    with PILImage.open(img_file) as img:
                        orig_w, orig_h = img.size
                    aspect = orig_w / orig_h
                    aw, ah = int(w_right), int(bh)
                    if aw / aspect <= ah:
                        fw, fh = aw, int(aw / aspect)
                    else:
                        fh, fw = ah, int(ah * aspect)
                    cx = int(right_x) + (aw - fw) // 2
                    cy = int(by) + (ah - fh) // 2
                    slide.shapes.add_picture(img_file, cx, cy, width=fw, height=fh)
                    rendered = True
                except:
                    pass

    if not rendered:
        print(f"   ⚠️ [detail_sections] 우측 콘텐츠 없음 — diagram, image_path, 또는 architecture/ 이미지를 지정해주세요")


# 15. Table + Callout (KMS PPT 슬라이드 6 참조)
def render_table_callout(slide, data):
    """비교 테이블 + 하단 추천/결론 콜아웃 박스 레이아웃

    KMS PPT 참조: 상단에 비교 테이블, 하단에 결론/추천 강조 박스
    테이블 열 수 자동 대응 (2~5열)
    """
    wrapper = data.get('data', {})
    y_start = draw_body_header_and_get_y(slide, wrapper.get('body_title'), wrapper.get('body_desc'))
    bx, by, bw, bh = calculate_dynamic_rect(y_start)
    content = wrapper.get('data', {})

    columns = content.get('columns', [])
    rows = content.get('rows', [])
    callout = content.get('callout', {})

    if not columns:
        return

    n_cols = len(columns)

    # 공간 분배: 테이블 65%, 콜아웃 35% (콜아웃 없으면 테이블 100%)
    callout_h = Inches(1.3) if callout else 0
    callout_gap = Inches(0.2) if callout else 0
    table_h = bh - callout_h - callout_gap

    # ── 테이블 영역 ──
    gap = Inches(0.15)
    w_col = (bw - (gap * (n_cols - 1))) / n_cols

    # 헤더 행
    header_h = Inches(0.7)
    for i, col in enumerate(columns):
        x = bx + i * (w_col + gap)
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, by, w_col, header_h)
        shp.fill.solid()
        shp.fill.fore_color.rgb = COLORS["PRIMARY"]
        shp.line.color.rgb = COLORS["PRIMARY"]
        shp.line.width = Pt(1.0)

        tf = shp.text_frame; tf.clear()
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = col.get('title', '') if isinstance(col, dict) else str(col)
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(15); p.font.color.rgb = COLORS["BG_WHITE"]
        p.alignment = PP_ALIGN.CENTER

    # 데이터 행
    if rows:
        row_area_h = table_h - header_h - Inches(0.15)
        row_h = row_area_h / len(rows)

        for row_idx, row in enumerate(rows):
            row_y = by + header_h + Inches(0.15) + (row_idx * row_h)
            values = row if isinstance(row, list) else row.get('values', [])

            for col_idx in range(n_cols):
                x = bx + col_idx * (w_col + gap)
                value = values[col_idx] if col_idx < len(values) else ''

                shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, row_y, w_col, row_h - Inches(0.05))
                shp.fill.solid()
                shp.fill.fore_color.rgb = COLORS["BG_WHITE"]
                shp.line.color.rgb = COLORS["BORDER"]
                shp.line.width = Pt(1.0)

                tf = shp.text_frame; tf.clear()
                tf.margin_left = Inches(0.15); tf.margin_right = Inches(0.15)
                tf.margin_top = Inches(0.08); tf.margin_bottom = Inches(0.08)
                tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE

                p = tf.paragraphs[0]
                p.text = str(value)
                p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(13)
                p.font.color.rgb = COLORS["BLACK"]; p.alignment = PP_ALIGN.CENTER

    # ── 콜아웃 박스 (하단 추천/결론) ──
    if callout:
        callout_y = by + table_h + callout_gap

        # 배경 박스
        cb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, callout_y, bw, callout_h)
        cb.fill.solid()
        cb.fill.fore_color.rgb = COLORS["CALLOUT_BG"]
        cb.line.color.rgb = COLORS["CALLOUT_BG"]

        # 아이콘 (이모지 또는 텍스트)
        icon_text = callout.get('icon', '💡')
        icon_w = Inches(0.7)
        tb_icon = slide.shapes.add_textbox(bx + Inches(0.25), callout_y + Inches(0.15), icon_w, Inches(0.6))
        p = tb_icon.text_frame.paragraphs[0]
        p.text = icon_text
        p.font.size = Pt(30); p.alignment = PP_ALIGN.CENTER

        # 제목 + 본문
        text_x = bx + Inches(1.1)
        text_w = bw - Inches(1.5)

        # 콜아웃 제목
        tb_title = slide.shapes.add_textbox(text_x, callout_y + Inches(0.15), text_w, Inches(0.4))
        p = tb_title.text_frame.paragraphs[0]
        p.text = callout.get('title', '')
        p.font.name = FONTS["BODY_TITLE"]; p.font.bold = True
        p.font.size = Pt(16); p.font.color.rgb = COLORS["BG_WHITE"]

        # 콜아웃 본문
        callout_body = callout.get('body', '')
        if callout_body:
            tb_body = slide.shapes.add_textbox(text_x, callout_y + Inches(0.55), text_w, callout_h - Inches(0.7))
            tf = tb_body.text_frame; tf.word_wrap = True
            for i, line in enumerate(callout_body.split('\n')):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.text = line
                p.font.name = FONTS["BODY_TEXT"]; p.font.size = Pt(14)
                p.font.color.rgb = COLORS["CALLOUT_TEXT"]; p.space_after = Pt(3)


