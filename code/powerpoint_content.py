# -*- coding: utf-8 -*-
# 라우터 전용 — 실제 렌더러는 각 layouts 파일에 있음
from powerpoint_utils import *
from powerpoint_layouts1 import (
    render_3_cards, render_bento_grid, render_grid_2x2, render_quad_matrix,
    render_challenge_solution, render_timeline_steps, render_process_arrow,
    render_phased_columns, render_architecture_wide, render_image_left,
    render_comparison_vs, render_key_metric, render_detail_image,
    render_comparison_table, render_detail_sections, render_table_callout,
)
from powerpoint_layouts2 import (
    render_full_image, render_before_after, render_icon_grid,
    render_numbered_list, render_stats_dashboard, render_quote_highlight,
    render_pros_cons, render_do_dont, render_split_text_code,
    render_pyramid_hierarchy, render_cycle_loop, render_venn_diagram,
    render_swot_matrix, render_center_radial, render_funnel,
    render_zigzag_timeline,
)
from powerpoint_layouts3 import (
    render_fishbone_cause_effect, render_org_chart, render_temple_pillars,
    render_infinity_loop, render_mind_map, render_checklist_2col,
    render_kanban_board, render_exec_summary, render_risk_table,
    render_speedometer_gauge,
)


def render_slide_content(slide, layout, data):
    clean_body_placeholders(slide)

    renderers = {
        "bento_grid": render_bento_grid, "3_cards": render_3_cards,
        "grid_2x2": render_grid_2x2, "quad_matrix": render_quad_matrix,
        "timeline_steps": render_timeline_steps, "process_arrow": render_process_arrow, "phased_columns": render_phased_columns,
        "architecture_wide": render_architecture_wide, "image_left": render_image_left,
        "comparison_vs": render_comparison_vs, "key_metric": render_key_metric,
        "challenge_solution": render_challenge_solution, "detail_image": render_detail_image,
        "comparison_table": render_comparison_table,
        "detail_sections": render_detail_sections, "table_callout": render_table_callout,
        "full_image": render_full_image, "before_after": render_before_after,
        "icon_grid": render_icon_grid, "numbered_list": render_numbered_list,
        "stats_dashboard": render_stats_dashboard,
        "quote_highlight": render_quote_highlight, "pros_cons": render_pros_cons,
        "do_dont": render_do_dont, "split_text_code": render_split_text_code,
        "pyramid_hierarchy": render_pyramid_hierarchy, "cycle_loop": render_cycle_loop,
        "venn_diagram": render_venn_diagram, "swot_matrix": render_swot_matrix,
        "center_radial": render_center_radial, "funnel": render_funnel,
        "zigzag_timeline": render_zigzag_timeline, "fishbone_cause_effect": render_fishbone_cause_effect,
        "org_chart": render_org_chart, "temple_pillars": render_temple_pillars,
        "infinity_loop": render_infinity_loop,
        "mind_map": render_mind_map,
        "checklist_2col": render_checklist_2col,
        "kanban_board": render_kanban_board,
        "risk_table": render_risk_table,
        "exec_summary": render_exec_summary,
        "speedometer_gauge": render_speedometer_gauge,
    }

    func = renderers.get(layout)
    if func:
        try: func(slide, data)
        except Exception as e: create_content_box(slide, Inches(1), Inches(3), Inches(10), Inches(2), "Error", str(e))
    else:
        create_content_box(slide, Inches(1), Inches(3), Inches(10), Inches(2), "Layout Not Found", str(data))
