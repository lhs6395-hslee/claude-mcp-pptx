# -*- coding: utf-8 -*-
"""
flat MCP 입력 → 엔진의 data.data.data 형식 변환

flat (Claude → MCP):
    {"layout": "...", "title": "...", "description": "...",
     "body_title": "...", "body_desc": "...", "content": {...}}

engine (powerpoint_content.py):
    {"l": "...", "t": "...", "d": "...",
     "data": {"body_title": "...", "body_desc": "...", "data": {...}}}

예외: challenge_solution, before_after는 content를 wrapper 레벨에 병합
"""

WRAPPER_LEVEL_LAYOUTS = {"challenge_solution", "before_after"}


def flat_to_engine_format(slide: dict) -> dict:
    """flat slide dict → engine format dict 변환"""
    layout = slide.get("layout", "bento_grid")
    content = slide.get("content", {})

    if layout in WRAPPER_LEVEL_LAYOUTS:
        wrapper = {
            "body_title": slide.get("body_title", ""),
            "body_desc": slide.get("body_desc", ""),
            **content,
        }
    else:
        wrapper = {
            "body_title": slide.get("body_title", ""),
            "body_desc": slide.get("body_desc", ""),
            "data": content,
        }

    return {
        "l": layout,
        "t": slide.get("title", ""),
        "d": slide.get("description", ""),
        "data": wrapper,
    }
