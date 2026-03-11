# claude-mcp-pptx

MCP 서버로 LLM이 JSON → PPT 직접 생성. Python 스크립트 작성 금지.

## 프로젝트 구조

```
server.py          # MCP 서버 (FastMCP) — 도구 4개 노출
transform.py       # flat JSON → 엔진 포맷 변환
code/              # PPT 렌더링 엔진
  powerpoint_content.py
  powerpoint_cover.py
  powerpoint_toc.py
template/          # PPT 템플릿
icons/             # 아이콘 PNG
results/           # 생성된 .pptx 출력
```

---

## MCP 도구 4개

### `create_presentation` — 신규 생성 / 누적 추가

파일이 없으면 새로 생성, 있으면 같은 파일에 섹션 누적 추가.

```json
{
  "cover_title": "제목\n부제목",
  "cover_subtitle": "부제",
  "sections": [
    {
      "section_title": "1. 섹션명",
      "slides": [
        {
          "layout": "3_cards",
          "title": "1-1. 슬라이드 제목",
          "description": "우측 설명 (선택)",
          "body_title": "본문 제목 (선택)",
          "body_desc": "본문 설명 (선택)",
          "content": {}
        }
      ]
    }
  ],
  "output_name": "파일명"
}
```

> `sections` 누락 시 100% 에러. 한 번에 5-10개 슬라이드씩 나눠서 호출.

---

### `update_slide` — 특정 슬라이드 교체

```json
{
  "output_name": "my_presentation",
  "slide_number": 8,
  "slide_data": { "layout": "3_cards", "title": "...", "content": {} }
}
```

---

### `delete_slide` — 특정 슬라이드 삭제

```json
{
  "output_name": "my_presentation",
  "slide_number": 8
}
```

---

### `insert_slide` — 특정 위치에 삽입

```json
{
  "output_name": "my_presentation",
  "after_slide_number": 8,
  "slide_data": { "layout": "3_cards", "title": "...", "content": {} }
}
```

> `after_slide_number: 0` 이면 맨 앞에 삽입.

---

## 핵심 규칙

### Python 스크립트 금지
- ❌ 루트 또는 `code/` 외 위치에 `.py` 파일 생성
- ❌ `python3 -c` 직접 실행
- ✅ MCP 도구만 사용
- ✅ 디버깅이 꼭 필요한 경우 `debug/` 폴더에만 생성, 작업 후 정리

## 텍스트 오버플로우 방지 — 핵심 원칙

> ⚠️ 렌더러는 텍스트를 자르거나 요약하지 않는다. LLM이 처음부터 한도 내로 작성해야 한다.

### 작성 원칙
- 한 줄 = 한글 기준 약 20자 이내 (영문 35자)
- 불필요한 조사/수식어 제거, 핵심 키워드 중심으로 작성
- 줄 수 초과 시 항목 수를 줄이거나 슬라이드를 분리
- `body_title` / `body_desc`도 1줄로 유지

### 레이아웃별 텍스트 한도

| 레이아웃 | 항목 수 | 줄/항목 | 비고 |
|----------|---------|---------|------|
| `bento_grid` | main 1, sub 2 | main ≤ 6줄, sub ≤ 4줄 | sub는 compact 모드 |
| `3_cards` / `key_metric` | 3개 | body ≤ 6줄 | |
| `grid_2x2` / `quad_matrix` | 4개 | body ≤ 4줄 | compact 모드 |
| `phased_columns` | 3-4개 | body ≤ 4줄 | |
| `timeline_steps` | ≤ 6개 | desc ≤ 2줄 | |
| `zigzag_timeline` | ≤ 6개 | desc ≤ 3줄, title ≤ 10자 | |
| `process_arrow` | ≤ 5개 | body ≤ 3줄 | |
| `comparison_vs` | 2개 | body ≤ 8줄 | |
| `comparison_table` | 정확히 3컬럼 | 셀당 ≤ 2줄 | 행 ≤ 6개 |
| `before_after` | 2개 | body ≤ 6줄 | |
| `pros_cons` / `do_dont` | 각 ≤ 5개 | text ≤ 1줄, detail ≤ 1줄 | |
| `numbered_list` | ≤ 5개 | desc ≤ 2줄 | 5개면 desc 1줄 |
| `stats_dashboard` | ≤ 4개 | desc ≤ 1줄 | |
| `icon_grid` | ≤ 6개 | desc ≤ 2줄 | |
| `exec_summary` | ≤ 5개 섹션 | body ≤ 3줄 | `body_desc` 금지 |
| `split_text_code` | — | 코드 ≤ 14줄, bullets ≤ 5개 | |
| `checklist_2col` | ≤ 8개 항목 | subitem ≤ 4개 | |
| `kanban_board` | 컬럼 3개 | 카드 ≤ 4개/컬럼 | |
| `detail_sections` | 3섹션 | body ≤ 4줄 | |
| `table_callout` | 3컬럼 | 셀당 ≤ 2줄, 행 ≤ 5개 | |
| `pyramid_hierarchy` | ≤ 5레벨 | desc ≤ 1줄 | |
| `fishbone_cause_effect` | ≤ 4카테고리 | causes ≤ 3개 | |
| `org_chart` | ≤ 4children | items ≤ 3개 | |
| `mind_map` | ≤ 5branches | sub_branches ≤ 4개 | |
| `center_radial` | ≤ 6directions | desc ≤ 1줄 | |
| `cycle_loop` | ≤ 6steps | desc ≤ 1줄 | |
| `venn_diagram` | 3circles | desc ≤ 2줄 | |
| `swot_matrix` | 4개 | items ≤ 4개 | |
| `funnel` | ≤ 5stages | label ≤ 8자 | |
| `temple_pillars` | ≤ 4pillars | desc ≤ 2줄 | |
| `quote_highlight` | — | quote ≤ 2줄, 40자/줄 | |
| `architecture_wide` | 3컬럼 | body ≤ 4줄 | |
| `image_left` | — | bullets ≤ 6개 | |



### 슬라이드 제목 형식
- 반드시 `{섹션번호}-{슬라이드번호}. 제목` 형식
- 예: `1-1. 개요`, `2-3. 아키텍처`
- 긴 제목은 `\n`으로 개행 (폰트 크기 줄이지 말 것)

### 레이아웃 다양성
- 같은 레이아웃 최대 2회 사용
- 41개 레이아웃 활용 → `LAYOUTS.md` 참조

### 아이콘 필수
- `icons/` 폴더 아이콘 사용
- 없으면 웹에서 PNG 다운로드 후 `icons/`에 저장
- 파란색 원형 방치 금지

---

## 대용량 생성 전략 (누적 모드)

10슬라이드 이상은 5-10개씩 나눠서 같은 `output_name`으로 반복 호출.

```
1차 호출: 섹션 1 (7개) → 새 파일 생성
2차 호출: 섹션 2 (6개) → 누적
3차 호출: 섹션 3 (6개) → 누적
4차 호출: 섹션 4 (4개) → 누적
```

목차는 자동 병합, 엔딩 슬라이드는 항상 마지막 배치.

---

## 생성 후 체크리스트

- [ ] 슬라이드 수 확인 (표지 + 목차 + 본문 + 엔딩)
- [ ] 목차 항목 수 = 섹션 수
- [ ] 텍스트 오버플로우 없음 (각 슬라이드 줄 수 확인)
- [ ] 아이콘 모두 표시됨 (파란 원형 없음)

---

## 자주 틀리는 키

| Layout | 틀린 것 | 올바른 것 |
|--------|---------|----------|
| `comparison_vs` | `left/right` dict | `item_a_title`, `item_a_body`, `item_b_title`, `item_b_body` |
| `comparison_table` | 4+ columns | 정확히 3개 컬럼 |
| `image_left` | `points[]` | `bullets[]` |
| `timeline_steps` | `steps[].title` | `steps[]{date, desc}` |
| `architecture_wide` | `columns[]` | `col1/col2/col3` 개별 dict |
| `swot_matrix` | `S/W/O/T` dict | `quadrants[]` list |
| `pyramid_hierarchy` | `layers[].title` | `levels[]{label, desc, color}` |
| `checklist_2col` | 커스텀 summary | `"X/Y Passed    Z Warning"` 패턴 필수 |
| `exec_summary` | `body_desc` 사용 | `body_desc` 금지, `body_title`만 사용 |
| `icon_grid` | `search_q` | `icon` 키 사용 |
| `funnel` | `title` | `label`, `value` |
| `quote_highlight` | author에 "—" 포함 | "—" 금지 (렌더러 자동 추가) |

---

## 아이콘 목록 (icons/ 폴더)

analysis, aurora, auto_mode, aws_account, billing, chat, cicd, cli,
cluster_delete, config, console, container, cutover, dashboard, database,
deploy, dms, eks, eksctl, encryption, gitops, helm, k8s_version, kubectl,
kubernetes, load_balancer, microservices, migration, monitoring, network,
performance, pipeline, schema, security, server, service, storage,
terraform, timeline, verification

레이아웃 상세 → `LAYOUTS.md` / 예시 → `EXAMPLES.md`
