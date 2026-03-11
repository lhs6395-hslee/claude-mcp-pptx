# claude-mcp-pptx

mcp-pptx PowerPoint 생성 엔진을 MCP(Model Context Protocol) 서버로 래핑한 프로젝트.

Claude가 Python 스티어링 파일을 작성하는 대신 JSON 데이터를 직접 MCP 도구로 전달해 PPT를 생성합니다.

## 핵심 차이점

| | 기존 (mcp-pptx) | MCP (claude-mcp-pptx) |
|---|---|---|
| 입력 | Python .py 파일 작성 | JSON 데이터 직접 전달 |
| data 중첩 | data.data.data (3단계) | content (1단계, 서버가 변환) |
| 실행 | `python3 generate.py file.py` | MCP 도구 호출 |
| 출력 | 동일 | 동일 |
| 렌더링 엔진 | powerpoint_content.py | 동일 (import) |

## 구조

```
claude-mcp-pptx/              ← 독립 실행 가능 (mcp-pptx 불필요)
  server.py                    MCP 서버 (FastMCP, stdio transport)
  transform.py                 flat JSON → engine data.data.data 변환
  code/
    powerpoint_content.py      41개 렌더러
    powerpoint_cover.py        표지 생성
    powerpoint_toc.py          목차 생성
  template/
    2025_PPT_Template_FINAL.pptx
  icons/                       아이콘 이미지 (47개)
  screenshots/                 스크린샷 이미지
  results/                     생성된 PPT 출력
  CLAUDE.md                    Claude Code 자동 로딩 컨텍스트
  LAYOUTS.md                   41개 레이아웃별 content 형식 레퍼런스
  EXAMPLES.md                  프롬프트 예시 + MCP 호출 예시
```

## 설치

```bash
pip3 install -r requirements.txt
```

## Claude Code 등록

```bash
claude mcp add --transport stdio pptx-generator -- python3 /path/to/claude-mcp-pptx/server.py
```

또는 `.kiro/settings/mcp.json`에 직접 추가:

```json
{
  "mcpServers": {
    "pptx-generator": {
      "command": "python3",
      "args": ["/Users/rayhli/Documents/ide/claude-mcp-pptx/server.py"],
      "disabled": false
    }
  }
}
```

## MCP 도구

### `create_presentation`

| 파라미터 | 타입 | 필수 | 설명 |
|---|---|---|---|
| cover_title | string | O | 표지 제목 (`\n`으로 줄바꿈) |
| cover_subtitle | string | O | 표지 부제 |
| sections | array | O | 섹션 배열 (아래 참조) |
| output_name | string | X | 출력 파일명 (확장자 없이, 기본: 타임스탬프) |

**section 구조:**
```json
{
  "section_title": "1. 섹션 제목",
  "slides": [
    {
      "layout": "레이아웃명",
      "title": "슬라이드 헤더 (좌)",
      "description": "슬라이드 헤더 (우)",
      "body_title": "본문 제목 (선택)",
      "body_desc": "본문 설명 (선택)",
      "content": { ... }
    }
  ]
}
```

**content 형식**: 레이아웃별로 다름 → `LAYOUTS.md` 참조

**출력**: 생성된 PPTX 절대 경로 (예: `/Users/.../claude-mcp-pptx/results/output.pptx`)

## 41개 레이아웃

### 카드/그리드 계열
- `3_cards`, `bento_grid`, `grid_2x2`, `quad_matrix`, `key_metric`

### 타임라인/프로세스 계열
- `timeline_steps`, `zigzag_timeline`, `process_arrow`, `phased_columns`, `cycle_loop`, `infinity_loop`

### 비교/대조 계열
- `comparison_vs`, `challenge_solution`, `before_after`, `pros_cons`, `do_dont`

### 테이블 계열
- `comparison_table`, `table_callout`, `risk_table`

### 이미지/다이어그램 계열
- `detail_image`, `image_left`, `full_image`, `architecture_wide`

### 다이어그램 계열
- `pyramid_hierarchy`, `venn_diagram`, `swot_matrix`, `center_radial`, `funnel`, `fishbone_cause_effect`, `org_chart`, `temple_pillars`, `mind_map`, `speedometer_gauge`

### 텍스트/리스트 계열
- `numbered_list`, `stats_dashboard`, `quote_highlight`, `icon_grid`, `split_text_code`, `exec_summary`

### 체크리스트/보드 계열
- `checklist_2col`, `kanban_board`

### 복합 레이아웃
- `detail_sections`

## 아이콘 (47개)

```
analysis, analytics, aurora, auto_mode, availability, aws_account, billing, 
chat, cicd, cli, cloud, cluster_delete, config, console, container, cutover, 
dashboard, database, deploy, dms, eks, eksctl, encryption, gitops, helm, 
iot, k8s_version, kafka, kubectl, kubernetes, lambda, load_balancer, 
microservices, migration, monitoring, network, performance, pipeline, 
scale, schema, security, server, service, storage, streaming, terraform, 
timeline, verification
```

- Format: PNG, 512×512 pixels, transparent background
- Naming: lowercase, underscores for spaces
- Location: `icons/` folder
- Fallback: 아이콘 파일 없으면 파란색 원형 표시

## 디자인 시스템

### 폰트
- HEAD_TITLE: "프리젠테이션 7 Bold" (28pt)
- HEAD_DESC: "프리젠테이션 5 Medium" (12pt)
- BODY_TITLE/TEXT: "Freesentation"

### 색상
- PRIMARY: RGB(0, 67, 218) - 제목, 강조
- BLACK: RGB(0, 0, 0) - 본문
- GRAY: RGB(80, 80, 80) - 설명
- BG_BOX: RGB(248, 249, 250) - 박스 배경

### 레이아웃 좌표
- SLIDE_TITLE_Y: 0.6"
- BODY_START_Y: 2.0"
- BODY_LIMIT_Y: 7.2"
- MARGIN_X: 0.5"
- SLIDE_W: 13.333"

## 템플릿 요구사항

PPT 템플릿은 다음 슬라이드 구조를 가져야 함:
- Index 0: Cover slide (표지)
- Index 1: TOC slide (목차)
- Index 7: Body slide (본문 - 이 레이아웃을 복제)
- Last slide: Ending slide (감사합니다 - 보존됨)
- Slide dimensions: 13.333" × 7.500"

## 주요 규칙

1. **Python 스크립트 생성 금지** - MCP 도구 직접 호출
2. **텍스트 박스 넘침 금지** - 텍스트는 간결하게, 코드는 슬라이드 분할
3. **슬라이드 제목 번호 필수** - "섹션번호-슬라이드번호. 제목" 형식
4. **레이아웃 다양성 필수** - 동일 레이아웃 최대 2번까지만
5. **아이콘 필수** - 파란색 원형 금지, 없으면 다운로드

자세한 내용은 `CLAUDE.md`, `LAYOUTS.md`, `EXAMPLES.md` 참조.
