# 레이아웃 레퍼런스 (41개)

`create_presentation` MCP 도구의 `content` 필드에 들어가는 레이아웃별 데이터 형식.

> **content는 flat 형식입니다.** 기존 data.data.data 3중 중첩 불필요 — MCP 서버가 자동 변환합니다.

## ⚠️ 중요: 슬라이드 제목 형식

**모든 슬라이드 제목에는 반드시 "섹션번호-슬라이드번호. 제목" 형식 사용:**

```json
{
  "section_title": "1. MSK 개요",
  "slides": [
    {
      "layout": "3_cards",
      "title": "1-1. AWS MSK란?",  // ✅ 올바른 형식
      "description": "Apache Kafka 완전 관리형 서비스",
      "content": {...}
    }
  ]
}
```

**잘못된 예:**
- ❌ "AWS MSK란?" (번호 없음)
- ❌ "1. AWS MSK란?" (슬라이드 번호 누락)

---

## 카드/그리드 계열

### 3_cards
```json
{"card_1": {"title": "", "body": "", "search_q": ""}, "card_2": {...}, "card_3": {...}}
```

### bento_grid
```json
{"main": {"title": "", "body": "", "search_q": ""}, "sub1": {"title": "", "body": ""}, "sub2": {"title": "", "body": ""}}
```

### grid_2x2
```json
{"item1": {"title": "", "body": ""}, "item2": {...}, "item3": {...}, "item4": {...}}
```

### quad_matrix
grid_2x2와 동일.

### key_metric
3_cards와 동일.

---

## 타임라인/프로세스 계열

### timeline_steps
```json
{"steps": [{"date": "Week 1", "desc": "설명"}]}
```
> `date` 키 사용 (title 아님)

### zigzag_timeline
```json
{"steps": [{"date": "03/03", "title": "킥오프", "desc": "설명"}]}
```
> MM/DD 형식이면 과거=파랑, 미래=회색 자동 적용
> **⚠️ 텍스트 제약 (카드 2.0×1.3in):**
> - `desc` 1줄 최대 ~16자 (한글 9pt), 최대 4줄
> - `title` 최대 ~12자 (한글 12pt)
> - 이모지 절대 금지 — 카드 넘침 원인

### process_arrow
```json
{"steps": [{"title": "", "body": "", "search_q": ""}]}
```

### phased_columns
```json
{"steps": [{"title": "1. 분석", "body": "현황 파악", "search_q": "analysis"}]}
```

### cycle_loop
```json
{"center_label": "DevOps", "steps": [{"label": "Plan", "desc": ""}, {"label": "Build"}, ...]}
```

### infinity_loop
```json
{"left_label": "Dev", "right_label": "Ops", "center_label": "CI/CD",
 "left_loop": [{"label": "Plan"}, {"label": "Code"}],
 "right_loop": [{"label": "Deploy"}, {"label": "Monitor"}]}
```

---

## 비교/대조 계열

### comparison_vs
```json
{"item_a_title": "옵션 A", "item_a_body": "설명", "item_b_title": "옵션 B", "item_b_body": "설명"}
```
> flat 키 (nested dict 아님)

### challenge_solution (**예외: wrapper 레벨**)
```json
{"challenge": {"title": "CHALLENGE", "body": "문제 설명"}, "solution": {"title": "SOLUTION", "body": "해결 방안"}}
```

### before_after (**예외: wrapper 레벨**)
```json
{"before_title": "Before", "before_body": "기존 상태", "after_title": "After", "after_body": "개선 결과"}
```

### pros_cons
```json
{"subject": "클라우드 마이그레이션", "pros": ["비용 절감", "확장성"], "cons": ["초기 비용", "학습 곡선"]}
```

### do_dont
```json
{"do_items": [{"text": "파라미터 쿼리 사용", "detail": "SQL 인젝션 방지"}],
 "dont_items": [{"text": "하드코딩 금지", "detail": "보안 위험"}]}
```

---

## 테이블 계열

### comparison_table
```json
{"columns": [{"title": "항목"}, {"title": "Plan A"}, {"title": "Plan B"}],
 "rows": [["가격", "$10", "$20"], ["속도", "빠름", "더 빠름"]]}
```
> **정확히 3개 컬럼** 필수

### table_callout
```json
{"columns": ["항목", "옵션A", "옵션B"],
 "rows": [["속도", "100ms", "50ms"]],
 "callout": {"icon": "💡", "title": "권장", "body": "옵션B 선택"}}
```

### risk_table
```json
{"summary": "⬤ 2 Yellow   |   ⬤ 3 Red",
 "columns": ["상태", "항목", "설명", "담당자"],
 "rows": [{"level": "critical", "item": "데이터 유실", "desc": "백업 없음", "owner": "DBA"},
          {"level": "high", "item": "성능 저하", "desc": "인덱스 미적용", "owner": "DB"}]}
```
> level 색상: `critical` → 빨강, `high` → **노랑**, `orange` → 주황
> summary 텍스트와 실제 circle 색상 일치 필수

---

## 이미지/다이어그램 계열

### detail_image
```json
{"title": "아키텍처 개요", "body": "설명 텍스트", "search_q": "aws_architecture"}
```

### image_left
```json
{"image_path": "screenshots/app.png", "bullets": ["포인트 1", "포인트 2"]}
```
> `bullets[]` 사용 (points 아님)

### full_image
```json
{"image_path": "architecture/diagram.png", "caption": "Figure 1: 시스템 아키텍처"}
```

### architecture_wide
```json
{"diagram_path": "architecture/diagram.png",
 "col1": {"title": "레이어 1", "body": "설명", "search_q": "database"},
 "col2": {"title": "레이어 2", "body": "설명"},
 "col3": {"title": "레이어 3", "body": "설명"}}
```
> col1/col2/col3 개별 dict (columns[] 아님)

---

## 다이어그램 계열

### pyramid_hierarchy
```json
{"levels": [{"label": "전략", "desc": "비전", "color": "primary"},
            {"label": "전술", "desc": "실행"},
            {"label": "운영", "desc": "일상"}]}
```
> 위→아래 순서 (첫 번째=꼭대기)

### venn_diagram
```json
{"circles": [{"label": "보안", "desc": "암호화", "color": "blue"},
             {"label": "성능", "desc": "저지연", "color": "red"},
             {"label": "비용", "desc": "최적화", "color": "green"}],
 "center_label": "균형"}
```

### swot_matrix
```json
{"quadrants": [
  {"label": "S", "title": "Strengths", "items": ["기술력"], "color": "blue"},
  {"label": "W", "title": "Weaknesses", "items": ["인력 부족"], "color": "red"},
  {"label": "O", "title": "Opportunities", "items": ["시장 성장"], "color": "green"},
  {"label": "T", "title": "Threats", "items": ["경쟁"], "color": "orange"}
]}
```
> quadrants[] 리스트 (S/W/O/T dict 아님)

### center_radial
```json
{"center": {"label": "핵심 전략", "desc": "디지털 전환"},
 "directions": [{"label": "기술", "desc": "클라우드", "color": "blue"},
                {"label": "프로세스", "desc": "자동화", "color": "green"},
                {"label": "인력", "desc": "교육", "color": "orange"},
                {"label": "문화", "desc": "혁신", "color": "red"}]}
```

### funnel
```json
{"stages": [{"label": "리드", "value": "1000", "color": "blue"},
            {"label": "검증", "value": "300", "color": "green"},
            {"label": "계약", "value": "50", "color": "primary"}]}
```

### fishbone_cause_effect
```json
{"effect": "마이그레이션 실패",
 "categories": [{"label": "인력", "causes": ["교육 부족", "저항"], "color": "blue"},
                {"label": "프로세스", "causes": ["롤백 계획 없음"], "color": "green"}]}
```

### org_chart
```json
{"root": {"label": "CTO", "desc": "기술총괄"},
 "children": [{"label": "개발팀", "desc": "제품개발", "items": ["FE", "BE", "QA"], "color": "blue"},
              {"label": "운영팀", "desc": "인프라", "items": ["SRE", "DBA"], "color": "green"}]}
```

### temple_pillars
```json
{"roof": {"label": "비즈니스 가치"},
 "pillars": [{"label": "보안", "desc": "암호화", "color": "blue"},
             {"label": "성능", "desc": "저지연", "color": "green"},
             {"label": "안정성", "desc": "HA/DR", "color": "orange"}],
 "foundation": {"label": "클라우드 인프라"}}
```

### mind_map
```json
{"center": {"label": "클라우드 마이그레이션"},
 "branches": [{"label": "컴퓨팅", "sub_branches": ["EC2", "Lambda"], "color": "blue"},
              {"label": "스토리지", "sub_branches": ["S3", "EBS"], "color": "green"}]}
```

### speedometer_gauge
```json
{"value": 75, "title": "마이그레이션 진척률",
 "segments": [{"label": "위험", "color": "red"}, {"label": "주의", "color": "orange"},
              {"label": "양호", "color": "green"}]}
```

---

## 텍스트/리스트 계열

### numbered_list
```json
{"items": [{"title": "1단계: 계획", "desc": "범위 정의\n리스크 식별"}, {"title": "2단계: 실행"}]}
```

### stats_dashboard
```json
{"metrics": [{"value": "99.9", "unit": "%", "label": "가용성", "desc": "연간 SLA"},
             {"value": "50", "unit": "ms", "label": "레이턴시"}]}
```

### quote_highlight
```json
{"quote": "미래를 예측하는 가장 좋은 방법은 직접 만드는 것이다.",
 "author": "Peter Drucker", "role": "경영 컨설턴트"}
```

### icon_grid
```json
{"items": [{"icon": "kubernetes", "title": "K8s", "desc": "컨테이너 오케스트레이션"},
           {"icon": "docker", "title": "Docker", "desc": "컨테이너화"}]}
```

### split_text_code
```json
{"description": "배포 자동화 스크립트",
 "bullets": ["빠른 배포", "안정적"],
 "code_title": "deploy.sh",
 "lang": "bash",
 "code": "#!/bin/bash\nkubectl apply -f manifest.yaml"}
```
> `lang` 옵션: `"yaml"` | `"python"` | `"bash"` | `"sh"` — 생략 시 단색(초록)

### exec_summary
```json
{"sections": [{"label": "상황", "body": "프로젝트 정상 진행 중", "color": "gray"},
              {"label": "핵심 발견사항", "body": "■  발견 1\n    상세 내용 1줄만\n■  발견 2\n    상세 내용 1줄만", "color": "blue"},
              {"label": "권고사항", "body": "➤  권고 1\n➤  권고 2\n➤  권고 3", "color": "navy"}]}
```
> body에 `■` → 불릿(heading+body 1줄 쌍), `➤` → 화살표 스타일 자동 적용
>
> **⚠️ 중요 제약:**
> - `body_title`만 사용, **`body_desc` 절대 금지** (레이아웃 깨짐)
> - `bullet_plain` (■): ■당 heading 1줄 + body **1줄만** 렌더링 (추가 줄 무시됨)
> - `bullet_arrow` (➤): 서브 들여쓰기 줄도 카운트됨 — 총 3줄 이하 권장
> - 3개 섹션(inline·bullet_plain·bullet_arrow) 기준: ■ 3개 + ➤ 3줄 = 압축 없음(scale=1.0)

---

## 체크리스트/보드 계열

### checklist_2col
```json
{"summary": "1/10 Passed    9 Warning",
 "items": [{"title": "WBS 1.1 Setup", "status": "done",
            "subitems": [{"text": "DB 설치", "badge": ""}, {"text": "복제 구성", "badge": "HIGH"}]}]}
```
> status: `done` | `in_progress` | `todo`
> **summary는 반드시 `"X/Y Passed    Z Warning"` 패턴** — 렌더러가 파싱해 progress bar 생성. 커스텀 텍스트 사용 시 bar 깨짐.
> subitem text 1줄 권장 (sub_row_h=0.30in 고정)

### kanban_board
```json
{"columns": [{"title": "To Do (3)", "color": "navy", "cards": [{"title": "Task 1\n03/10", "badge": "Critical"}]},
             {"title": "In Progress (2)", "color": "blue", "cards": [{"title": "Task 2"}]},
             {"title": "Done (5)", "color": "green", "cards": []}]}
```

---

## 상세: detail_sections (복합 레이아웃)

왼쪽 3섹션 + 오른쪽 다이어그램/이미지:

```json
{"overview": {"title": "개요", "body": "요약 텍스트"},
 "highlight": {"title": "핵심 발견", "body": "중요\n상세", "color": "red"},
 "condition": {"title": "조건", "bullets": ["조건 1", "조건 2"]},
 "diagram": {"type": "flow", "items": [
   {"label": "Step 1", "color": "blue"},
   {"type": "arrow", "label": ""},
   {"label": "Step 2", "color": "green"}
 ]}}
```

다이어그램 type: `flow`(기본) | `layers` | `compare` | `process`
다이어그램 대신 `image_path` 또는 `search_q` 사용 가능.
