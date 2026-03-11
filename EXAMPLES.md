# 예시: MCP 도구로 프레젠테이션 생성

## 1. 사용자 프롬프트 예시

### 간단한 요청
```
DB 마이그레이션 1주차 계획 PPT를 만들어줘.
- 전체 로드맵 (5 Phase)
- 1주차 일별 일정
- 핵심 지표 (KPI 4개)
- Phase 1 체크리스트
- 칸반 보드
- 리스크 테이블
- 목표 & Next Step
```

**중요: 슬라이드 제목에는 반드시 "섹션번호-슬라이드번호. 제목" 형식 사용!**

### 구체적인 요청
```
다음 내용으로 PPT를 만들어줘. 파일명은 ssts_1juca_plan으로.

표지: DB 마이그레이션 1주차 실행 계획
부제: 신성통상 GWM — Phase 1 분석 + Phase 2 인프라 착수

섹션 1: Roadmap & 핵심 지표
- zigzag_timeline: 프로젝트 전체 5단계 로드맵
- zigzag_timeline: 1주차 일별 로드맵 (03/03~06)
- stats_dashboard: 핵심 KPI 4개

섹션 2: 체크리스트 & 작업 현황판
- checklist_2col: Phase 1 작업 체크리스트
- kanban_board: To Do / Doing / Done
- risk_table: 리스크 5건

섹션 3: 목표 & Next Step
- exec_summary: 상황/핵심발견/권고
```

---

## 2. MCP 도구 호출 예시

아래는 위 요청을 Claude가 `create_presentation` MCP 도구로 변환한 실제 호출 예시입니다.

```json
{
  "cover_title": "DB 마이그레이션\n1주차 실행 계획",
  "cover_subtitle": "신성통상 GWM — Phase 1 분석 + Phase 2 인프라 착수 (03/03 ~ 03/06)",
  "output_name": "ssts_1juca_plan",
  "sections": [
    {
      "section_title": "1. Roadmap & 핵심 지표",
      "slides": [
        {
          "layout": "zigzag_timeline",
          "title": "1-1. 프로젝트 전체 로드맵",
          "description": "5단계 실행 계획 — 현재 Phase 1·2 진행중",
          "body_title": "프로젝트 재개 로드맵 (2026.03.03 ~ 04.03)",
          "body_desc": "WBS 기준 5단계, 총 1개월 — D-Day 03/26(목) 컷오버",
          "content": {
            "steps": [
              {
                "date": "03/03~09",
                "title": "Phase 1: 분석",
                "desc": "킥오프·R&R 확인\nGap 분석·오브젝트 현행화\n호환성 재검증 (PG14→17)\nAurora 파라미터 확인"
              },
              {
                "date": "03/03~20",
                "title": "Phase 2: 운영 인프라 구축",
                "desc": "Aurora/DMS 인프라 구성\n보안·네트워크 점검\n스키마 사전 이관 (DDL)\nDatadog 대시보드 구성"
              },
              {
                "date": "03/10~24",
                "title": "Phase 3: 이관 전 테스트",
                "desc": "통테 DMS Full Load\n운영계정 DMS+CDC 테스트\n개발팀 통합테스트\n실제 운영 Full Load+CDC"
              },
              {
                "date": "03/26",
                "title": "Phase 4: 컷오버",
                "desc": "DMS Full Load (PK 미존재)\nGlue 암호화 35테이블 적재\nFK/Trigger·정합성 검증\n서비스 전환 (약 4~5h)"
              },
              {
                "date": "03/26~04/03",
                "title": "Phase 5: 안정화",
                "desc": "앱 정상 동작 확인\nDatadog 집중 모니터링\n종료 보고·인수인계"
              }
            ]
          }
        },
        {
          "layout": "stats_dashboard",
          "title": "1-2. 핵심 지표",
          "description": "WBS (재개) 기준 1주차 작업 규모",
          "body_title": "Week 1 Key Metrics",
          "body_desc": "Phase 1 분석 단계 — WBS 1.1~1.5 해당",
          "content": {
            "metrics": [
              {"value": "10", "unit": "Task", "label": "1주차 활성 Task", "desc": "Phase 1: 1.1~1.5 (5개) + Phase 2: 2.1~2.5 (5개)"},
              {"value": "28", "unit": "건", "label": "전체 WBS Task", "desc": "1.0 분석(5) + 2.0 인프라·이관(15) + 3.0 컷오버·안정화(8)"},
              {"value": "1", "unit": "개월", "label": "프로젝트 기간", "desc": "03/03(화) ~ 04/03(금) — D-Day 03/26"},
              {"value": "1", "unit": "완료", "label": "1.1 킥오프 완료", "desc": "03/04 완료 | 나머지 9개 Task 진행중·예정"}
            ]
          }
        }
      ]
    },
    {
      "section_title": "2. 체크리스트 & 작업 현황판",
      "slides": [
        {
          "layout": "checklist_2col",
          "title": "2-1. Phase 1 작업 체크리스트",
          "description": "WBS Phase 1 세부 항목 (1주차 해당분)",
          "body_title": "Phase 1 분석 — 1주차 Task 목록",
          "body_desc": "WBS (재개) 기준 — 1.1~1.5 상세",
          "content": {
            "summary": "1/10 Passed    9 Warning",
            "items": [
              {
                "title": "WBS 1.1 재개 킥오프 ✅ 완료",
                "status": "done",
                "subitems": [
                  {"text": "재개 범위·일정 합의, R&R 재확인", "badge": "CRITICAL"},
                  {"text": "중단 기간 변경사항 브리핑", "badge": "HIGH"}
                ]
              },
              {
                "title": "WBS 1.2 시스템 변경 내역 추적",
                "status": "in_progress",
                "subitems": [
                  {"text": "스키마/데이터 변경사항 전수 파악", "badge": "CRITICAL"},
                  {"text": "변경 내역 추적 결과 보고서", "badge": "HIGH"}
                ]
              },
              {
                "title": "WBS 1.3 오브젝트 현황 반영",
                "status": "in_progress",
                "subitems": [
                  {"text": "Table, Index, View 최신화", "badge": "CRITICAL"},
                  {"text": "오브젝트 현황표 업데이트", "badge": "HIGH"}
                ]
              }
            ]
          }
        },
        {
          "layout": "kanban_board",
          "title": "2-2. 작업 현황판",
          "description": "WBS Task 기준 Kanban Board",
          "body_title": "Week 1 Task Board",
          "body_desc": "Trello 기준 (03/05 현재)",
          "content": {
            "columns": [
              {
                "title": "To Do (3)", "color": "navy",
                "cards": [
                  {"title": "2.4 스키마 사전 이관\n03/06~09, DB, 2일", "badge": "내일 착수"},
                  {"title": "2.5 Datadog 모니터링\n03/06~20, Infra, 11일", "badge": "내일 착수"}
                ]
              },
              {
                "title": "Doing (7)", "color": "blue",
                "cards": [
                  {"title": "1.2 시스템 변경 내역 추적\n03/03~06, DB, 4일", "badge": "Critical"},
                  {"title": "1.3 오브젝트 현황 반영\n03/03~06, DB, 4일", "badge": "Critical"}
                ]
              },
              {
                "title": "Done (1)", "color": "green",
                "cards": [
                  {"title": "1.1 재개 킥오프 미팅\n03/04, 전체, 1일", "badge": "완료"}
                ]
              }
            ]
          }
        },
        {
          "layout": "risk_table",
          "title": "2-3. 리스크 & 전제 조건",
          "description": "WBS Risk/Note 기반 사전 점검",
          "body_title": "사전 점검 항목",
          "body_desc": "WBS Risk/Note 컬럼 + 비고 사항 기반",
          "content": {
            "summary": "⬤ 4 Yellow   |   ⬤ 1 Red",
            "columns": ["상태", "항목", "설명", "담당자"],
            "rows": [
              {"level": "high", "item": "시스템 변경 내역 추적 범위", "desc": "중단 기간 중 스키마/데이터 변경사항 전수 파악, 변경 폭에 따라 후속 일정 영향", "owner": "DB"},
              {"level": "critical", "item": "컷오버 다운타임", "desc": "서비스 컷오버(03/26) 최대 5시간 소요, 야간 작업 사전 승인 필요", "owner": "전체"}
            ]
          }
        }
      ]
    },
    {
      "section_title": "3. 목표 & Next Step",
      "slides": [
        {
          "layout": "exec_summary",
          "title": "3-1. 목표 & Next Step",
          "description": "1주차 완료 기준 및 2주차 연계",
          "body_title": "Executive Summary",
          "content": {
            "sections": [
              {
                "label": "상황",
                "body": "프로젝트 재개 1주차 (03/03~06). 전체 WBS 28개 Task 중 10개 착수.",
                "color": "gray"
              },
              {
                "label": "핵심 발견사항",
                "body": "■  1주차 완료 기준\n    1.1 킥오프 완료 + 1.2~1.4 분석 완료 예정 + 2.1 인프라 구성 완료\n■  핵심 리스크\n    1.2 변경 내역 추적 시 스키마 변경 폭이 클 경우 후속 일정 영향",
                "color": "blue"
              },
              {
                "label": "권고사항",
                "body": "➤  03/06(금) 주간 리뷰: 1.2~1.4 완료 결과 공유, 2주차 계획 확정\n➤  개발팀 Task 일정 사전 협의\n➤  컷오버(03/26) 야간 작업 사전 승인",
                "color": "navy"
              }
            ]
          }
        }
      ]
    }
  ]
}
```

---

## 3. 기존 Python 방식과의 비교

### Before (mcp-pptx — Python steering file)

```python
# data.data.data 3중 중첩 필요
{
    "l": "zigzag_timeline",
    "t": "1-1. 전체 로드맵",
    "d": "5단계 실행 계획",
    "data": {
        "body_title": "...",
        "body_desc": "...",
        "data": {
            "steps": [{"date": "03/03", "title": "Phase 1", "desc": "..."}]
        }
    }
}
```

### After (claude-mcp-pptx — flat JSON)

```json
{
    "layout": "zigzag_timeline",
    "title": "1-1. 전체 로드맵",
    "description": "5단계 실행 계획",
    "body_title": "...",
    "body_desc": "...",
    "content": {
        "steps": [{"date": "03/03", "title": "Phase 1", "desc": "..."}]
    }
}
```

**차이점:**
- `l/t/d` → `layout/title/description` (읽기 쉬운 키)
- `data.data.data` 3중 중첩 → `content` 1단계 (서버가 변환)
- Python 파일 작성 불필요 → JSON 직접 전달
- PRESENTATION_GUIDE.md 로딩 불필요 → 토큰 ~70% 절감

---

## 4. 레이아웃별 간단 예시

### 3_cards
```json
{
  "layout": "3_cards",
  "title": "핵심 영역",
  "content": {
    "card_1": {"title": "보안", "body": "암호화 적용", "search_q": "security"},
    "card_2": {"title": "성능", "body": "레이턴시 최적화", "search_q": "performance"},
    "card_3": {"title": "안정성", "body": "HA/DR 구성", "search_q": "reliability"}
  }
}
```

### comparison_vs
```json
{
  "layout": "comparison_vs",
  "title": "방안 비교",
  "content": {
    "item_a_title": "방안 A: DMS",
    "item_a_body": "AWS 네이티브\n자동 CDC 지원\n비용 효율적",
    "item_b_title": "방안 B: Debezium",
    "item_b_body": "오픈소스\n커스텀 가능\n학습 비용 높음"
  }
}
```

### process_arrow
```json
{
  "layout": "process_arrow",
  "title": "마이그레이션 절차",
  "content": {
    "steps": [
      {"title": "분석", "body": "Gap 분석\n호환성 검증"},
      {"title": "인프라", "body": "Aurora 구성\nDMS 설정"},
      {"title": "테스트", "body": "Full Load\nCDC 검증"},
      {"title": "컷오버", "body": "서비스 전환\n정합성 확인"}
    ]
  }
}
```

### comparison_table
```json
{
  "layout": "comparison_table",
  "title": "방안 비교표",
  "content": {
    "columns": [{"title": "항목"}, {"title": "현재 (PG14)"}, {"title": "목표 (PG17)"}],
    "rows": [
      ["버전", "14.9", "17.2"],
      ["HA", "Single", "Aurora Multi-AZ"],
      ["모니터링", "수동", "Datadog APM"]
    ]
  }
}
```

### challenge_solution (예외: wrapper 레벨)
```json
{
  "layout": "challenge_solution",
  "title": "문제 & 해결",
  "content": {
    "challenge": {"title": "CHALLENGE", "body": "PG14→17 호환성 이슈\nExtension 미지원 가능성"},
    "solution": {"title": "SOLUTION", "body": "사전 호환성 검증 수행\n대체 Extension 확보"}
  }
}
```
