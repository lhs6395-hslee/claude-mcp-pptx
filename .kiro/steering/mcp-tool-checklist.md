---
inclusion: auto
---

# MCP 도구 호출 및 검증 체크리스트

## create_presentation 도구 호출 전 필수 확인

**🚨 이 체크리스트를 거치지 않으면 100% 에러 발생!**

**⚠️ 가장 중요한 규칙:**
- **sections 파라미터는 데이터가 아무리 크더라도 반드시 작성해야 함**
- **작성을 시작했으면 반드시 완료해야 함**
- **중간에 멈추거나 생략하면 절대 안 됨**
- **생성 후 반드시 파일을 열어서 내용 검증**

### 1단계: 파라미터 준비 확인
```
[ ] cover_title 작성 완료
[ ] cover_subtitle 작성 완료
[ ] sections 배열 작성 완료 ← 가장 중요!
[ ] output_name 작성 (선택)
```

### 2단계: sections 배열 검증
```
[ ] sections가 빈 배열이 아님 (최소 1개 섹션)
[ ] 각 섹션에 section_title 있음
[ ] 각 섹션에 slides 배열 있음 (최소 1개 슬라이드)
[ ] 각 슬라이드에 layout, title, content 있음
[ ] content에 충분한 텍스트 포함 (짧은 텍스트 금지)
```

### 3단계: 도구 호출
```python
# ✅ 올바른 호출
mcp_pptx_generator_create_presentation(
    cover_title="제목",
    cover_subtitle="부제",
    sections=[...],  # 반드시 포함!
    output_name="파일명"
)

# ❌ 잘못된 호출 - sections 누락
mcp_pptx_generator_create_presentation(
    cover_title="제목",
    cover_subtitle="부제",
    output_name="파일명"
)
# → Error: Field required [type=missing, input_value=..., input_type=dict]
```

### 4단계: 생성 후 검증 (필수!)

**파일 생성 후 반드시 수행:**

1. **슬라이드 개수 확인**
   ```python
   from pptx import Presentation
   prs = Presentation("results/파일명.pptx")
   print(f'총 슬라이드: {len(prs.slides)}개')
   ```

2. **각 슬라이드 제목 확인**
   ```python
   for i, slide in enumerate(prs.slides, 1):
       title = 'No title'
       for shape in slide.shapes:
           if shape.has_text_frame and shape.text.strip():
               title = shape.text.strip()[:50]
               break
       print(f'{i}. {title}')
   ```

3. **내용 검증 - 빈 슬라이드 찾기**
   ```python
   empty_slides = []
   for i, slide in enumerate(prs.slides, 1):
       text_count = 0
       total_chars = 0
       for shape in slide.shapes:
           if hasattr(shape, "text_frame") and shape.text.strip():
               text_count += 1
               total_chars += len(shape.text.strip())
       
       if i > 2 and i < len(prs.slides):  # 표지/목차/엔딩 제외
           if text_count < 3 or total_chars < 100:
               empty_slides.append(f"Slide {i}: 내용 부족")
   
   if empty_slides:
       print("⚠️ 내용 부족한 슬라이드:")
       for slide in empty_slides:
           print(f"  - {slide}")
   ```

4. **텍스트 오버플로우 검증**
   ```python
   issues = []
   for slide_num, slide in enumerate(prs.slides, 1):
       for shape in slide.shapes:
           if hasattr(shape, "text_frame"):
               text = shape.text_frame.text.strip()
               if len(text) > 800:
                   issues.append(f"Slide {slide_num}: 텍스트 너무 김")
               lines = text.split('\n')
               if len(lines) > 15:
                   issues.append(f"Slide {slide_num}: 줄 수 많음")
   
   if issues:
       print("⚠️ 텍스트 오버플로우:")
       for issue in issues:
           print(f"  - {issue}")
   ```

## 자주 발생하는 에러와 해결

### 에러 1: sections field required
```
Error: 1 validation error for create_presentationArguments
sections
  Field required
```

**원인:** sections 파라미터를 작성하지 않고 도구 호출

**해결:** sections 배열을 반드시 포함해서 호출

### 에러 2: 도구 호출 중간에 멈춤 또는 sections 파라미터가 너무 큼
**원인:** sections 데이터가 너무 커서 validation error 발생

**해결: 5-10개 슬라이드씩 나눠서 누적 생성**

**전략:**
1. **한 번에 5-10개 슬라이드만 생성** (sections 파라미터 크기 제한)
2. 같은 `output_name` 사용하여 하나의 파일에 누적
3. 각 호출마다 sections 파라미터를 완전히 작성
4. 각 호출 후 검증 수행

**예시: 41개 레이아웃 쇼케이스 생성**
- 1차 호출: 섹션 1 (5개) → 새 파일 생성
- 2차 호출: 섹션 2 (7개) → 같은 파일에 누적
- 3차 호출: 섹션 3 (7개) → 같은 파일에 누적
- 4차 호출: 섹션 4 (7개) → 같은 파일에 누적
- 5차 호출: 섹션 5 (11개) → 같은 파일에 누적
- 6차 호출: 섹션 6 전반부 (5개) → 같은 파일에 누적
- 7차 호출: 섹션 6 후반부 (6개) → 같은 파일에 누적
- 8차 호출: 섹션 6 마지막 (1개) → 같은 파일에 누적

**중요:**
- 섹션이 10개 이상의 슬라이드를 포함하면 2-3번으로 나눠서 호출
- 누적 모드는 자동으로 목차를 병합하고 엔딩 슬라이드를 마지막에 배치

### 에러 3: 내용이 비어있는 슬라이드
**원인:** content에 충분한 텍스트를 넣지 않음

**해결:**
- 각 카드/아이템에 최소 50자 이상의 설명 포함
- 짧은 단어만 넣지 말고 완전한 문장 사용
- 레이아웃별 권장 텍스트 길이 준수

### 에러 4: content 데이터 형식 오류
**원인:** 레이아웃별 content 형식이 잘못됨

**해결:** LAYOUTS.md에서 해당 레이아웃의 정확한 형식 확인

## 권장 작업 순서

1. **최소 버전으로 테스트**
   - 1개 섹션, 1개 슬라이드로 먼저 생성
   - 정상 작동 확인

2. **점진적 확장 (누적 기능 사용)**
   - 같은 파일명으로 5-10개 슬라이드씩 추가
   - 각 단계마다 검증 수행

3. **최종 검증**
   - 모든 슬라이드 내용 확인
   - 텍스트 오버플로우 검증
   - 빈 슬라이드 없는지 확인

## 최소 테스트 예제

```json
{
  "cover_title": "테스트",
  "cover_subtitle": "최소 버전",
  "sections": [
    {
      "section_title": "1. 테스트",
      "slides": [
        {
          "layout": "3_cards",
          "title": "1-1. 테스트",
          "content": {
            "card1_title": "클라우드 컴퓨팅",
            "card1_text": "클라우드 컴퓨팅은 인터넷을 통해 IT 리소스를 온디맨드로 제공하는 서비스입니다",
            "card2_title": "데이터베이스",
            "card2_text": "관리형 데이터베이스 서비스로 운영 부담을 줄이고 확장성을 확보합니다",
            "card3_title": "보안",
            "card3_text": "다층 보안 아키텍처로 데이터와 애플리케이션을 안전하게 보호합니다"
          }
        }
      ]
    }
  ],
  "output_name": "test"
}
```

이 최소 버전이 정상 작동하면, 점진적으로 슬라이드를 추가하세요.
