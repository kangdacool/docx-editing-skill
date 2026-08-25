---
name: docx-editing
description: >-
  Build, edit, and QA Word (.docx) documents with python-docx without re-inventing
  tools that already exist or breaking typesetting. Use this whenever a task involves
  a .docx file, Word / 워드 / 원고 / manuscript / 보고서 / 브리프 / 구성안 / 초록, or
  writing a builder that emits .docx. Ships its own toolbox (scripts/docx_kit.py),
  the genre→table-module rule, the render-and-look step, and points at the audit gate.
---

# docx — 만들거나 고치기 전에

형제: [[hwpx-editing]] · [[pptx-editing]]. 같은 자리에 있는 세 번째 형식이다.

**감사에는 단일 진입점(`audit.py`)이 있어서 작동한다. 제작에는 없어서 실패한다.**
2026-08-25 하루에 세 번, 「맞는 답이 이미 있는데 못 찾은」 실패가 났다 — 마크다운→docx를
다시 짜고, 렌더가 죽은 원인을 오진하고, 표를 다른 장르 모듈로 짰다.

이 파일은 **가리키기만 한다.** 내용은 도구 안에 있다.

## 1. 코드는 이 스킬 안에 있다

```python
import os, sys
sys.path.insert(0, os.path.expanduser(
    os.path.join("~", ".claude", "skills", "docx-editing", "scripts")))
from docx_kit import render_markdown, add_journal_table, brief_table, render_docx
```

⚠️ **기계 고정 경로(`D:\...`)를 쓰지 않는다.** 형제 스킬의 소비자도 `expanduser`를 쓴다.

`scripts/` 안에 전부 있다 — `docx_kit.py`(진입점) · `document_shell.py`(골격) ·
`md_to_docx.py` · `manuscript_table.py` · `brief_builder.py`. `docx_kit.py` 머리말에
쓰는 법과 함정이 있으니 **직접 짜기 전에 그 파일을 연다.**

> *(내력)* 2026-08-25까지 이 코드는 스킬 밖의 공용 도구 폴더에 있었다. 형제(hwpx·pptx)와 달랐고, 그래서
> 못 찾아 다시 짰다. 옛 경로에는 **얇은 shim만** 남아 있다(기존 import 보호용) —
> 새 코드는 위 경로로 직접 가져온다.

## 2. 장르가 골격과 표를 정한다 — 틀리면 문서 전체를 다시 짠다

**먼저 이 표에서 자기 장르를 찾는다.** 골격이 있으면 골격부터 부른다.

| 장르 | 골격 | 표 |
|---|---|---|
| 저널 투고 원고 (totale) | **`manuscript_shell`** | `add_journal_table` |
| 연구요약·브리프·사례보고서 | **`brief_shell`** | `brief_table` |
| 전시물 구성안 (표·그림만) | — (부품) | `add_journal_table` |
| 정부·감독자 보고서 | — (부품) | `brief_table` |
| 수업자료·문항·케이스 | — (부품) | 장르별 |
| 공부문서·가이드 (자유서식) | `render_markdown` 자체가 골격 | — |

⚠️ 「보기 좋으니까」로 고르지 않는다. **원고에 격자 표를 넣으면 저널 관습 위반이고,
브리프에 저널 표를 넣으면 밋밋해 보이는 게 아니라 정보가 덜 보인다.**

### 골격이 왜 따로 있나 — 부품만으로는 안 멈췄다

실측(2026-08-25): docx를 만드는 스크립트 **64개 중 kit을 쓴 것은 11개**였고, 직접 짠 53개 중
**51개가 「굵게 처리」를 다시 짰다** — kit에 있는데도. 부품이 모자란 게 아니라 **골격을 아무도
안 줘서** 매번 표지·쪽나눔·표번호·캡션·게이트를 처음부터 엮다가 부품까지 다시 짠 것이다.
같은 골격의 원고 totale 빌더가 **7개**(254~741줄)로 흩어져 있었다.

```python
res = manuscript_shell(
    out, TITLE, "manuscript/sections/*.md",
    cover=review_cover,          # 검토 표지 — 쪽을 끊어 붙으므로 투고 시 통째로 뺀다
    abstract=abstract_md,
    tables=[(caption, rows, widths)], figures=[(caption, path)],
    values=VALUES,               # 미치환이면 GateError로 «멈춘다»
    word_limit=5000, no_count=("06_declarations",),
)
print(res["body_words"], res["refs"])   # 보고는 호출자가 한다
```

**골격이 없는 장르는 «측정하지 않아서» 없다.** 같은 골격이 셋 이상 반복되는 것이 보이면
그때 `document_shell.py`로 올린다 — 그 파일이 생긴 근거가 그것이다.

⚠️ **골격은 문장을 만들지 않는다.** 산문은 사람이 쓰고 골격은 배치·번호·게이트만 맡는다.

## 3. 숫자는 타이핑하지 않는다

산출물 수치를 본문에 손으로 적지 않는다. 사전 + 자리표시자 + **미치환이면 멈추는 게이트**.
**행을 «고르는» 이름(모형 라벨 등)도 숫자다** — 구성이 바뀌면 값은 그대로인 채 다른 행을
읽는다. 구현 패턴은 `references/docx-conventions.md` §6.

## 4. 만들었으면 렌더해서 «눈으로» 본다

```python
render_docx("out/문서.docx", "scratch/_render")     # PNG 쪽별
```

조판 결함(글씨 축소·쪽 넘김·칸 밖으로 나간 글자·별표가 글자로 남음)은 **구조 검사로
안 잡히고 렌더에서만 보인다.** 2026-08-25에 이 단계가 6종을 잡았다.

## 5. 끝나면 감사

조판 감사는 이 스킬 «밖»에 있다 — 하나의 진입점이 장르(docx·pptx·hwpx·md)로 라우팅하고,
건너뛴 검사와 그 이유까지 보고한다. 이 환경에서는:

```bash
python agent/tools/audit.py <파일>       # 없으면 이 절은 건너뛴다
```

글자 크기·칸 폭·고아 표시물(인용되지 않은 표·그림)·본문↔표 수치 일치를 본다.
`PostToolUse` 훅이 자동으로도 돈다. **감사는 렌더 확인(§4)을 대체하지 않는다** —
서로 다른 것을 잡는다.

---

## 더 읽을 것

- **`references/docx-conventions.md`** — 저널 표 규칙, 마크다운→docx 결함 넷, 그림 크기,
  렌더 함정(`Dispatch`/`DispatchEx`), 숫자 게이트. *직접 조립해야 할 때 여는 문서.*
- `scripts/docx_kit.py` 머리말 — 각 함수의 쓰는 법과 그것이 막는 실패.

## 자기검사

```bash
python ~/.claude/skills/docx-editing/scripts/docx_kit_selftest.py
```

절반은 «하지 않는 것»을 증명한다 — 줄마다 문단을 만들지 않는가, 별표를 글자로 남기지
않는가, 저널 표에 세로줄을 넣지 않는가, 렌더가 `Dispatch`(맨몸)를 쓰지 않는가.
