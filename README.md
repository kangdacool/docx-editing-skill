<div align="center">

# 📄 DOCX Editing Skill

**LLM 에이전트가 Word(`.docx`) 문서를 학술 조판 관습대로 만들게 해주는 Agent Skill.**

`python-docx`로 문서를 조립할 때마다 반복되는 결함 — 줄마다 문단이 생겨 세로로 늘어지는 원고,
글자로 남은 `**굵게**`, 세로줄이 그어진 저널 표, 조용히 죽는 docx→PDF 렌더 — 을 검증된 규칙과
실제로 돌아가는 Python 도구로 묶었습니다. Claude Code · Codex · Cursor · Gemini CLI 등에서
그대로 씁니다.

_A portable Agent Skill that teaches AI coding agents to build Word documents that follow academic
typesetting conventions: journal tables without vertical rules, markdown that becomes real
paragraphs instead of one-line-per-paragraph, and a render-and-look QA step that catches what
structural checks cannot._

[![CI](https://github.com/kangdacool/docx-editing-skill/actions/workflows/ci.yml/badge.svg)](https://github.com/kangdacool/docx-editing-skill/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue)
![Agent Skill](https://img.shields.io/badge/format-SKILL.md-8A2BE2)
![Works with](https://img.shields.io/badge/agents-Claude%20Code%20%C2%B7%20Codex%20%C2%B7%20Cursor%20%C2%B7%20Gemini-orange)

**[한국어](#한국어)** · **[English](#english)**

</div>

---

## 한국어

### 왜 필요한가

`python-docx`로 문단을 넣고 표를 만드는 방법은 어디에나 있습니다. 없는 것은 **「그래서 학술
문서가 어떻게 생겨야 하는가」** 이고, 에이전트는 매번 거기서 같은 자리를 틀립니다.

- **마크다운을 그대로 넣으면 문서가 세로로 늘어진다.** 원본의 하드랩(`\n`)이 문단 경계로
  오해되어 3,800단어짜리 원고가 **413문단**이 됩니다. 빈 줄만 문단 경계입니다.
- **`**굵게**`가 글자로 남는다.** 별표를 지우고 run 을 나눠야 하는데, 대개 그 처리가 빠집니다.
- **저널 표에 세로줄이 그어진다.** `table.style = 'Table Grid'` 한 줄이 주범입니다. 출판된 저널
  표는 세로줄이 **하나도 없고** 가로줄도 셋(top / header / bottom)뿐입니다.
- **docx→PDF 렌더가 원인 없이 죽는다.** Word COM을 `Dispatch`로 잡으면 **이미 떠 있는** 인스턴스에
  붙습니다. 그게 숨은 채 멈춰 있으면 렌더가 조용히 실패하고, 원인이 «파일»에 있는 것처럼 보입니다.
  (실제로 「Word가 망가졌다」고 오진한 적이 있습니다. `DispatchEx`가 답입니다.)
- **장르를 틀리면 표 전체를 다시 짠다.** 저널 원고의 표와 기관 브리프의 표는 **다른 물건**입니다.
  원고에 격자·색 헤더를 넣으면 관습 위반이고, 브리프에 저널 표를 넣으면 정보가 덜 보입니다.

### 무엇이 들어 있나

```
skills/docx-editing/
├── SKILL.md                              에이전트가 읽는 규칙 (가리키기만 한다)
├── references/
│   └── docx-conventions.md               저널 표 규칙 · 렌더 함정 · 숫자 게이트
└── scripts/
    ├── docx_kit.py                       ★ 단 하나의 진입점 (재수출 + 렌더)
    ├── md_to_docx.py                     마크다운 → docx (문단 합치기·굵게·목록·표)
    ├── manuscript_table.py               저널 3선표 (세로줄 없음)
    ├── brief_builder.py                  브리프 표 (격자·색 헤더)
    └── docx_kit_selftest.py              절반이 «하지 않는 것»을 증명한다
```

### 설치

`~/.claude/skills/` (또는 쓰는 에이전트의 스킬 폴더)에 `skills/docx-editing/`를 복사합니다.

```bash
git clone https://github.com/kangdacool/docx-editing-skill
cp -r docx-editing-skill/skills/docx-editing ~/.claude/skills/
python ~/.claude/skills/docx-editing/scripts/docx_kit_selftest.py
```

`SKILL.md`를 읽지 않는 에이전트라면 `scripts/`만 가져다 라이브러리로 써도 됩니다.

### 쓰는 법

```python
import os, sys
sys.path.insert(0, os.path.expanduser(
    os.path.join("~", ".claude", "skills", "docx-editing", "scripts")))
from docx_kit import render_markdown, add_journal_table, brief_table, render_docx

from docx import Document
doc = Document()

render_markdown(doc, open("manuscript.md", encoding="utf-8").read())

add_journal_table(doc, rows, col_widths=[2.15, 1.15, 1.15, 0.7])   # 저널 원고
brief_table(doc, rows)                                             # 국문 브리프

doc.save("out/manuscript.docx")
render_docx("out/manuscript.docx", "scratch/_render")              # 쪽별 PNG — 눈으로 본다
```

**장르가 표를 정합니다.** 「보기 좋으니까」로 고르지 않습니다.

### 핵심 규율 넷

1. **빈 줄만 문단 경계다.** 하드랩은 이어 붙인다.
2. **저널 표에 세로줄을 긋지 않는다.** 가로줄은 셋뿐. 구분은 볼드 라벨 + 여백으로.
3. **숫자를 타이핑하지 않는다.** 값 사전 + 자리표시자 + **미치환이면 빌드 실패**.
   ⚠️ 행을 «고르는» 이름(모형 라벨 등)도 숫자다 — 구성이 바뀌면 값은 그대로인 채 다른 행을 읽는다.
4. **만들었으면 렌더해서 눈으로 본다.** 글씨 축소·쪽 넘김·표 잘림·별표 잔존은 구조 검사로
   안 잡히고 렌더에서만 보인다.

자세한 근거와 함정은 [`references/docx-conventions.md`](skills/docx-editing/references/docx-conventions.md).

### 자기검사

```bash
python skills/docx-editing/scripts/docx_kit_selftest.py
```

절반은 **하지 않는 것**을 증명합니다 — 줄마다 문단을 만들지 않는가, 별표를 글자로 남기지
않는가, 저널 표에 세로줄을 넣지 않는가, 렌더가 맨몸 `Dispatch`를 쓰지 않는가.

### 요구사항

`python-docx` · Python 3.10+.
`render_docx()`만 추가로 필요합니다 — Windows + Word(`pywin32`), 또는 LibreOffice,
그리고 PNG 변환에 `PyMuPDF`. **import 시점이 아니라 호출 시점에** 잡으므로 렌더를 안 쓰면
아무것도 설치할 필요가 없습니다.

---

## English

### Why

Recipes for `python-docx` are everywhere. What is missing is **what an academic document should
actually look like** — and agents get the same handful of things wrong every time.

- **Markdown pasted in stretches the document vertically.** Hard wraps (`\n`) get read as paragraph
  breaks: a 3,800-word manuscript becomes **413 paragraphs**. Only blank lines end a paragraph.
- **`**bold**` survives as literal asterisks** unless you strip the markers and split runs.
- **Journal tables come out gridded.** `table.style = 'Table Grid'` is the culprit. Published
  journal tables have **no vertical rules at all** and exactly three horizontal ones.
- **docx→PDF rendering dies for no visible reason.** `win32.Dispatch` attaches to an *existing*
  Word instance; if a hidden one is stuck, rendering fails silently and it looks like the file is
  broken. It isn't. Use `DispatchEx`.
- **Getting the genre wrong means rebuilding every table.** Journal manuscripts and institutional
  briefs need *different* tables — a grid in a manuscript violates convention, and a journal table
  in a brief hides information rather than looking austere.

### Install

```bash
git clone https://github.com/kangdacool/docx-editing-skill
cp -r docx-editing-skill/skills/docx-editing ~/.claude/skills/
python ~/.claude/skills/docx-editing/scripts/docx_kit_selftest.py
```

Agents that don't read `SKILL.md` can use `scripts/` as a plain library.

### Use

```python
from docx_kit import render_markdown, add_journal_table, brief_table, render_docx
```

One entry point. `render_markdown` for prose, `add_journal_table` for manuscripts (no vertical
rules, caption above, notes below), `brief_table` for reports where a grid is the intended design,
`render_docx` to look at the result page by page.

### The four rules

1. **Only blank lines break paragraphs.** Join hard wraps.
2. **No vertical rules in journal tables.** Three horizontal rules; separate groups with bold
   labels and whitespace.
3. **Never type a number.** Value dictionary + placeholders + a gate that **fails the build** on any
   unsubstituted placeholder. Row *selectors* (model labels and the like) are numbers too.
4. **Render it and look.** Shrunken text, bad page breaks, clipped tables and leftover asterisks are
   invisible to structural checks.

See [`references/docx-conventions.md`](skills/docx-editing/references/docx-conventions.md).

### Requirements

`python-docx`, Python 3.10+. `render_docx()` additionally needs Word via `pywin32` (Windows) or
LibreOffice, plus `PyMuPDF` for PNG conversion — imported lazily, so you need none of it unless you
render.

---

## License

MIT — see [LICENSE](LICENSE).
