#!/usr/bin/env python
"""md_to_docx.py -- render a plain markdown document to a readable .docx.

Built for the data-request paperwork, which has to be read and edited by a person
and pasted into web forms, not typeset for a journal. Deliberately small: headings,
paragraphs, bold/italic runs, bullet and numbered lists, block quotes and pipe
tables. No figures, no styles beyond the built-ins.

    python tools/md_to_docx.py <in.md> [out.docx]
"""
import os
import re
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

INLINE = re.compile(r"(\*\*.+?\*\*|`.+?`|\*[^*]+?\*)")


def add_runs(par, text):
    """Render **bold**, `code` and *italic* inside one paragraph."""
    for piece in INLINE.split(text):
        if not piece:
            continue
        if piece.startswith("**") and piece.endswith("**"):
            par.add_run(piece[2:-2]).bold = True
        elif piece.startswith("`") and piece.endswith("`"):
            r = par.add_run(piece[1:-1])
            r.font.name = "Consolas"
            r.font.size = Pt(9.5)
        elif piece.startswith("*") and piece.endswith("*") and len(piece) > 2:
            par.add_run(piece[1:-1]).italic = True
        else:
            par.add_run(piece)


def unwrap(lines):
    """Join hard-wrapped continuation lines back onto their list item or quote.

    A bold span that crosses a line break inside a list item -- '**MCPS\\n
    publication**' -- otherwise ends up split across two paragraphs and the
    asterisks are rendered literally. Found exactly that way in the MCPS
    registration draft, so the joining happens before parsing rather than being
    patched per-case afterwards.

    두 종류를 «둘 다» 이어야 한다 -- 인용문 «아래의 들여쓴» 줄과, 인용문 «자신의»
    '>' 로 시작하는 연속줄. 오래도록 앞엣것만 처리하면서 docstring 은 "list item or
    quote" 라고 말하고 있었다(2026-09-07 실측: 보고서 서론에서 '**인구당 병원 수가
    자원을 / 연속량으로 가정**' 이 별표째 찍히고, 인용문 하나가 문단 다섯 개로
    쪼개졌다). '>' 뒤가 빈 줄이면 인용문 «안의» 문단 구분이므로 잇지 않는다.
    """
    out = []
    for raw in lines:
        line = raw.rstrip()
        prev = out[-1] if out else ""
        # two spaces is a valid list continuation in markdown, not three -- requiring
        # three silently split '**bold\n  text**' across paragraphs and leaked the
        # asterisks into the rendered document.
        is_cont = (line.startswith(("  ", "\t")) and line.strip()
                   and not line.strip().startswith(("|", ">", "-", "*", "#", "```"))
                   and not re.match(r"^\s*\d+\.\s", line))
        parent_is_item = bool(re.match(r"^(\s*)([-*]|\d+\.)\s+", prev)) or \
            prev.lstrip().startswith(">")
        # 인용문 자신의 연속줄. 위 is_cont 는 '>' 를 제외하므로 여기서 따로 잡는다 --
        # 안 이으면 '> **인구당 병원 수가\n> 연속량으로 가정**' 의 짝이 갈라져 별표가
        # 그대로 찍히고, 인용문 하나가 문단 여럿으로 쪼개진다(2026-09-07 실측).
        # '>' 뒤가 빈 줄이면 인용문 «안의» 문단 구분이므로 잇지 않는다.
        quote_cont = (line.lstrip().startswith(">")
                      and prev.lstrip().startswith(">")
                      and line.lstrip().lstrip(">").strip()
                      and prev.lstrip().lstrip(">").strip())
        if is_cont and parent_is_item:
            out[-1] = prev + " " + line.strip()
        elif quote_cont:
            out[-1] = prev + " " + line.lstrip().lstrip(">").strip()
        else:
            out.append(line)
    return out


def flush_table(doc, rows):
    rows = [r for r in rows if not re.fullmatch(r"\s*\|[\s:|-]+\|\s*", r)]
    cells = [[c.strip() for c in r.strip().strip("|").split("|")] for r in rows]
    if not cells:
        return
    t = doc.add_table(rows=len(cells), cols=max(len(c) for c in cells))
    t.style = "Light Grid Accent 1"
    for i, row in enumerate(cells):
        for j, val in enumerate(row):
            par = t.cell(i, j).paragraphs[0]
            add_runs(par, val)
            if i == 0:
                for r in par.runs:
                    r.bold = True
    doc.add_paragraph()


def manuscript_typography(doc, font="Times New Roman", size=11):
    """문서 전체를 «원고 조판»으로 맞춘다 — 흑백, 한 서체, 군더더기 여백 없음.

    왜 필요한가 (2026-09-03): `render_markdown` 은 `doc.add_heading()` 을 쓰고, 그것은
    Word 기본 Heading 스타일이라 **제목이 파랗게** 나온다. 학회 초록·원고처럼 «인쇄해서
    읽는» 문서에 그 파란색과 Calibri 는 맞지 않는다. 사용자 지적:
    *"원고처럼, 글자는 흑백, 글씨체도 원고처럼 두고, 쓸데없는 구간 넣지 마."*

    **`render_markdown` 을 부르기 «전에»** 호출한다 — 스타일을 먼저 갈아둬야 그 위에
    쌓인다. 매 빌더에서 다시 짜지 말 것(그 재작성이 이 킷이 생긴 이유다).
    """
    from docx.enum.text import WD_LINE_SPACING

    def _set(st, pt, bold=None, space_before=0, space_after=6):
        st.font.name = font
        st.font.size = Pt(pt)
        st.font.color.rgb = RGBColor(0, 0, 0)          # 흑백 — Heading 의 파란색을 지운다
        if bold is not None:
            st.font.bold = bold
        pf = st.paragraph_format
        pf.space_before = Pt(space_before)
        pf.space_after = Pt(space_after)
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        # 동아시아 글꼴도 같이 지정하지 않으면 한글만 다른 서체로 나온다.
        rpr = st.element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.append(rf)
        # ⚠ Word 의 Heading 스타일은 글꼴을 «테마»로 참조한다(`w:asciiTheme="majorHAnsi"`).
        #   테마 속성이 남아 있으면 `w:ascii` 를 새로 써도 **테마가 이긴다** — 본문만 서체가
        #   바뀌고 제목은 산세리프 그대로 나온다(2026-09-03 렌더에서 확인). 먼저 지운다.
        for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            if rf.get(qn(attr)) is not None:
                del rf.attrib[qn(attr)]
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            rf.set(qn(attr), font)

    _set(doc.styles["Normal"], size, space_after=6)
    # 제목 단계는 «크기»가 아니라 굵기로 가른다 — 초록·원고는 위계가 얕다.
    for lvl, pt in ((1, size + 2), (2, size), (3, size), (4, size)):
        try:
            _set(doc.styles["Heading %d" % lvl], pt, bold=True, space_before=8, space_after=3)
        except KeyError:
            pass
    return doc


def render_markdown(doc, text):
    """마크다운 «본문»을 주어진 doc에 이어 붙인다. 문단·굵게·목록·인용·표·코드.

    ⭐⭐ 2026-08-25에 main()에서 뺐다. 이 루프가 main() 안에 갇혀 있어서 다른
       빌더가 재사용하지 못하고 «더 못한 판»을 다시 짰다(굵게만 처리하고 하드랩
       문단 합치기를 새로 구현). 동작은 그대로다.

    ⚠️ 핵심은 buf/flush_para다 -- 88자에서 손으로 줄바꿈한 산문은 «빈 줄»까지
       모아 한 문단으로 합쳐야 한다. 줄마다 add_paragraph 하면 한 문장이 여러
       문단으로 쪼개지고, 괄호 짝 검사 같은 것이 전부 오탐이 된다.
    """
    lines = unwrap(text.splitlines())
    buf, table = [], []

    def flush_para():
        if buf:
            add_runs(doc.add_paragraph(), " ".join(buf))
            buf.clear()

    in_code = False
    for raw in lines:
        line = raw.rstrip()
        # fenced code block: keep the layout, drop the fences, no inline markup.
        # Without this the ``` lines and any * or ` inside them are rendered literally,
        # which is how a checkbox flow diagram ends up full of stray asterisks.
        if line.strip().startswith("```"):
            flush_para()
            in_code = not in_code
            continue
        if in_code:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.left_indent = Pt(14)
            r = p.add_run(raw if raw.strip() else " ")
            r.font.name = "Consolas"
            r.font.size = Pt(9)
            continue
        if line.strip().startswith("|"):
            flush_para()
            table.append(line)
            continue
        if table:
            flush_table(doc, table)
            table = []
        if not line.strip():
            flush_para()
            continue
        if line.startswith("#"):
            flush_para()
            lvl = len(line) - len(line.lstrip("#"))
            doc.add_heading(line.lstrip("#").strip(), level=min(lvl, 4))
            continue
        if line.strip() in ("---", "***"):
            flush_para()
            continue
        if line.lstrip().startswith(">"):
            flush_para()
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(22)
            add_runs(p, line.lstrip().lstrip(">").strip())
            for r in p.runs:
                r.italic = True
                r.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
            continue
        m = re.match(r"^(\s*)([-*]|\d+\.)\s+(.*)$", line)
        # 줄바꿈된 산문이 연도로 시작하면 목록이 아니다. "…in July\n2016. The dental…"
        # 에서 둘째 줄이 목록으로 잡혀 문단이 끊기고 연도가 산출물에서 사라졌다
        # (2026-09-07 실측, 통독에서만 보였다). 문단을 모으는 중이고 번호가 100 이상이면
        # 산문으로 본다 -- 진짜 목록은 앞에 빈 줄/제목이 와서 buf 가 비어 있다.
        if m and buf and m.group(2).endswith(".") and m.group(2)[:-1].isdigit() \
                and int(m.group(2)[:-1]) >= 100:
            m = None
        if m:
            flush_para()
            bullet = "List Bullet" if m.group(2) in "-*" else "List Number"
            p = doc.add_paragraph(style=bullet)
            add_runs(p, m.group(3))
            continue
        buf.append(line.strip())

    flush_para()
    if table:
        flush_table(doc, table)

    return doc


def main():
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(src)[0] + ".docx"
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10.5)
    style.paragraph_format.space_after = Pt(7)

    render_markdown(doc, open(src, encoding="utf-8").read())

    doc.save(out)
    print("wrote %s" % out)
    print("  paragraphs: %d   tables: %d" % (len(doc.paragraphs), len(doc.tables)))


if __name__ == "__main__":
    main()
