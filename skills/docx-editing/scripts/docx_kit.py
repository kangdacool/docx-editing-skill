#!/usr/bin/env python3
"""docx_kit -- .docx 를 «만들 때» 여는 단 하나의 문.

    import os, sys
    sys.path.insert(0, os.path.expanduser(
        os.path.join("~", ".claude", "skills", "docx-editing", "scripts")))
    from docx_kit import render_markdown, add_journal_table, brief_table, render_docx

왜 이 파일이 있나 (2026-08-25)
──────────────────────────────
감사에는 단일 진입점(`audit.py`)이 있어서 작동했다. **제작에는 없어서** 하루에 세 번
「맞는 답이 이미 있는데 못 찾은」 실패가 났다:
  · 마크다운→docx 를 다시 짰다 -- 굵게만 처리하고 하드랩 문단 합치기를 새로 구현.
    `md_to_docx.py`가 이미 다 하고 있었다(그 루프가 main() 안에 갇혀 있었을 뿐이라
    같은 날 `render_markdown()`으로 뺐다).
  · docx→PDF 렌더가 죽었다. 원인은 Word가 아니라 `win32.Dispatch`였다(아래).
  · 표를 브리프 모듈로 짰다 -- 장르가 다르면 표가 다르다(아래).

**교훈은 그 일을 하는 «단 하나의 도구» 안에 산다.** 길이 하나면 틀릴 길이 없다.
지식을 다른 곳에 또 적지 않는다 -- 여기 있는 것이 전부이고, 나머지는 가리킨다.

부품이냐 골격이냐
────────────────
    부품  render_markdown · add_journal_table · brief_table · render_docx · page_count
    골격  manuscript_shell(원고 totale) · brief_shell(국문 브리프)
          -- 표지·쪽나눔·표번호·그림캡션·게이트까지 «배치»를 맡는다.
2026-08-25 실측: 부품만 있을 때 64개 빌더 중 11개만 kit을 썼고, 직접 짠 53개 중
51개가 «굵게 처리»를 다시 짰다. 골격이 없어 매번 처음부터 엮었기 때문이다.

장르 → 표 모듈 (틀리면 표 전체를 다시 짠다)
──────────────────────────────────────────
    저널 투고 원고·구성안   `add_journal_table`   세로줄 없음 · 색 없음 ·
                                                 제목은 표 «위» · 각주는 «아래»
    국문 브리프·사례보고서  `brief_builder.data_table`  격자·색 헤더가 «의도된» 디자인
    자유서식 문서           `render_markdown`     .md 를 그대로 읽히게

표 열 폭 — widths 를 안 주면 이제 두 표 모듈 다 «내용 비례»가 기본값 (2026-09-09)
──────────────────────────────────────────
    균등 분할 금지("표 width 는 내가 맨날 하는 말" — 사용자가 표 12개를 손으로
    재배분한 날 코드가 됐다). 명시적 widths 인자가 언제나 이긴다. 계산기는 장르별 둘:
      저널 표(영문)   `manuscript_table.col_widths_for`   문자수·감쇠·water-filling
      국문·격자 표    `col_widths.content_col_widths`     CJK 표시폭·snug/지배·균등 유지
    새 국문 빌더를 직접 짜면 `from col_widths import content_col_widths` 를 쓴다 —
    ⚠ 두 함수를 합치거나 서로 바꿔 쓰지 말 것(col_widths.py 머리말의 판정 참조).

만든 뒤에는 반드시 «렌더해서 눈으로» 본다 -- render_docx(). 구조 검사로 안 잡히고
렌더에서만 보이는 결함이 있다. 조판 규율의 근거는 references/docx-conventions.md.
"""
import shutil
import sys
import tempfile
from pathlib import Path

_HERE = Path(__file__).resolve().parent   # 형제 모듈이 같은 폴더에 있다(2026-08-25 이동)
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

# ── 재수출: 이미 있는 것을 다시 짜지 않는다 ──────────────────────────────────
from md_to_docx import (add_runs, manuscript_typography, render_markdown,  # noqa: E402,F401
                        unwrap)
from manuscript_table import add_journal_table                      # noqa: E402,F401
from col_widths import content_col_widths, doc_text_width_cm        # noqa: E402,F401
# 골격 -- 부품 위의 한 층. 장르가 반복되면 여기로 올라온다.
from document_shell import (GateError, brief_shell,                 # noqa: E402,F401
                           fill_values, manuscript_shell)

__all__ = ["add_runs", "unwrap", "render_markdown", "manuscript_typography",
           "add_journal_table", "content_col_widths", "doc_text_width_cm",
           "render_docx", "page_count", "brief_table",
           "manuscript_shell", "brief_shell", "fill_values", "GateError", "finalize_docx"]


def finalize_docx(path, author="Kang Seo"):
    """⛔ 사람에게 넘기는 모든 .docx 의 «마지막 단계» — Word(새 프로세스)로 열어 그대로 다시 저장한다.

    연구자 지시(2026-10-02): 「앞으로도 word 재저장을 꼭 해라. 모든 연구 파이프라인 산출물 규칙으로」.
    python-docx 의 기본 템플릿은 저장 파일에 남의 흔적을 남긴다 — 실측:
      app.xml  Application=「Microsoft Macintosh Word」 · AppVersion 14 · 단어/문단 수 0
      docProps/thumbnail.jpeg  템플릿에 딸려 온 남의 썸네일
    core.xml 의 creator 를 고쳐도 이것들은 남는다(audit_doc_properties 는 core 만 본다).
    Word 로 재저장하면 이 PC 의 Word 값(16.0 · 실제 단어 수)이 되고 썸네일이 빠진다.
    실제 일은 `agent/tools/office_finalize.py` 가 한다(.pptx·.xlsx 도 같은 문 — 길은 하나).
    내용이 재저장 전후로 다르면 원본을 그대로 두고 멈춘다.
    """
    sys.path.insert(0, r"D:\onedrive\claude\agent\tools")
    from office_finalize import finalize
    return finalize(str(Path(path).resolve()), author)


def brief_table(*a, **k):
    """국문 브리프용 격자 표. 저널 원고에는 쓰지 않는다(장르가 다르다)."""
    from brief_builder import data_table
    return data_table(*a, **k)


def _to_pdf(src):
    """docx → PDF. (임시 docx, PDF) 경로를 돌려준다 — 둘 다 부른 쪽이 지운다.

    render_docx 와 page_count 가 같은 경로를 쓰게 하려고 뽑았다. `DispatchEx` 과
    «사본을 연다» 는 두 규칙이 한 곳에만 있어야 한 쪽이 빠지지 않는다(render_docx 주석 참고).
    """
    import win32com.client as win32

    src = Path(src).resolve()
    tmp = Path(tempfile.gettempdir()) / ("_dk_" + src.name)
    shutil.copy2(src, tmp)
    pdf = tmp.with_suffix(".pdf")

    word = win32.DispatchEx("Word.Application")     # ⚠️ Ex -- render_docx 주석 참고
    word.Visible = False
    try:
        doc = word.Documents.Open(str(tmp), ReadOnly=True, AddToRecentFiles=False)
        doc.SaveAs(str(pdf), FileFormat=17)         # 17 = PDF
        doc.Close(False)
    finally:
        word.Quit()
    return tmp, pdf


def page_count(*srcs):
    """쪽수만 센다 — PNG 를 만들지 않는다. [(이름, 쪽, 표, 표칸수), ...] 반환.

    «분량이 긴가» 는 혼자서는 답이 안 나온다. 여러 파일을 한 번에 받게 한 것은 그 때문이다 —
    같은 과목의 다른 산출물·이전 판을 같은 줄에 놓고 본다.
    ⚠ 그러려면 비교 대상이 «같은 규칙으로 만들어졌는가» 를 먼저 확인해야 한다. 표 칸수를
    함께 내는 이유가 그것이다 — 표가 2개인 문서와 20개인 문서는 쪽수를 비교할 사이가 아니다.
    (2026-08-28 실측: 5쪽짜리 «동료 문서» 셋을 21쪽 문서 옆에 놓았는데, 그것들은 표가 2개인
     다른 골격의 문서였다. «우리 것이 너무 길다» 는 거짓 결론을 그 표가 만들었다.)
    """
    import fitz
    from docx import Document

    out = []
    for src in srcs:
        src = Path(src)
        if not src.exists():
            out.append((src.name, None, None, None))
            continue
        tmp, pdf = _to_pdf(src)
        with fitz.open(pdf) as d:
            n = d.page_count
        dx = Document(str(src))
        tbl = len(dx.tables)
        cells = sum(len(t.rows) * len(t.columns) for t in dx.tables)
        for f in (tmp, pdf):
            try:
                f.unlink()
            except OSError:
                pass
        out.append((src.name, n, tbl, cells))
    return out


def render_docx(src, out_dir=None, dpi=110):
    """docx → PDF → 쪽별 PNG. 만든 PNG 경로 목록을 돌려준다.

    조판 결함(글씨 축소·쪽 넘김·상자 밖으로 나간 글자)은 구조 검사로 안 잡히고
    **렌더에서만 보인다.** 그래서 사람에게 넘기기 전에 이걸 돌린다.

    ⚠️⚠️ **`DispatchEx` 를 쓴다. `Dispatch` 를 쓰지 마라.**
       `Dispatch`는 «이미 떠 있는» Word 인스턴스에 붙는다. 자동화가 남긴 숨은
       인스턴스(창 없음, Visible=False)가 한 번 막히면 그 뒤 모든 렌더가 그것을
       물려받아 실패한다 -- 그리고 증상이 「Word가 망가졌다」로 보인다.
       2026-08-25에 그렇게 오진해 사용자 Word를 죽일 뻔했다. 실제로는 새 프로세스를
       띄우는 `DispatchEx` 한 글자 차이였다.
       (가르는 시험: «예전에 성공했던 파일»을 지금 열어 본다. 그것도 실패하면
        파일이 아니라 인스턴스 문제다.)
    ⚠️ 원본을 직접 열지 않고 임시 사본을 연다 -- 동기화·다른 Word가 잡고 있으면
       원본 열기가 실패하고, 원본에 잠금/최근문서 흔적을 남기지 않는 편이 낫다.
    """
    import fitz

    src = Path(src).resolve()
    out = Path(out_dir) if out_dir else src.parent / "_render"
    out.mkdir(parents=True, exist_ok=True)
    tmp, pdf = _to_pdf(src)

    made = []
    with fitz.open(pdf) as d:
        for i, page in enumerate(d):
            p = out / f"{src.stem}_p{i + 1}.png"
            page.get_pixmap(dpi=dpi).save(p)
            made.append(p)
    for f in (tmp, pdf):
        try:
            f.unlink()
        except OSError:
            pass
    return made


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--pages":
        print("%-46s %5s %5s %8s" % ("", "쪽", "표", "표칸수"))
        for name, n, t, c in page_count(*sys.argv[2:]):
            print("%-46s %5s %5s %8s" % (name[:46], n if n is not None else "없음",
                                         t if t is not None else "-", c if c is not None else "-"))
    elif len(sys.argv) > 1:
        for p in render_docx(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None):
            print(p)
    else:
        print(__doc__)
