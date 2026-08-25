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
    부품  render_markdown · add_journal_table · brief_table · render_docx
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
from md_to_docx import add_runs, render_markdown, unwrap            # noqa: E402,F401
from manuscript_table import add_journal_table                      # noqa: E402,F401
# 골격 -- 부품 위의 한 층. 장르가 반복되면 여기로 올라온다.
from document_shell import (GateError, brief_shell,                 # noqa: E402,F401
                           fill_values, manuscript_shell)

__all__ = ["add_runs", "unwrap", "render_markdown", "add_journal_table",
           "render_docx", "brief_table",
           "manuscript_shell", "brief_shell", "fill_values", "GateError"]


def brief_table(*a, **k):
    """국문 브리프용 격자 표. 저널 원고에는 쓰지 않는다(장르가 다르다)."""
    from brief_builder import data_table
    return data_table(*a, **k)


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
    import win32com.client as win32
    import fitz

    src = Path(src).resolve()
    out = Path(out_dir) if out_dir else src.parent / "_render"
    out.mkdir(parents=True, exist_ok=True)

    tmp = Path(tempfile.gettempdir()) / ("_dk_" + src.name)
    shutil.copy2(src, tmp)
    pdf = tmp.with_suffix(".pdf")

    word = win32.DispatchEx("Word.Application")     # ⚠️ Ex -- 위 주석 참고
    word.Visible = False
    try:
        doc = word.Documents.Open(str(tmp), ReadOnly=True, AddToRecentFiles=False)
        doc.SaveAs(str(pdf), FileFormat=17)         # 17 = PDF
        doc.Close(False)
    finally:
        word.Quit()

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
    if len(sys.argv) > 1:
        for p in render_docx(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None):
            print(p)
    else:
        print(__doc__)
