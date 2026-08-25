#!/usr/bin/env python3
"""docx_kit 자체시험. 절반은 «하지 않는 것»을 증명한다(연구실 규율).

    python ~/.claude/skills/docx-editing/scripts/docx_kit_selftest.py

렌더(Word COM)는 여기서 돌리지 않는다 -- Word가 없는 기기에서도 나머지를 검사할 수
있어야 하고, 렌더는 `docx_kit.py <파일>`로 사람이 확인한다. 대신 «DispatchEx를 쓰는가»를
소스에서 검사한다: 그 한 글자가 2026-08-25 사고의 전부였다.
"""
import re
import sys
from pathlib import Path

# 콘솔이 cp949면 «»·한글이 인쇄되다 죽는다 -- 검사 «내용»과 무관한 실패다.
for _s in (sys.stdout, sys.stderr):
    if hasattr(_s, "reconfigure"):
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from docx import Document                                    # noqa: E402
from docx_kit import add_journal_table, render_markdown      # noqa: E402

fails = []


def check(name, ok, detail=""):
    print(f"  [{'OK  ' if ok else '실패'}] {name}" + (f"  -- {detail}" if not ok else ""))
    if not ok:
        fails.append(name)


print("=== render_markdown ===")
doc = Document()
render_markdown(doc, "# 제목\n\n손으로 줄바꿈한\n산문이 **굵게**를\n걸쳐 있다.\n\n둘째 문단.\n")
texts = [p.text for p in doc.paragraphs if p.text.strip()]

# ── 해야 하는 것 ──────────────────────────────────────────────────────────
check("하드랩 산문을 한 문단으로 합친다",
      any("손으로 줄바꿈한 산문이 굵게를 걸쳐 있다." == t for t in texts),
      f"실제: {texts}")
check("문단을 빈 줄로 가른다", "둘째 문단." in texts, f"실제: {texts}")
check("제목을 heading으로", any(p.style.name.startswith("Heading")
                             for p in doc.paragraphs if p.text == "제목"))
_runs = [r for p in doc.paragraphs for r in p.runs]
check("**굵게**를 굵은 run으로", any(r.bold and "굵게" in r.text for r in _runs))

# ── «하지 않아야» 하는 것 ─────────────────────────────────────────────────
check("별표를 글자로 남기지 «않는다»", not any("**" in t for t in texts),
      f"실제: {texts}")
check("줄마다 문단을 만들지 «않는다»", len(texts) == 3,
      f"문단 {len(texts)}개 -- 줄 단위로 쪼개면 5개가 된다")

print("\n=== add_journal_table ===")
doc2 = Document()
t = add_journal_table(doc2, [["항목", "값"], ["가", "1"], ["나", "2"]],
                      col_widths=[2.0, 1.0], font_size=9)
check("표가 붙는다", len(doc2.tables) == 1)
check("행 수가 맞는다", len(t.rows) == 3, f"실제 {len(t.rows)}")
_borders = t._tbl.xml
check("세로줄을 «넣지 않는다»(저널 관습)",
      'w:val="single"' not in _borders.split("insideV")[1][:200]
      if "insideV" in _borders else True)

print("\n=== 렌더러 -- 소스 검사 ===")
src = (HERE / "docx_kit.py").read_text(encoding="utf-8")
check("DispatchEx를 쓴다", re.search(r"win32\.DispatchEx\(", src) is not None)
check("맨몸 Dispatch를 «쓰지 않는다»",
      re.search(r"win32\.Dispatch\(", src) is None,
      "Dispatch는 떠 있는 인스턴스에 붙어 한 번 막히면 계속 실패한다")
check("원본이 아니라 사본을 연다", "shutil.copy2" in src)

print()
if fails:
    print(f"★ {len(fails)}건 실패: {', '.join(fails)}")
    sys.exit(1)
print("전부 통과")
