# -*- coding: utf-8 -*-
"""인용문 패치가 «고쳤는가»와 «다른 것을 깨뜨리지 않았는가»를 같이 본다.

⛔ 고치는 중에 새 결함을 만드는 것이 이 랩의 실측 최대 부류다 — 그래서 대조군을 넣는다.
"""
import sys

for _s in (sys.stdout, sys.stderr):
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")

from pathlib import Path

sys.path.insert(0, r"C:\Users\User\.claude\skills\docx-editing\scripts")
from md_to_docx import unwrap  # noqa: E402

CASES = [
    ## ① 실사고 — 강조가 인용문 줄을 넘는다
    ("인용문 안 강조가 줄을 넘음",
     ["> 그리고 이 사실이 잘 보이지 않는 이유는, 표준 지표인 **인구당 병원 수가 자원을",
      "> 연속량으로 가정**해 불가분성을 담지 못하기 때문이다."],
     lambda out: len(out) == 1 and "**인구당 병원 수가 자원을 연속량으로 가정**" in out[0]),

    ## ② 대조군 — 인용문 «안의» 문단 구분('>' 만 있는 줄)은 «지켜야» 한다
    ("인용문 안 문단 구분은 유지",
     ["> 첫 문단이다.", ">", "> 둘째 문단이다."],
     lambda out: len(out) == 3),

    ## ③ 대조군 — 서로 다른 인용문 두 개(빈 줄로 분리)를 붙이면 안 된다
    ("빈 줄로 갈린 인용문 둘",
     ["> 앞 인용문.", "", "> 뒤 인용문."],
     lambda out: len(out) == 3),

    ## ④ 대조군 — 목록 이어붙이기가 여전히 되는가(기존 기능)
    ("목록 연속줄(기존 기능)",
     ["- **MCPS", "  publication** 항목"],
     lambda out: len(out) == 1 and "**MCPS publication**" in out[0]),
]

bad = 0
for name, lines, ok in CASES:
    out = unwrap(lines)
    good = ok(out)
    bad += 0 if good else 1
    print("%s  %s" % ("✅" if good else "⛔ 실패", name))
    if not good:
        for o in out:
            print("      %r" % o)

print()
if bad:
    sys.exit("⛔ %d개 실패" % bad)

## 실사고가 난 실제 문서로 한 번 더 — «있으면» 본다.
## 이 스킬은 랩 공용이므로 특정 프로젝트 파일에 «의존»하면 안 된다(없으면 건너뛴다).
src = Path(r"D:\Onedrive\2026-2\보건의료경제학\과제_병원폐업\report\01_intro.md")
if not src.exists():
    print("(실문서 확인 건너뜀 — %s 없음)" % src.name)
    print("\n\u2705 4/4 — 고쳤고, 대조군 셋은 그대로다")
    raise SystemExit(0)
joined = unwrap(src.read_text(encoding="utf-8").split("\n"))
stray = [l for l in joined
         if l.lstrip().startswith(">") and l.count("**") % 2 == 1]
print("보고서 서론 인용문 중 «짝이 안 맞는 별표»: %d줄" % len(stray))
for l in stray:
    print("   %s" % l[:70])
print("\n✅ 4/4 — 고쳤고, 대조군 셋은 그대로다")
