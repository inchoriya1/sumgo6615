# -*- coding: utf-8 -*-
u"""
근태원본 출근·퇴근을 '글자' 에서 '진짜 시각' 으로 바꿉니다.

  python tools/text-to-time.py 급여대장.xlsx            # 보기만
  python tools/text-to-time.py 급여대장.xlsx --write    # 새 파일 만들기
  python tools/text-to-time.py 급여대장.xlsx --write --out 같은이름.xlsx   # 덮어쓰기

무엇이 문제였나
  근태기록 .xls 에서 복사해 온 출퇴근이 **글자로 들어와 있습니다.**
  칸에는 `07:49` 라고 보이지만 그것은 **"07:49" 라는 글자**이지 시각이 아닙니다.

  급여는 지금도 맞게 나옵니다 — 엑셀이 뺄셈할 때 글자를 숫자로
  알아서 바꿔 주기 때문입니다(`"16:32" - TIME(8,30,0)` 이 계산됩니다).
  그래서 **이 작업으로 급여 숫자는 한 원도 바뀌지 않습니다.**

  다만 글자로 남아 있으면 —
  · 시각으로 **정렬·필터**하면 뜻대로 안 됩니다 (글자 순서로 줄을 섭니다)
  · `h:mm` **서식이 아무 일도 안 합니다** — 이미 글자라 입힐 것이 없습니다
  · [근퇴계] 가 이 값을 끌어다 쓰는데, 앞으로 **숫자 시각**이 섞여 들어오면
    그때부터 `0.354166666666667` 처럼 보이기 시작합니다

어떻게 고치나
  `"07:49"` → `0.32569…` (하루의 몇 분의 몇). 엑셀이 시각을 담는 방식 그대로입니다.
  서식은 이미 `h:mm` 이라 **화면에 보이는 모양은 똑같습니다.**
  바뀌는 것은 **속**입니다.

무엇을 건드리지 않나
  · **D·E 두 열만** 바꿉니다. F~J·M 은 비어 있고, O(지각차감)는 원래 숫자입니다.
  · 수식 칸은 손대지 않습니다.
  · `H:MM` 으로 읽히지 않는 값은 **건너뛰고 그대로 둡니다** — 끝에 알려 줍니다.
"""
import argparse
import re

import openpyxl

RAW = u"근태원본"
FIRST, LAST = 2, 5953
COLS = ((4, u"출근"), (5, u"퇴근"))
TIME_FMT = u"h:mm"

# 07:49 · 7:49 · 07:49:30 — 24시를 넘는 값(철야)도 받아 둡니다
PAT = re.compile(r"^\s*(\d{1,3}):([0-5]\d)(?::([0-5]\d))?\s*$")


def to_fraction(text):
    m = PAT.match(text)
    if not m:
        return None
    h, mi, s = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
    return (h * 3600 + mi * 60 + s) / 86400.0


def run(path, out_path, do_write):
    wb = openpyxl.load_workbook(path)
    ws = wb[RAW]

    done, skipped, sample = 0, [], []
    for col, label in COLS:
        for r in range(FIRST, LAST + 1):
            c = ws.cell(row=r, column=col)
            v = c.value
            if not isinstance(v, str) or v == "":
                continue
            if v.startswith("="):            # 수식은 건드리지 않습니다
                continue
            frac = to_fraction(v)
            if frac is None:
                skipped.append((c.coordinate, v))
                continue
            if len(sample) < 3:
                sample.append((c.coordinate, v, frac))
            c.value = frac
            c.number_format = TIME_FMT
            done += 1

    print(u"출근·퇴근 %d칸을 진짜 시각으로 바꿨습니다 (서식 %s 는 그대로)"
          % (done, TIME_FMT))
    for coord, before, after in sample:
        print(u"   %-7s '%s'(글자)  →  %.10f (시각)" % (coord, before, after))
    if skipped:
        print(u"★ H:MM 으로 읽히지 않아 그대로 둔 칸 %d개: %s"
              % (len(skipped), skipped[:5]))
    else:
        print(u"   읽지 못해 남긴 칸: 없음")
    print(u"\n※ 급여 숫자는 바뀌지 않습니다 — 엑셀은 원래도 글자를 숫자로 바꿔 계산했습니다.")

    if not do_write:
        print(u"※ 보기만 했습니다. 실제로 만들려면 --write 를 붙이세요.")
        return
    wb.save(out_path)
    print(u"만들었습니다: %s" % out_path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("book")
    ap.add_argument("--out", default=None)
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args()
    out = a.out or re.sub(r"\.xlsx$", "", a.book) + u"_시각.xlsx"
    run(a.book, out, a.write)


if __name__ == "__main__":
    main()
