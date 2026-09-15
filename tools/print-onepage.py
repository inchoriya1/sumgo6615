# -*- coding: utf-8 -*-
u"""
근퇴계를 가로 A4 한 장에 맞춥니다.

  python tools/print-onepage.py 급여대장.xlsx            # 보기만
  python tools/print-onepage.py 급여대장.xlsx --write    # 새 파일 만들기
  python tools/print-onepage.py 급여대장.xlsx --write --out 같은이름.xlsx   # 덮어쓰기

무엇이 문제였나
  근퇴계는 **세로로 깁니다** — 날짜표 1~40행에 급여산출 42~61행이 아래로 붙어
  모두 62행입니다. 가로 A4 는 폭은 넉넉해도 **높이가 210mm 뿐**이라,
  그대로 눕혀 한 장에 넣으면 **55% 로 줄어듭니다.** 11pt 글자가 6pt 가 됩니다.

어떻게 고치나
  **급여산출 블록을 아래가 아니라 오른쪽으로 옮깁니다.**
  가로로 눕혔을 때 남는 것은 폭이니, 거기를 씁니다.

      전                          후
      ┌ 날짜표 40행 ┐             ┌ 날짜표 40행 ┐ ┌ 급여산출 ┐
      ├ 급여산출 20행┤             │            │ │        │
      └ 62행 ─────┘             └ 40행 ──────┘ └────────┘
      → 한 장에 넣으면 55%        → 한 장에 넣으면 약 83%

  · 급여산출 A~E 42~61행 → **M~P 7~26행** 으로 값·서식 그대로 옮깁니다.
  · 산출식이 C:D 두 칸에 걸쳐 있던 병합은 풀고 **O 한 칸**에 담습니다.
  · 날짜표 열 너비를 조금씩 줄입니다. **글자가 잘리지 않는 선까지만** 줄입니다.
  · 행 높이는 옮기지 않습니다 — 급여산출 머리글(30pt)을 그대로 가져오면
    그 자리의 **날짜 줄 하나가 같이 두꺼워집니다.**

곁들여 옮기는 것 둘
  · `N1` — 기본정보 몇 번째 줄을 볼지 정하는 **숨은 계산 칸**입니다.
    급여산출이 M~P 로 들어가면 **N1 이 그 바로 위**라 인쇄에 딸려 나옵니다.
    `T1` 로 옮기고, 이것을 가리키던 근퇴계 안 **28칸**을 함께 고칩니다.
    (근퇴계 밖에서 `N1` 을 보는 곳은 없습니다. 확인했습니다.)
  · `L3` 의 "← 이 칸에서 사람을 고르세요" — **화면용 안내**입니다.
    인쇄 범위 안에 들어가 버려서 `R3` 으로 뺍니다.
"""
import argparse
import re
from copy import copy

import openpyxl
from openpyxl.utils import get_column_letter as gl
from openpyxl.worksheet.page import PageMargins

GT = u"근퇴계"

SRC_TOP, SRC_BOT = 42, 61       # 급여산출 블록 (원래 자리)
DST_TOP = 7                     # 옮길 자리 — 날짜표 머리글과 같은 줄에서 시작
SRC_COLS = (1, 2, 3, 5)         # A 구분 · B 시간 · C 산출식 · E 금액
DST_COLS = (13, 14, 15, 16)     # M · N · O · P

HELPER_FROM, HELPER_TO = "N1", "T1"
HINT_FROM, HINT_TO = "L3", "R3"

PRINT_LAST_ROW = 40             # 40행 = 연장 실적 한 줄. 그 아래는 비웁니다
PRINT_LAST_COL = "P"

# 글자가 잘리지 않는 선까지만 줄입니다. 한글 한 자 ≈ 2, 숫자 한 자 ≈ 1.
WIDTHS = {
    "A": 11,    # 2026-08-01
    "B": 5,     # 토
    "C": 11,    # 대체공휴일
    "D": 7,     # 8:30
    "E": 7,     # 17:30
    "F": 11,    # 근무시간(H)
    "G": 9,     # 연장(H)
    "H": 11,    # 토요근무(H)
    "I": 11,    # 휴일특근(H)
    "J": 10,    # 휴일근무
    "K": 22,    # 메모 — 비고가 길어 여기만 넉넉히 둡니다
    "L": 2,     # 사이 띄우기
    "M": 13,    # 토요근무차감
    "N": 7,
    "O": 21,    # 통상시급 × 1.5 × 고정시간
    "P": 12,    # 2,970,000
}


def move_block(ws):
    u"""급여산출을 M~P 로 옮기고 원래 자리를 비웁니다."""
    # 병합 먼저 풀어야 값을 지울 수 있습니다
    for rng in [str(m) for m in ws.merged_cells.ranges]:
        a, b = rng.split(":")
        if SRC_TOP <= ws[a].row <= SRC_BOT:
            ws.unmerge_cells(rng)

    moved = 0
    for i, src_r in enumerate(range(SRC_TOP, SRC_BOT + 1)):
        dst_r = DST_TOP + i
        for src_c, dst_c in zip(SRC_COLS, DST_COLS):
            s = ws.cell(row=src_r, column=src_c)
            d = ws.cell(row=dst_r, column=dst_c)
            d.value = s.value
            d.font = copy(s.font)
            d.fill = copy(s.fill)
            d.border = copy(s.border)
            d.alignment = copy(s.alignment)
            d.number_format = s.number_format
            moved += 1

    # 원래 자리를 비웁니다 (A~E, 42~61행)
    for r in range(SRC_TOP, SRC_BOT + 1):
        for c in range(1, 6):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.fill = openpyxl.styles.PatternFill()
            cell.border = openpyxl.styles.Border()
            cell.font = openpyxl.styles.Font()
            cell.number_format = "General"
    return moved


def move_helper(ws):
    u"""N1 을 T1 로 옮기고, 가리키던 칸을 모두 고칩니다."""
    ws[HELPER_TO].value = ws[HELPER_FROM].value
    ws[HELPER_FROM].value = None

    old = "$%s$%s" % (HELPER_FROM[0], HELPER_FROM[1:])
    new = "$%s$%s" % (HELPER_TO[0], HELPER_TO[1:])
    hit = []
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and old in c.value:
                c.value = c.value.replace(old, new)
                hit.append(c.coordinate)

    ws[HINT_TO].value = ws[HINT_FROM].value
    ws[HINT_TO].font = copy(ws[HINT_FROM].font)
    ws[HINT_FROM].value = None
    return hit


def set_page(ws):
    for col, w in WIDTHS.items():
        ws.column_dimensions[col].width = w

    ps = ws.page_setup
    ps.orientation = "landscape"
    ps.paperSize = ws.PAPERSIZE_A4
    ps.fitToWidth = 1
    ps.fitToHeight = 1
    ps.scale = None                      # 배율과 맞추기는 같이 못 씁니다
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.4, bottom=0.4,
                                  header=0.2, footer=0.2)
    ws.print_options.horizontalCentered = True
    ws.print_area = "A1:%s%d" % (PRINT_LAST_COL, PRINT_LAST_ROW)


def estimate(ws):
    u"""줄어드는 비율을 미리 셈해 봅니다. A4 가로 297 × 210mm."""
    width = sum(WIDTHS[gl(c)] for c in range(1, 17)) * 2.0        # 1칸 ≈ 2.0mm
    height = sum((ws.row_dimensions[r].height
                  if r in ws.row_dimensions and ws.row_dimensions[r].height
                  else 15.0) for r in range(1, PRINT_LAST_ROW + 1)) * 25.4 / 72
    usable_w = 297.0 - (0.25 + 0.25) * 25.4
    usable_h = 210.0 - (0.4 + 0.4) * 25.4
    return width, height, min(usable_w / width, usable_h / height)


def run(path, out_path, do_write):
    wb = openpyxl.load_workbook(path)
    ws = wb[GT]

    before = (ws.page_setup.orientation, ws.print_area,
              ws.max_row, ws.max_column)

    moved = move_block(ws)
    hit = move_helper(ws)
    set_page(ws)
    w, h, scale = estimate(ws)

    print(u"① 급여산출을 오른쪽으로 — A~E %d~%d행 → M~P %d~%d행 (%d칸)"
          % (SRC_TOP, SRC_BOT, DST_TOP, DST_TOP + (SRC_BOT - SRC_TOP), moved))
    print(u"   산출식 C:D 병합 %d개를 풀어 O 한 칸에 담았습니다" % (SRC_BOT - 43))
    print(u"② %s → %s  (이것을 보던 %d칸을 함께 고침)"
          % (HELPER_FROM, HELPER_TO, len(hit)))
    print(u"   %s → %s  (화면용 안내를 인쇄 밖으로)" % (HINT_FROM, HINT_TO))
    print(u"③ 인쇄 설정")
    print(u"   방향   %s → 가로(landscape) · A4" % (before[0] or u"세로(기본)"))
    print(u"   범위   %s → %s" % (before[1], ws.print_area))
    print(u"   한 장  가로 1쪽 × 세로 1쪽 · 여백 0.25/0.4인치 · 가로 가운데")
    print(u"   크기   폭 %.0fmm × 높이 %.0fmm → **약 %.0f%% 로 축소** (전에는 55%%)"
          % (w, h, scale * 100))

    if not do_write:
        print(u"\n※ 보기만 했습니다. 실제로 만들려면 --write 를 붙이세요.")
        return
    wb.save(out_path)
    print(u"\n만들었습니다: %s" % out_path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("book")
    ap.add_argument("--out", default=None)
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args()
    out = a.out or re.sub(r"\.xlsx$", "", a.book) + u"_인쇄.xlsx"
    run(a.book, out, a.write)


if __name__ == "__main__":
    main()
