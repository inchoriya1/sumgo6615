# -*- coding: utf-8 -*-
u"""
장마다 종합문제를 한 개씩 만듭니다.

  python tools/build-final-exam.py            # 보기만
  python tools/build-final-exam.py --write    # 실제로 만들기

복습 파일과 무엇이 다른가
  `N장_복습.xlsx` 은 **기술을 하나씩** 확인합니다 — 시트마다 한 강의.
  종합문제는 **엉망인 표 한 장**을 주고, 그 장에서 배운 것을 **전부 동원해**
  쓸 수 있는 표로 만들게 합니다. 실무는 기술이 하나씩 오지 않습니다.

  · `1장_종합문제.xlsx`  새봄전자 3월 수주현황 — 채우기·오류·연산자·참조
  · `2장_종합문제.xlsx`  누리유통 지점실적 — 병합·빈칸·나누기·중복·이름

순서가 곧 함정입니다
  앞 단계를 건너뛰면 뒤가 **조용히 틀립니다.** 오류가 안 뜹니다.
  그래서 [최종확인] 다섯 칸만 맞으면 앞 단계는 다 맞은 것이고,
  틀린 번호로 **어느 단계에서 샜는지** 알 수 있게 해 두었습니다.
"""
import argparse
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

FONT = u"맑은 고딕"
INK_HEAD = u"FF1F4E79"
YELLOW = u"FFFFE699"
ORANGE = u"FFF8CBAD"
NOTE = u"FFFFF9E6"
GREY = u"FF808080"
THIN = Border(*(Side(style="thin"),) * 4)

FIRST = 6


def put(ws, r, c, v, fill=None, bold=False, fmt=None, wrap=False, align="center"):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(name=FONT, size=11, bold=bold)
    cell.border = THIN
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)
    if fmt:
        cell.number_format = fmt
    return cell


def head(ws, r, c0, labels):
    f = Font(name=FONT, size=9, bold=True, color="FFFFFFFF")
    fill = PatternFill("solid", fgColor=INK_HEAD)
    for i, lab in enumerate(labels):
        cell = ws.cell(row=r, column=c0 + i, value=lab)
        cell.font, cell.fill, cell.border = f, fill, THIN
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def sheet_guide(wb, title_text, sub, body, height=430):
    ws = wb.create_sheet(u"안내")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 104
    t = ws.cell(row=1, column=2, value=title_text)
    t.font = Font(name=FONT, size=18, bold=True, color=INK_HEAD)
    s = ws.cell(row=2, column=2, value=sub)
    s.font = Font(name=FONT, size=10, color=GREY)
    ws.merge_cells("B4:B32")
    c = ws.cell(row=4, column=2, value=body)
    c.font = Font(name=FONT, size=11)
    c.fill = PatternFill("solid", fgColor=NOTE)
    c.alignment = Alignment(vertical="top", wrap_text=True)
    c.border = THIN
    ws.row_dimensions[4].height = height


def sheet_check(wb, questions):
    ws = wb.create_sheet(u"최종확인")
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 3, "B": 46, "C": 18}.items():
        ws.column_dimensions[col].width = w
    t = ws.cell(row=1, column=2, value=u"최종확인")
    t.font = Font(name=FONT, size=16, bold=True, color=INK_HEAD)
    s = ws.cell(row=2, column=2,
                value=u"주황 칸에 답을 적으세요. 다섯 칸이 맞으면 앞 단계는 다 맞은 것입니다.")
    s.font = Font(name=FONT, size=10, color=GREY)
    head(ws, 4, 2, [u"묻는 것", u"내 답"])
    for i, (q, _, fmt, _) in enumerate(questions):
        r = 5 + i
        put(ws, r, 2, q, align="left")
        put(ws, r, 3, None, fill=ORANGE, fmt=fmt)
    n = ws.cell(row=5 + len(questions) + 1, column=2,
                value=u"다 적었으면 [정답] 시트를 여세요.")
    n.font = Font(name=FONT, size=10, color=GREY)


def sheet_answer(wb, steps, questions):
    ws = wb.create_sheet(u"정답")
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 3, "B": 5, "C": 18, "D": 9, "E": 76}.items():
        ws.column_dimensions[col].width = w
    t = ws.cell(row=1, column=2, value=u"정답")
    t.font = Font(name=FONT, size=16, bold=True, color=INK_HEAD)
    s = ws.cell(row=2, column=2, value=u"먼저 다 풀고 나서 보세요.")
    s.font = Font(name=FONT, size=10, color=GREY)

    head(ws, 4, 2, [u"단계", u"무엇", u"강의", u"어떻게"])
    for i, (no, what, lec, how) in enumerate(steps):
        r = 5 + i
        put(ws, r, 2, no); put(ws, r, 3, what, align="left"); put(ws, r, 4, lec)
        put(ws, r, 5, how, align="left", wrap=True)
        ws.row_dimensions[r].height = 32

    r0 = 5 + len(steps) + 1
    h = ws.cell(row=r0, column=2, value=u"최종확인 답")
    h.font = Font(name=FONT, size=12, bold=True)
    head(ws, r0 + 1, 2, [u"번호", u"답", u"", u"틀렸다면"])
    rows = []
    for i, (q, a, fmt, why) in enumerate(questions):
        r = r0 + 2 + i
        put(ws, r, 2, u"%d" % (i + 1))
        put(ws, r, 3, a, bold=True, fmt=fmt)
        put(ws, r, 4, None)
        put(ws, r, 5, why, align="left", wrap=True)
        ws.row_dimensions[r].height = 30
        rows.append(r)
    return rows


# ══════════ 1장 ══════════════════════════════════════════════════
C1_ROWS = [
    (u"2026-03-02", u"한빛전자 (HB.Kim)",   120, 45000),
    (u"2026-03-03", u"미래상사 (MR.Lee)",    95, 45000),
    (u"2026-03-04", u"대성물산 (DS.Park)",  140, 78000),
    (u"2026-03-05", u"신영테크 (SY.Choi)",  160, 78000),
    (u"2026-03-06", u"우진산업 (WJ.Jung)",  210, 32000),
    (u"2026-03-07", u"가온시스템 (GO.Kim)", 175, 32000),
]
C1_TEXT = (1, 3, 5)          # 수량을 글자로 넣을 줄
C1_RATE, C1_GOAL = 0.03, 7000000

C1_GUIDE = u"""■ 상황

  새봄전자 2026년 3월 수주현황입니다. 앞자리 두 줄만 적혀 있고
  나머지는 **손대다 만 상태**입니다. 계산이 안 됩니다.

  [문제] 를 채우고 [최종확인] 의 다섯 칸에 답을 적으세요.

■ 여덟 단계 — 순서대로

  ① 번호 채우기        B6·B7 을 잡고 끌어 내립니다                (1-05)
  ② 수주일 채우기      C6·C7 을 잡고 끌어 내립니다                 (1-05)
  ③ 영문 이름 뽑기     F열에 괄호 안 영문만                        (1-07)
  ④ 숫자 고치기        G열 수량 중 글자로 된 것을 숫자로            (1-11)
  ⑤ 금액 내기          I열 = 수량 × 단가                         (1-10)
  ⑥ 목표 달성          J열 = 금액이 목표(C4) 이상인가 → TRUE/FALSE  (1-10)
  ⑦ 커미션 내기        K열 = 금액 × 커미션비율(C3)                 (1-09)
  ⑧ 최종확인 채우기    [최종확인] 다섯 칸                          —

■ 순서를 지켜야 하는 이유

  · ④를 건너뛰면 ⑤의 곱셈이 글자 × 숫자가 되어 세 줄이 샙니다.
  · ⑥·⑦에서 $ 를 안 붙이면 아래로 갈수록 엉뚱해집니다.
  · ⑥은 IF 가 아닙니다. 비교연산자 하나면 TRUE/FALSE 가 나옵니다.

■ 채점

  노란 칸 = 채울 곳 · 주황 칸 = 답을 적는 곳 · 흰 칸은 자료입니다."""

C1_STEPS = [
    (u"①", u"번호 채우기", u"1-05", u"B6 에 1, B7 에 2 가 있습니다. 두 칸을 함께 잡고 B11 까지 끕니다."),
    (u"②", u"수주일 채우기", u"1-05", u"C6·C7 두 칸을 잡고 C11 까지 끕니다. 하루씩 늘어납니다."),
    (u"③", u"영문 이름", u"1-07", u"F6 에 HB.Kim 을 직접 치고 Ctrl+E. 괄호 안만 뽑힙니다."),
    (u"④", u"숫자 고치기", u"1-11", u"G7 · G9 · G11 이 글자입니다. G6:G11 을 잡고 느낌표 → 「숫자로 변환」."),
    (u"⑤", u"금액", u"1-10", u"I6 = G6*H6 을 넣고 I11 까지 끕니다."),
    (u"⑥", u"목표 달성", u"1-10", u"J6 = I6>=$C$4 입니다. IF 가 아니라 비교연산자만 쓰면 TRUE/FALSE 가 나옵니다."),
    (u"⑦", u"커미션", u"1-09", u"K6 = I6*$C$3 을 넣고 K11 까지. $ 를 빠뜨리면 아래가 0 이 됩니다."),
    (u"⑧", u"최종확인", u"—", u"[최종확인] 주황 칸 다섯 개를 채웁니다."),
]

C1_Q = [
    (u"① 수량 합계는?", 900, u"0",
     u"④를 건너뛰면 글자로 된 세 줄이 빠져 470 만 나옵니다."),
    (u"② 금액 합계는?", 45395000, u"#,##0",
     u"④를 건너뛰면 세 줄이 0 이 되어 23,040,000 으로 샙니다."),
    (u"③ 커미션 합계(3%)는?", 1361850, u"#,##0",
     u"⑦에서 $ 를 안 붙이면 아래로 갈수록 0 이 되어 훨씬 작습니다."),
    (u"④ 목표(700만) 이상인 건수는?", 2, u"0",
     u"대성물산·신영테크 둘뿐입니다. J열의 TRUE 를 세면 됩니다."),
    (u"⑤ 거래처 성(영문)이 Kim 인 곳은?", 2, u"0",
     u"③을 해야 셀 수 있습니다. HB.Kim · GO.Kim 둘."),
]


def build_ch1():
    wb = Workbook(); wb.remove(wb.active)
    sheet_guide(wb, u"1장 종합문제 — 새봄전자 3월 수주현황",
                u"1장에서 배운 채우기·빠른채우기·오류·연산자·셀참조를 한 표에서 씁니다.",
                C1_GUIDE)
    ws = wb.create_sheet(u"문제")
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 3, "B": 7, "C": 13, "D": 3, "E": 20, "F": 13,
                   "G": 9, "H": 11, "I": 14, "J": 12, "K": 13}.items():
        ws.column_dimensions[col].width = w
    t = ws.cell(row=1, column=2, value=u"새봄전자 2026년 3월 수주현황")
    t.font = Font(name=FONT, size=14, bold=True)
    put(ws, 3, 2, u"커미션 비율", bold=True, align="left"); put(ws, 3, 3, C1_RATE, fmt=u"0.0%")
    put(ws, 4, 2, u"목표 금액", bold=True, align="left"); put(ws, 4, 3, C1_GOAL, fmt=u"#,##0")

    head(ws, 5, 2, [u"번호", u"수주일"])
    head(ws, 5, 5, [u"거래처", u"거래처영문", u"수량", u"단가", u"금액", u"목표달성", u"커미션"])
    for i, (d, nm, qty, price) in enumerate(C1_ROWS):
        r = FIRST + i
        if i < 2:                                  # 앞 두 줄만 주어집니다
            put(ws, r, 2, i + 1)
            put(ws, r, 3, d, fmt=u"yyyy-mm-dd")
        else:
            put(ws, r, 2, None, fill=YELLOW)
            put(ws, r, 3, None, fill=YELLOW, fmt=u"yyyy-mm-dd")
        put(ws, r, 5, nm, align="left")
        put(ws, r, 6, None, fill=YELLOW)
        put(ws, r, 7, u"%d" % qty if i in C1_TEXT else qty, fmt=u"#,##0")
        put(ws, r, 8, price, fmt=u"#,##0")
        for c in (9, 10, 11):
            put(ws, r, c, None, fill=YELLOW, fmt=u"#,##0" if c != 10 else u"General")
    last = FIRST + len(C1_ROWS) - 1
    put(ws, last + 2, 5, u"합계", bold=True)
    put(ws, last + 2, 9, u"=SUM(I%d:I%d)" % (FIRST, last), bold=True, fmt=u"#,##0")
    put(ws, last + 2, 11, u"=SUM(K%d:K%d)" % (FIRST, last), bold=True, fmt=u"#,##0")
    g = ws.cell(row=3, column=6, value=u"노란 칸을 채우세요. [안내] 의 여덟 단계를 순서대로.")
    g.font = Font(name=FONT, size=10, color=GREY)

    sheet_check(wb, C1_Q)
    rows = sheet_answer(wb, C1_STEPS, C1_Q)
    return wb, rows


# ══════════ 2장 ══════════════════════════════════════════════════
C2_ROWS = [
    (u"수도권", u"강남점",   u"김세민\n정다온", u"SU-01-2026", 4200, 3800, 4500),
    (None,     u"신촌점",   u"박정화",         u"SU-02-2026", 3100, 3300, 2900),
    (None,     u"구로점",   u"이서우\n최리",    u"SU-03-2026", 2800, 2600, 3000),
    (u"영남",   u"해운대점", u"김준용",         u"YN-01-2026", 3600, 3900, 3700),
    (None,     u"서면점",   u"정진하\n김병민",  u"YN-02-2026", 2400, 2500, 2200),
    (u"호남",   u"상무점",   u"박희선",         u"HN-01-2026", 1900, 2100, 2000),
    (None,     u"수완점",   u"김수호\n이유림",  u"HN-01-2026", 1700, 1800, 1600),
]

C2_GUIDE = u"""■ 상황

  누리유통 1분기 지점실적입니다. **보기에는 잘 정리된 표**인데
  정렬도 필터도 안 되고, 합계가 맞지 않습니다.

  [문제] 를 쓸 수 있는 표로 바꾸고 [최종확인] 다섯 칸에 답을 적으세요.

■ 여덟 단계 — 순서대로

  ① 병합 풀기          B열 지역의 병합을 풉니다                    (2-03)
  ② 빈 칸 채우기        드러난 빈 칸을 위 값으로 채웁니다            (2-04)
  ③ 담당자 나누기       D열 한 칸 두 줄을 E·F 두 칸으로             (2-13)
  ④ 지점코드 나누기     H열을 I·J·K 세 칸으로                      (2-12)
  ⑤ 중복 찾기          H열에서 겹치는 지점코드를 찾습니다            (2-09)
  ⑥ 이름 붙이기        L6:N12 에 「매출」 이라는 이름을 붙입니다       (2-10)
  ⑦ 합계 내기          O열 = 1~3월 합계. 맨 아래는 =SUM(매출)       (2-10)
  ⑧ 최종확인 채우기    [최종확인] 다섯 칸                          —

■ 순서를 지켜야 하는 이유

  · ②를 건너뛰면 지역별 합계가 첫 줄만 잡힙니다.
  · ③을 안 하면 담당자가 몇 명인지 셀 수 없습니다.
  · ⑤는 눈으로 찾지 마세요. 조건부 서식 → 중복 값 을 쓰면 한 번에 보입니다.

■ 채점

  노란 칸 = 채울 곳 · 주황 칸 = 답을 적는 곳 · 흰 칸은 자료입니다."""

C2_STEPS = [
    (u"①", u"병합 풀기", u"2-03", u"B6:B12 를 잡고 「병합하고 가운데 맞춤」 을 다시 눌러 끕니다."),
    (u"②", u"빈 칸 채우기", u"2-04",
     u"B6:B12 를 잡고 F5 → 빈 셀 → = 치고 ↑ → Ctrl+Enter. 그다음 값으로 굳힙니다."),
    (u"③", u"담당자 나누기", u"2-13",
     u"D6:D12 를 잡고 텍스트 나누기 → 구분 기호 → 기타 에서 Ctrl+J. 화면에는 안 보입니다."),
    (u"④", u"지점코드 나누기", u"2-12",
     u"H6:H12 를 잡고 텍스트 나누기 → 구분 기호 → 기타 에 - 를 넣습니다."),
    (u"⑤", u"중복 찾기", u"2-09",
     u"H6:H12 를 잡고 조건부 서식 → 셀 강조 규칙 → 중복 값. 한 쌍이 칠해집니다."),
    (u"⑥", u"이름 붙이기", u"2-10", u"L6:N12 를 잡고 이름 상자에 매출 이라고 치고 Enter."),
    (u"⑦", u"합계", u"2-10", u"O6 = SUM(L6:N6) 을 끌고, 맨 아래 총합은 =SUM(매출)."),
    (u"⑧", u"최종확인", u"—", u"[최종확인] 주황 칸 다섯 개를 채웁니다."),
]

C2_Q = [
    (u"① 수도권 세 지점의 1~3월 합계는?", 30200, u"#,##0",
     u"②를 건너뛰면 강남점만 잡혀 12,500 이 나옵니다."),
    (u"② 일곱 지점 전체 합계는?", 59600, u"#,##0",
     u"⑥·⑦이 맞으면 =SUM(매출) 과 O열 합계가 같아야 합니다."),
    (u"③ 중복된 지점코드는 몇 건?", 1, u"0",
     u"HN-01-2026 이 두 번 있습니다. 한 쌍이니 1건입니다."),
    (u"④ 담당자는 모두 몇 명?", 11, u"0",
     u"③을 해야 셀 수 있습니다. 두 명인 지점이 넷, 한 명인 지점이 셋."),
    (u"⑤ 지점코드 앞 두 글자는 몇 종류?", 3, u"0",
     u"④를 해야 보입니다. SU · YN · HN 셋."),
]


def build_ch2():
    wb = Workbook(); wb.remove(wb.active)
    sheet_guide(wb, u"2장 종합문제 — 누리유통 1분기 지점실적",
                u"2장에서 배운 병합·빈칸·나누기·중복·이름을 한 표에서 씁니다.",
                C2_GUIDE)
    ws = wb.create_sheet(u"문제")
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 3, "B": 10, "C": 11, "D": 13, "E": 11, "F": 11, "G": 3,
                   "H": 14, "I": 7, "J": 7, "K": 8, "L": 9, "M": 9, "N": 9,
                   "O": 11}.items():
        ws.column_dimensions[col].width = w
    t = ws.cell(row=1, column=2, value=u"누리유통 2026년 1분기 지점실적")
    t.font = Font(name=FONT, size=14, bold=True)
    put(ws, 3, 2, u"단위 : 만원", bold=True, align="left")

    head(ws, 5, 2, [u"지역", u"지점", u"담당자", u"담당자1", u"담당자2"])
    head(ws, 5, 8, [u"지점코드", u"지역", u"번호", u"연도", u"1월", u"2월", u"3월", u"합계"])
    for i, (reg, br, who, code, a, b, c) in enumerate(C2_ROWS):
        r = FIRST + i
        put(ws, r, 2, reg)
        put(ws, r, 3, br)
        cell = put(ws, r, 4, who)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        put(ws, r, 5, None, fill=YELLOW); put(ws, r, 6, None, fill=YELLOW)
        put(ws, r, 8, code)
        for cc in (9, 10, 11):
            put(ws, r, cc, None, fill=YELLOW)
        for j, v in enumerate((a, b, c)):
            put(ws, r, 12 + j, v, fmt=u"#,##0")
        put(ws, r, 15, None, fill=YELLOW, fmt=u"#,##0")
        ws.row_dimensions[r].height = 30
    ws.merge_cells(start_row=FIRST, start_column=2, end_row=FIRST + 2, end_column=2)
    ws.merge_cells(start_row=FIRST + 3, start_column=2, end_row=FIRST + 4, end_column=2)
    ws.merge_cells(start_row=FIRST + 5, start_column=2, end_row=FIRST + 6, end_column=2)

    last = FIRST + len(C2_ROWS) - 1
    put(ws, last + 2, 2, u"전체 합계", bold=True)
    put(ws, last + 2, 15, None, fill=YELLOW, bold=True, fmt=u"#,##0")
    g = ws.cell(row=3, column=8, value=u"노란 칸을 채우세요. [안내] 의 여덟 단계를 순서대로.")
    g.font = Font(name=FONT, size=10, color=GREY)

    sheet_check(wb, C2_Q)
    rows = sheet_answer(wb, C2_STEPS, C2_Q)
    return wb, rows


# ══════════ 3장 ══════════════════════════════════════════════════
# (지점, 1월~6월 매출, 목표)
C3_ROWS = [
    (u"강남점", [31_450_000, 28_900_000, 0, 33_100_000, 35_600_000, 30_200_000], 190_000_000),
    (u"서초점", [24_300_000, 26_700_000, 25_100_000, 0, 27_900_000, 26_400_000], 150_000_000),
    (u"명동점", [18_200_000, 17_500_000, 19_800_000, 20_100_000, 18_900_000, 19_300_000], 110_000_000),
    (u"일산점", [12_600_000, 13_100_000, 0, 14_200_000, 13_800_000, 12_900_000], 80_000_000),
    (u"용산점", [29_700_000, 31_200_000, 30_500_000, 28_800_000, 0, 32_100_000], 200_000_000),
]
C3_SUM = [(nm, sum(ms), goal) for nm, ms, goal in C3_ROWS]
C3_TOTAL = sum(s for _, s, _ in C3_SUM)
C3_UNDER = sum(1 for _, s, g in C3_SUM if s < g)
C3_TOP = max(C3_SUM, key=lambda x: x[1])[0]
C3_ZEROS = sum(1 for _, ms, _ in C3_ROWS for v in ms if v == 0)

C3_GUIDE = u"""■ 상황

  한빛물산 상반기 영업보고입니다. **값은 다 들어 있습니다.**
  그런데 아무것도 안 읽힙니다 — 자릿수가 열두 자리이고,
  0 이 잔뜩이고, 어디를 봐야 할지 알 수가 없습니다.

  숫자는 **하나도 고치지 마세요.** 보이는 방식만 바꿉니다.

■ 여덟 단계 — 순서대로

  ① 천 단위 기호      숫자 열 전부에 #,##0                        (3-03)
  ② 단위 접기         맨 뒤 쉼표로 천 단위 절삭                     (3-03)
  ③ 안내 문구 고치기  표 위 「(단위: 원)」 을 고칩니다               (3-03)
  ④ 0 감추기          0 을 하이픈으로 보이게                       (3-02)
  ⑤ 달성률            J열 = 상반기합계 ÷ 목표, 0.0% 서식            (3-02)
  ⑥ 줄 전체 강조      달성률 100% 미만인 줄 전체를 칠합니다          (3-04)
  ⑦ 막대와 트렌드     I열에 데이터 막대, K열에 스파크라인            (3-05·3-07)
  ⑧ 최종확인 채우기   [최종확인] 다섯 칸                            —

■ 순서를 지켜야 하는 이유

  · ②를 하고 ③을 안 하면 **천 배 틀린 보고서**가 나갑니다.
    1,250 이라고 적힌 값을 읽는 사람은 천이백오십 원으로 읽습니다.
  · ⑥에서 $ 를 열 앞에만 안 붙이면 줄 전체가 안 칠해집니다.
  · ⑦의 막대 범위에 합계 줄을 넣으면 나머지가 전부 짧아집니다.

■ 채점

  노란 칸 = 채울 곳 · 주황 칸 = 답을 적는 곳 · 흰 칸은 자료입니다.
  ①~④는 서식이라 값이 안 바뀝니다. 아무 칸이나 눌러
  **수식 입력줄에 원래 숫자가 그대로 있는지** 확인하세요."""

C3_STEPS = [
    (u"①", u"천 단위 기호", u"3-03", u"C6:I10 을 잡고 Ctrl+Shift+1. 또는 서식에 #,##0."),
    (u"②", u"단위 접기", u"3-03", u"같은 범위에 #,##0, 를 넣습니다. 맨 뒤 쉼표 하나가 천으로 나눕니다."),
    (u"③", u"안내 문구", u"3-03", u"표 위 C3 을 「(단위: 천원)」 으로 고칩니다. 이것을 빼면 천 배 틀립니다."),
    (u"④", u"0 감추기", u"3-02", u"#,##0,;-#,##0,;-;@ — 셋째 칸이 0 일 때입니다."),
    (u"⑤", u"달성률", u"3-02", u"J6 = I6/H6, 서식은 0.0%;[빨강]-0.0%;-;@ 입니다."),
    (u"⑥", u"줄 전체 강조", u"3-04", u"B6:J10 을 전부 잡고 새 규칙 → 수식 → =$J6<1."),
    (u"⑦", u"막대·트렌드", u"3-05", u"I6:I10 에 데이터 막대. K6 에 스파크라인(범위 C6:H6) 을 넣고 아래로 끕니다."),
    (u"⑧", u"최종확인", u"—", u"[최종확인] 주황 칸 다섯 개를 채웁니다."),
]

C3_Q = [
    (u"① 상반기 매출 합계는? (원)", C3_TOTAL, u"#,##0",
     u"서식은 값을 안 바꿉니다. ①~④를 다 해도 합계는 그대로여야 합니다."),
    (u"② 달성률이 100% 미만인 지점은 몇 곳?", C3_UNDER, u"0",
     u"⑤를 해야 셉니다. ⑥에서 칠해진 줄을 세도 같습니다."),
    (u"③ 상반기 매출이 가장 큰 지점은?", C3_TOP, u"General",
     u"I열 합계를 견주면 됩니다. 달성률과는 다른 답입니다 — 목표가 다르니까요."),
    (u"④ 매출이 0 인 칸은 몇 개?", C3_ZEROS, u"0",
     u"④에서 하이픈으로 바뀐 칸을 세면 됩니다. 값이 지워진 것이 아닙니다."),
    (u"⑤ ②를 한 뒤 표 위에 적어야 할 말은?", u"(단위: 천원)", u"General",
     u"이것을 안 적으면 읽는 사람은 천 배 작게 읽습니다. 이 장에서 제일 큰 사고입니다."),
]


def build_ch3():
    wb = Workbook(); wb.remove(wb.active)
    sheet_guide(wb, u"3장 종합문제 — 한빛물산 상반기 영업보고",
                u"값은 하나도 안 고칩니다. 보이는 방식만 바꿔 읽히게 만듭니다.",
                C3_GUIDE, height=470)
    ws = wb.create_sheet(u"문제")
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 3, "B": 12, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13,
                   "H": 13, "I": 15, "J": 11, "K": 14}.items():
        ws.column_dimensions[col].width = w
    t = ws.cell(row=1, column=2, value=u"한빛물산 2026 상반기 영업보고")
    t.font = Font(name=FONT, size=14, bold=True)
    u = ws.cell(row=3, column=3, value=u"(단위: 원)")
    u.font = Font(name=FONT, size=10, color=GREY)
    head(ws, 5, 2, [u"지점", u"1월", u"2월", u"3월", u"4월", u"5월", u"6월",
                    u"상반기 목표", u"상반기 합계", u"달성률", u"트렌드"])
    for i, (nm, ms, goal) in enumerate(C3_ROWS):
        r = FIRST + i
        put(ws, r, 2, nm)
        for j, v in enumerate(ms):
            put(ws, r, 3 + j, v)
        put(ws, r, 9, goal)
        put(ws, r, 10, sum(ms))
        put(ws, r, 11, None, fill=YELLOW)
        put(ws, r, 12, None, fill=YELLOW)
    last = FIRST + len(C3_ROWS) - 1
    put(ws, last + 2, 2, u"합계", bold=True)
    put(ws, last + 2, 10, C3_TOTAL, bold=True)
    g = ws.cell(row=3, column=9, value=u"노란 칸을 채우세요. [안내] 의 여덟 단계를 순서대로.")
    g.font = Font(name=FONT, size=10, color=GREY)
    sheet_check(wb, C3_Q)
    sheet_answer(wb, C3_STEPS, C3_Q)
    return wb, None


# ══════════ 4장 ══════════════════════════════════════════════════
C4_PRICE = [(u"OPD011", u"진라면", 780), (u"OPD023", u"오라면", 830),
            (u"OPD038", u"진짬뽕", 920), (u"OPD047", u"진짜장", 920),
            (u"OPD057", u"김치라면", 850)]
C4_MAP = dict((c, (n, p)) for c, n, p in C4_PRICE)
# (제품코드, 수량) — 수량이 글자인 줄과 제품표에 없는 코드가 섞여 있습니다
C4_ROWS = [
    (u"OPD023", 50, False), (u"OPD011", 20, True), (u"OPD038", 50, False),
    (u"OPD099", 25, False), (u"OPD047", 25, True), (u"OPD057", 35, False),
    (u"OPD088", 30, False), (u"OPD011", 35, True),
]
C4_QTY = sum(q for _, q, _ in C4_ROWS)
C4_AMT = sum(q * C4_MAP[c][1] for c, q, _ in C4_ROWS if c in C4_MAP)
C4_MISS = sum(1 for c, _, _ in C4_ROWS if c not in C4_MAP)
C4_JUNK_ROW = 480

C4_GUIDE = u"""■ 상황

  대림기업 3월 견적서입니다. **내일 아침에 나가야 합니다.**

  · 수량 중 세 줄이 **글자**로 들어와 있습니다
  · 제품표에 **없는 코드**가 섞여 있습니다
  · 아래쪽에 **빈 서식이 수백 행** 남아 파일이 무겁습니다
  · 인쇄하면 **여러 장으로 갈라지고** 둘째 장에 머리글이 없습니다

■ 여덟 단계 — 순서대로

  ① 빈 서식 지우기    Ctrl+End 로 확인 → 자료 아래 행 전체 삭제 → 저장   (4-07)
  ② 유효성 검사       제품코드는 목록에서, 수량은 정수 0 이상            (4-08)
  ③ 오류 메시지       「수량은 0보다 큰 정수만 입력 가능합니다」          (4-08)
  ④ 글자 숫자 고치기  수량 세 줄을 숫자로                              (1-11)
  ⑤ 단가·금액         INDEX/MATCH 또는 VLOOKUP + IFERROR               (7-04)
  ⑥ 인쇄 영역·눈금선  견적서 범위만, 눈금선은 끄기                      (4-02)
  ⑦ 인쇄 제목         둘째 장에도 머리글이 찍히게                        (4-12)
  ⑧ 최종확인 채우기   [최종확인] 다섯 칸                                —

■ 순서를 지켜야 하는 이유

  · ①을 나중에 하면 ⑥에서 잡은 **인쇄 영역이 어긋납니다.**
    빈 행을 지우면 표 크기가 바뀌기 때문입니다.
  · ④를 건너뛰면 ⑤의 곱셈에서 **세 줄이 샙니다.**
  · ⑤에 IFERROR 를 **처음부터** 씌우면 제품표에 없는 코드가
    **안 보입니다.** 다 만들고 맨 마지막에 씌우세요.

■ 채점

  노란 칸 = 채울 곳 · 주황 칸 = 답을 적는 곳 · 흰 칸은 자료입니다."""

C4_STEPS = [
    (u"①", u"빈 서식 지우기", u"4-07", u"Ctrl+End 가 자료 끝이 아닌 곳으로 갑니다. 그 아래 행 전체를 잡아 삭제하고 저장한 뒤 다시 엽니다."),
    (u"②", u"유효성 검사", u"4-08", u"제품코드는 목록(원본 =$K$6:$K$10), 수량은 정수 >= 0."),
    (u"③", u"오류 메시지", u"4-08", u"「오류 메시지」 탭에서 제목과 내용을 직접 씁니다. 기본 문구로는 뭘 하라는지 모릅니다."),
    (u"④", u"글자 숫자", u"1-11", u"수량 열을 잡고 느낌표 → 「숫자로 변환」. 왼쪽에 붙어 있던 것이 오른쪽으로 갑니다."),
    (u"⑤", u"단가·금액", u"7-04", u"단가 = INDEX($M$6:$M$10,MATCH(C6,$K$6:$K$10,0)), 금액 = 수량*단가."),
    (u"⑥", u"인쇄 영역", u"4-02", u"표 범위만 잡아 인쇄 영역 설정. 페이지 레이아웃 → 눈금선 → 인쇄 체크 해제."),
    (u"⑦", u"인쇄 제목", u"4-12", u"페이지 레이아웃 → 인쇄 제목 → 반복할 행에 $5:$5."),
    (u"⑧", u"최종확인", u"—", u"[최종확인] 주황 칸 다섯 개를 채웁니다."),
]

C4_Q = [
    (u"① 수량 합계는?", C4_QTY, u"0",
     u"④를 건너뛰면 글자 세 줄이 빠져 %d 만 나옵니다."
     % sum(q for _, q, t in C4_ROWS if not t)),
    (u"② 금액 합계는?", C4_AMT, u"#,##0",
     u"제품표에 없는 코드 두 줄은 0 입니다. 그 둘까지 억지로 채우면 값이 커집니다."),
    (u"③ 제품표에 없는 코드가 든 줄은?", C4_MISS, u"0",
     u"⑤에 IFERROR 를 먼저 씌우면 #N/A 가 가려져 이 줄을 못 찾습니다."),
    (u"④ 처음에 Ctrl+End 가 간 행 번호는?", C4_JUNK_ROW, u"0",
     u"자료는 %d행까지인데 훨씬 아래로 갑니다. 그 아래가 전부 파일에 저장되고 있었습니다."
     % (FIRST + len(C4_ROWS) - 1)),
    (u"⑤ 인쇄 제목 「반복할 행」에 넣을 값은?", u"$5:$5", u"General",
     u"머리글이 5행입니다. 제목까지 넣으려면 $1:$5 도 됩니다."),
]


def build_ch4():
    wb = Workbook(); wb.remove(wb.active)
    sheet_guide(wb, u"4장 종합문제 — 대림기업 3월 견적서",
                u"내일 아침에 나가야 하는 견적서 한 장을 쓸 수 있게 만듭니다.",
                C4_GUIDE, height=470)
    ws = wb.create_sheet(u"문제")
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 3, "B": 6, "C": 14, "D": 16, "E": 10, "F": 12, "G": 14,
                   "H": 3, "J": 3, "K": 12, "L": 14, "M": 10}.items():
        ws.column_dimensions[col].width = w
    t = ws.cell(row=1, column=2, value=u"대림기업 제품 발주서")
    t.font = Font(name=FONT, size=14, bold=True)
    head(ws, 5, 2, [u"No", u"제품코드", u"제품명", u"수량", u"단가", u"금액"])
    for i, (code, qty, as_text) in enumerate(C4_ROWS):
        r = FIRST + i
        put(ws, r, 2, i + 1)
        put(ws, r, 3, code)
        put(ws, r, 4, None, fill=YELLOW)
        c = put(ws, r, 5, (u"%d" % qty) if as_text else qty)
        if as_text:
            c.alignment = Alignment(horizontal="left", vertical="center")
        put(ws, r, 6, None, fill=YELLOW, fmt=u"#,##0")
        put(ws, r, 7, None, fill=YELLOW, fmt=u"#,##0")
    last = FIRST + len(C4_ROWS) - 1
    put(ws, last + 2, 3, u"합계", bold=True)
    put(ws, last + 2, 5, None, fill=YELLOW, bold=True)
    put(ws, last + 2, 7, None, fill=YELLOW, bold=True, fmt=u"#,##0")

    head(ws, 5, 11, [u"제품코드", u"제품명", u"단가"])
    for i, (code, nm, price) in enumerate(C4_PRICE):
        r = FIRST + i
        put(ws, r, 11, code); put(ws, r, 12, nm); put(ws, r, 13, price, fmt=u"#,##0")
    g = ws.cell(row=3, column=3, value=u"노란 칸을 채우세요. [안내] 의 여덟 단계를 순서대로.")
    g.font = Font(name=FONT, size=10, color=GREY)

    # 일부러 남겨 둔 빈 서식 — Ctrl+End 가 여기로 갑니다
    junk = ws.cell(row=C4_JUNK_ROW, column=20)
    junk.border = THIN
    junk.fill = PatternFill("solid", fgColor=NOTE)

    sheet_check(wb, C4_Q)
    sheet_answer(wb, C4_STEPS, C4_Q)
    return wb, None


# ══════════ 5장 ══════════════════════════════════════════════════
# (지역, 대분류, 담당자, 수량, 금액) — 사이사이에 소계 줄이 섞여 있습니다
C5_DATA = [
    (u"서울", u"사무용품", u"김희윤", 120, 1_240_000),
    (u"서울", u"가구", u"박예설", 95, 2_880_000),
    (u"서울", u"전자제품", u"이강은", 40, 1_060_000),
    (u"경기", u"사무용품", u"박수혁", 150, 3_050_000),
    (u"경기", u"가구", u"정지산", 80, 1_610_000),
    (u"경기", u"전자제품", u"최민슬", 60, 2_270_000),
    (u"경상", u"사무용품", u"이새은", 110, 1_930_000),
    (u"경상", u"가구", u"정다해", 75, 1_450_000),
    (u"전라", u"사무용품", u"최서훈", 90, 990_000),
    (u"전라", u"가구", u"김우진", 130, 3_390_000),
    (u"전라", u"전자제품", u"이해영", 55, 1_710_000),
]
C5_REAL = len(C5_DATA)
C5_TOTAL = sum(d[4] for d in C5_DATA)
C5_SEOUL = sum(d[4] for d in C5_DATA if d[0] == u"서울")
C5_ADV = sum(1 for d in C5_DATA
             if (d[1] == u"사무용품" and d[3] >= 100) or (d[1] == u"가구" and d[3] >= 80))
C5_TOP3 = sum(sorted((d[4] for d in C5_DATA), reverse=True)[:3])
C5_ORDER = [u"서울", u"경기", u"경상", u"전라"]

C5_GUIDE = u"""■ 상황

  세아물산 상반기 주문내역입니다. 팀장이 메일을 보냈습니다 —

    「사무용품은 100개 이상, 가구는 80개 이상 나간 건만 뽑아서
     지역은 서울·경기·경상·전라 순으로 세워 주세요.
     그리고 서울만 따로 합계도 알려 주시고요.」

  표에는 **지역마다 소계 줄이 섞여** 있습니다.

■ 일곱 단계 — 순서대로

  ① 소계 걷어내기   「소계」 로 필터해서 그 줄만 지웁니다            (5-06)
  ② 사용자 지정 목록 서울·경기·경상·전라 를 등록합니다               (5-04)
  ③ 정렬            지역(사용자 지정) → 대분류 → 금액 큰 순          (5-03)
  ④ 합계 고치기     J4 의 =SUM 을 =SUBTOTAL(9,…) 로                (5-08)
  ⑤ 필터로 확인     서울만 걸어 J4 가 따라 바뀌는지 봅니다            (5-06)
  ⑥ 고급 필터       조건표를 만들어 [추출] 시트로 뽑습니다            (5-10)
  ⑦ 최종확인 채우기 [최종확인] 다섯 칸                              —

■ 순서를 지켜야 하는 이유

  · ①을 안 하면 **합계가 정확히 두 배**가 됩니다. 오류가 안 뜹니다.
  · ①을 안 하고 정렬하면 소계 줄이 **엉뚱한 자리로 흩어집니다.**
  · ④를 안 하면 ⑤에서 **필터를 걸어도 합계가 안 바뀝니다.**
  · ⑥의 조건표 머리글은 **원본에서 복사**하세요. 한 글자만 달라도 안 잡힙니다.

■ 채점

  노란 칸 = 채울 곳 · 주황 칸 = 답을 적는 곳 · 흰 칸은 자료입니다."""

C5_STEPS = [
    (u"①", u"소계 걷어내기", u"5-06", u"담당자 열에서 「소계」 만 필터로 남기고, 보이는 줄을 전부 지웁니다. 필터 상태의 행 삭제는 보이는 줄만 지웁니다."),
    (u"②", u"사용자 지정 목록", u"5-04", u"파일 → 옵션 → 고급 → 맨 아래 「사용자 지정 목록 편집」 에 네 줄을 넣습니다."),
    (u"③", u"정렬", u"5-03", u"표 안 한 칸만 누르고 데이터 → 정렬. 기준 셋을 위에서부터 지역·대분류·금액."),
    (u"④", u"SUBTOTAL", u"5-08", u"=SUBTOTAL(9,F6:F16). 9 는 합계입니다."),
    (u"⑤", u"필터 확인", u"5-06", u"Ctrl+Shift+L 로 필터를 켜고 지역에서 서울만 고릅니다."),
    (u"⑥", u"고급 필터", u"5-10", u"[추출] 시트에서 「데이터 → 고급」 을 누르고, 목록 범위로 [문제] 시트를 잡습니다."),
    (u"⑦", u"최종확인", u"—", u"[최종확인] 주황 칸 다섯 개를 채웁니다."),
]

C5_Q = [
    (u"① 소계를 뺀 진짜 주문은 몇 줄?", C5_REAL, u"0",
     u"소계 줄이 지역마다 하나씩 섞여 있습니다. 담당자 칸을 보면 구별됩니다."),
    (u"② 전체 금액 합계는?", C5_TOTAL, u"#,##0",
     u"①을 건너뛰면 정확히 두 배인 %s 이 나옵니다." % format(C5_TOTAL * 2, ",")),
    (u"③ 서울만 걸었을 때 합계는?", C5_SEOUL, u"#,##0",
     u"④를 안 하면 필터를 걸어도 전체 합계가 그대로 나옵니다."),
    (u"④ 고급 필터 결과는 몇 줄?", C5_ADV, u"0",
     u"사무용품 100↑ 또는 가구 80↑ 입니다. 조건을 같은 줄에 적으면 「그리고」가 되어 0 줄이 됩니다."),
    (u"⑤ 금액 상위 세 건의 합은?", C5_TOP3, u"#,##0",
     u"숫자 필터 → 상위 10 에서 항목을 3 으로 바꾸면 됩니다."),
]


def build_ch5():
    wb = Workbook(); wb.remove(wb.active)
    sheet_guide(wb, u"5장 종합문제 — 세아물산 상반기 주문내역",
                u"소계가 섞인 표를 걷어내고, 세우고, 조건에 맞는 것만 뽑습니다.",
                C5_GUIDE, height=470)
    ws = wb.create_sheet(u"문제")
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 3, "B": 10, "C": 14, "D": 12, "E": 10, "F": 15,
                   "G": 3, "H": 14, "I": 3, "J": 16}.items():
        ws.column_dimensions[col].width = w
    t = ws.cell(row=1, column=2, value=u"세아물산 2026 상반기 주문내역")
    t.font = Font(name=FONT, size=14, bold=True)
    put(ws, 3, 9, u"합계", bold=True)
    put(ws, 3, 10, None, fill=YELLOW, fmt=u"#,##0")

    head(ws, 5, 2, [u"지역", u"대분류", u"담당자", u"수량", u"금액"])
    r = FIRST
    prev = None
    for region, cat, who, qty, amt in C5_DATA:
        if prev is not None and region != prev:
            sub = [d for d in C5_DATA if d[0] == prev]
            put(ws, r, 2, prev); put(ws, r, 3, u"소계"); put(ws, r, 4, u"소계")
            put(ws, r, 5, sum(d[3] for d in sub))
            put(ws, r, 6, sum(d[4] for d in sub), fmt=u"#,##0")
            r += 1
        put(ws, r, 2, region); put(ws, r, 3, cat); put(ws, r, 4, who)
        put(ws, r, 5, qty); put(ws, r, 6, amt, fmt=u"#,##0")
        prev = region
        r += 1
    sub = [d for d in C5_DATA if d[0] == prev]
    put(ws, r, 2, prev); put(ws, r, 3, u"소계"); put(ws, r, 4, u"소계")
    put(ws, r, 5, sum(d[3] for d in sub))
    put(ws, r, 6, sum(d[4] for d in sub), fmt=u"#,##0")

    g = ws.cell(row=3, column=2, value=u"노란 칸을 채우세요. [안내] 의 일곱 단계를 순서대로.")
    g.font = Font(name=FONT, size=10, color=GREY)

    ex = wb.create_sheet(u"추출")
    ex.sheet_view.showGridLines = False
    for col, w in {"A": 3, "B": 10, "C": 14, "D": 12, "E": 10, "F": 15}.items():
        ex.column_dimensions[col].width = w
    t = ex.cell(row=1, column=2, value=u"추출 — 조건표를 만들고 결과를 여기로")
    t.font = Font(name=FONT, size=14, bold=True)
    s = ex.cell(row=2, column=2,
                value=u"3행에 조건 머리글을, 4·5행에 조건을 적으세요. 머리글은 [문제] 에서 복사하세요.")
    s.font = Font(name=FONT, size=10, color=GREY)
    for rr in (3, 4, 5):
        for cc in range(2, 7):
            put(ex, rr, cc, None, fill=YELLOW)
    n = ex.cell(row=7, column=2, value=u"↓ 결과는 여기부터 나옵니다")
    n.font = Font(name=FONT, size=10, color=GREY)

    sheet_check(wb, C5_Q)
    sheet_answer(wb, C5_STEPS, C5_Q)
    return wb, None


# ══════════ 6장 ══════════════════════════════════════════════════
C6_YEARS = [2021, 2022, 2023, 2024, 2025]
# 지점 → 연도별 (매출, 이익)
C6_ROWS = [
    (u"강남점", [(30000, 1488), (8000, 1357), (26000, 1575), (19000, 1862), (24000, 1104)]),
    (u"명동점", [(37000, 2245), (6000, 1531), (25000, 1250), (22000, 1364), (31000, 1798)]),
    (u"서초점", [(33000, 1650), (11000, 1023), (33000, 1749), (28000, 1596), (26000, 1352)]),
    (u"일산점", [(17000, 1262), (17000, 1173), (15000, 1005), (20000, 1240), (18000, 1044)]),
    (u"용산점", [(39000, 2028), (9000, 1305), (11000, 913), (24000, 1512), (29000, 1682)]),
]
C6_LINES = len(C6_ROWS) * len(C6_YEARS)
C6_Y2023 = sum(v[2][0] for _, v in C6_ROWS)
C6_TOP = max(C6_ROWS, key=lambda x: sum(v[0] for v in x[1]))[0]
C6_SALES = sum(v[0] for _, vs in C6_ROWS for v in vs)
C6_PROFIT = sum(v[1] for _, vs in C6_ROWS for v in vs)
C6_FIELD = C6_PROFIT / float(C6_SALES)
C6_ITEM = sum(v[1] / float(v[0]) for _, vs in C6_ROWS for v in vs)

C6_GUIDE = u"""■ 상황

  세나상사 지점별 매출표입니다. **지점이 세로, 연도가 가로**로 되어 있습니다.
  이대로는 **연도별로 나눠 볼 수가 없습니다.**

  이것을 쓸 수 있는 자료로 눕히고, 피벗을 만들고, 이익률까지 냅니다.

■ 일곱 단계 — 순서대로

  ① 눕히기        [자료] 시트에 지점·연도·매출·이익 네 열로 옮깁니다   (6-03)
  ② 표로 바꾸기   Ctrl+T, 표 이름을 「지점매출」 로                    (6-01)
  ③ 피벗 만들기   행=지점, 열=연도, 값=매출 합계                      (6-04)
  ④ 계산 필드     이익률 = 이익 / 매출                               (6-11)
  ⑤ 계산 항목     같은 이익률을 계산 항목으로도 만들어 봅니다           (6-11)
  ⑥ 새로 고침     [자료] 에 한 줄 더하고 Alt+F5                       (6-06)
  ⑦ 최종확인 채우기 [최종확인] 다섯 칸                                —

■ 순서를 지켜야 하는 이유

  · ①을 안 하면 필드 목록에 **「2021」 「2022」 가 따로** 뜹니다.
    **「연도」 라는 필드가 없어** 연도별로 못 나눕니다.
  · ④와 ⑤는 **같은 이익률인데 총합계가 다릅니다.** 그것이 이 문제의 핵심입니다.
  · ⑥에서 새로 고침을 안 하면 **옛날 숫자가 그대로** 인쇄됩니다.

■ 채점

  노란 칸 = 채울 곳 · 주황 칸 = 답을 적는 곳 · 흰 칸은 자료입니다."""

C6_STEPS = [
    (u"①", u"눕히기", u"6-03", u"한 줄 = 한 사건. 「강남점의 2021년 매출은 30000, 이익은 1488」 이 한 줄입니다. 파워 쿼리 → 「다른 열 피벗 해제」 로도 됩니다."),
    (u"②", u"표로 바꾸기", u"6-01", u"Ctrl+T. 표 디자인 탭 왼쪽 끝에서 이름을 「지점매출」 로 바꿉니다."),
    (u"③", u"피벗", u"6-04", u"삽입 → 피벗 테이블 → 새 워크시트. 행=지점, 열=연도, 값=매출."),
    (u"④", u"계산 필드", u"6-11", u"피벗 분석 → 필드, 항목 및 집합 → 계산 필드. 이름 「이익률」, 수식 = 이익 / 매출."),
    (u"⑤", u"계산 항목", u"6-11", u"같은 곳의 「계산 항목」 으로도 만들어 총합계를 견줍니다. 값이 다릅니다."),
    (u"⑥", u"새로 고침", u"6-06", u"[자료] 에 한 줄 더하고 피벗 안에서 Alt+F5. 표로 바꿔 두었으니 범위가 따라옵니다."),
    (u"⑦", u"최종확인", u"—", u"[최종확인] 주황 칸 다섯 개를 채웁니다."),
]

C6_Q = [
    (u"① 눕히면 몇 줄이 되나?", C6_LINES, u"0",
     u"%d지점 × %d년 입니다. 머리글은 안 셉니다." % (len(C6_ROWS), len(C6_YEARS))),
    (u"② 2023년 전체 매출은?", C6_Y2023, u"#,##0",
     u"①을 안 하면 연도별로 나눌 수가 없어 이 값을 낼 수 없습니다."),
    (u"③ 5년 매출이 가장 큰 지점은?", C6_TOP, u"General",
     u"피벗의 총합계 열을 견주면 됩니다."),
    (u"④ 계산 필드로 낸 전체 이익률은? (%, 소수 둘째까지)",
     round(C6_FIELD * 100, 2), u"0.00",
     u"이익 합계 %s ÷ 매출 합계 %s. 합계를 먼저 내고 나눕니다 — 이쪽이 맞습니다."
     % (format(C6_PROFIT, ","), format(C6_SALES, ","))),
    (u"⑤ 계산 항목으로 내면 총합계가 얼마? (%, 소수 둘째까지)",
     round(C6_ITEM * 100, 2), u"0.00",
     u"줄마다 낸 이익률 %d개를 그냥 더한 값입니다. 아무 뜻이 없는데 피벗은 자신 있게 내놓습니다."
     % C6_LINES),
]


def build_ch6():
    wb = Workbook(); wb.remove(wb.active)
    sheet_guide(wb, u"6장 종합문제 — 세나상사 지점별 매출",
                u"가로로 퍼진 표를 눕히고, 피벗을 만들고, 이익률의 함정을 확인합니다.",
                C6_GUIDE, height=450)

    ws = wb.create_sheet(u"원본")
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 3, "B": 12}.items():
        ws.column_dimensions[col].width = w
    for i in range(len(C6_YEARS) * 2):
        ws.column_dimensions[chr(ord("C") + i)].width = 12
    t = ws.cell(row=1, column=2, value=u"세나상사 지점별 매출 / 이익 (단위: 만원)")
    t.font = Font(name=FONT, size=14, bold=True)
    for i, y in enumerate(C6_YEARS):
        c = ws.cell(row=4, column=3 + i * 2, value=u"%d년" % y)
        c.font = Font(name=FONT, size=10, bold=True)
        c.alignment = Alignment(horizontal="center")
        ws.merge_cells(start_row=4, start_column=3 + i * 2,
                       end_row=4, end_column=4 + i * 2)
    head(ws, 5, 2, [u"지점"] + [u"매출" if j % 2 == 0 else u"이익"
                                for j in range(len(C6_YEARS) * 2)])
    for i, (nm, vs) in enumerate(C6_ROWS):
        r = FIRST + i
        put(ws, r, 2, nm)
        for j, (s, p) in enumerate(vs):
            put(ws, r, 3 + j * 2, s, fmt=u"#,##0")
            put(ws, r, 4 + j * 2, p, fmt=u"#,##0")
    g = ws.cell(row=3, column=2, value=u"이 표는 손대지 마세요. [자료] 시트에 옮겨 적는 것이 ① 입니다.")
    g.font = Font(name=FONT, size=10, color=GREY)

    dt = wb.create_sheet(u"자료")
    dt.sheet_view.showGridLines = False
    for col, w in {"A": 3, "B": 12, "C": 10, "D": 12, "E": 12}.items():
        dt.column_dimensions[col].width = w
    t = dt.cell(row=1, column=2, value=u"자료 — 여기에 세로로 눕혀 적으세요")
    t.font = Font(name=FONT, size=14, bold=True)
    head(dt, 3, 2, [u"지점", u"연도", u"매출", u"이익"])
    for i in range(C6_LINES):
        r = 4 + i
        for cc in range(2, 6):
            put(dt, r, cc, None, fill=YELLOW)

    sheet_check(wb, C6_Q)
    sheet_answer(wb, C6_STEPS, C6_Q)
    return wb, None


# ══════════ 7장 ══════════════════════════════════════════════════
C7_PRICE = [(u"FQ-003", 8500), (u"JD-005", 12400), (u"KM-011", 6300),
            (u"PL-027", 21000), (u"SY-040", 4800)]
C7_MAP = dict(C7_PRICE)
# (거래처명, 앞공백여부, 날짜코드 YYMMDD, 제품코드, 수량)
C7_ROWS = [
    (u"한결상사", True, u"260115", u"FQ-003", 40),
    (u"미래산업", False, u"260228", u"JD-005", 25),
    (u"대성물산", True, u"260310", u"KM-011", 60),
    (u"신영테크", False, u"260402", u"PL-027", 15),
    (u"한결상사", True, u"260519", u"SY-040", 80),
    (u"미래산업", False, u"260620", u"FQ-003", 35),
    (u"대성물산", True, u"260711", u"JD-005", 20),
    (u"신영테크", False, u"260805", u"KM-011", 50),
]
C7_PAD = u"　　　"          # 전각 공백 세 개
C7_BASE = (2026, 9, 16)
C7_CLEAN = sum(1 for _, pad, _, _, _ in C7_ROWS if not pad)
C7_AMT = sum(q * C7_MAP[c] for _, _, _, c, q in C7_ROWS)


def _months(code):
    y, m, d = 2000 + int(code[:2]), int(code[2:4]), int(code[4:])
    n = (C7_BASE[0] - y) * 12 + (C7_BASE[1] - m)
    if C7_BASE[2] < d:
        n -= 1
    return n


C7_OLD = sum(1 for _, _, code, _, _ in C7_ROWS if _months(code) >= 6)
_by = {}
for nm, _, _, code, qty in C7_ROWS:
    _by[nm] = _by.get(nm, 0) + qty * C7_MAP[code]
C7_BEST = max(_by.items(), key=lambda kv: kv[1])[0]
C7_LEN = len(C7_PAD + u"한결상사")

C7_GUIDE = u"""■ 상황

  한결상사 거래처 정산표입니다. 앞으로 갈수록 **하나씩 막힙니다.**

  · 거래처명 앞에 **보이지 않는 공백**이 붙어 있습니다
  · 거래일이 **글자**(YYMMDD)로 들어 있습니다
  · 단가표는 **찾을 열이 오른쪽**에 있습니다

■ 여덟 단계 — 순서대로

  ① 글자 수 세기   LEN 으로 거래처명이 몇 글자인지 봅니다        (7-07)
  ② 공백 떼기      TRIM 으로 정리한 열을 따로 만듭니다            (7-07)
  ③ 날짜 살리기    DATE + LEFT/MID/RIGHT 로 진짜 날짜로          (7-11)
  ④ 단가 가져오기  INDEX + MATCH (VLOOKUP 은 왼쪽을 못 봅니다)    (7-20)
  ⑤ 금액          수량 × 단가                                  (7-01)
  ⑥ 거래처별 합계  SUMIFS 로 — 반드시 TRIM 한 열을 봅니다         (7-05)
  ⑦ 경과 개월      DATEDIF 로 기준일까지 몇 개월인지              (7-12)
  ⑧ 최종확인 채우기 [최종확인] 다섯 칸                            —

■ 순서를 지켜야 하는 이유

  · ②를 건너뛰면 ⑥의 SUMIFS 가 **공백 붙은 거래처를 못 찾습니다.**
    합계가 조용히 작아집니다.
  · ③을 건너뛰면 ⑦의 DATEDIF 가 **#VALUE!** 를 냅니다.
  · **IFERROR 는 맨 마지막에** 씌우세요. 먼저 씌우면
    안 찾아진 줄이 **빈칸으로 보여** 문제를 못 찾습니다.

■ 채점

  노란 칸 = 채울 곳 · 주황 칸 = 답을 적는 곳 · 흰 칸은 자료입니다."""

C7_STEPS = [
    (u"①", u"글자 수", u"7-07", u"=LEN(B6). 「한결상사」는 4여야 하는데 7 이 나옵니다. 앞에 공백이 셋 붙어 있습니다."),
    (u"②", u"공백 떼기", u"7-07", u"=TRIM(B6). 웹에서 온 자료라면 =TRIM(SUBSTITUTE(B6,CHAR(160),\" \")) 가 안전합니다."),
    (u"③", u"날짜", u"7-11", u"=DATE(LEFT(C6,2)+2000,MID(C6,3,2),RIGHT(C6,2)). +2000 을 빼먹으면 서기 26년이 됩니다."),
    (u"④", u"단가", u"7-20", u"=INDEX($L$6:$L$10,MATCH(D6,$K$6:$K$10,0)). 찾을 열이 오른쪽이라 VLOOKUP 은 안 됩니다."),
    (u"⑤", u"금액", u"7-01", u"=E6*F6."),
    (u"⑥", u"거래처별 합계", u"7-05", u"=SUMIFS($G$6:$G$13,$H$6:$H$13,J6) — 원본 B열이 아니라 TRIM 한 H열을 봐야 합니다."),
    (u"⑦", u"경과 개월", u"7-12", u"=DATEDIF(I6,$F$3,\"M\"). 자동 완성에 안 뜨니 손으로 전부 치세요."),
    (u"⑧", u"최종확인", u"—", u"[최종확인] 주황 칸 다섯 개를 채웁니다."),
]

C7_Q = [
    (u"① 공백이 안 붙은 거래처 줄은 몇 줄?", C7_CLEAN, u"0",
     u"②를 하기 전에 SUMIFS 를 걸면 이 %d줄만 잡힙니다. 나머지 %d줄이 조용히 빠집니다."
     % (C7_CLEAN, len(C7_ROWS) - C7_CLEAN)),
    (u"② 금액 합계는?", C7_AMT, u"#,##0",
     u"④·⑤를 다 해야 나옵니다. 단가를 못 찾은 줄이 있으면 작아집니다."),
    (u"③ 가장 많이 산 거래처는?", C7_BEST, u"General",
     u"⑥의 SUMIFS 결과를 견주면 됩니다. TRIM 을 안 하면 답이 달라집니다."),
    (u"④ 기준일까지 6개월 이상 지난 건은?", C7_OLD, u"0",
     u"③·⑦을 다 해야 셉니다. 기준일은 F3 에 적혀 있습니다."),
    (u"⑤ 공백이 붙은 거래처명의 LEN 값은?", C7_LEN, u"0",
     u"「한결상사」 4글자 + 앞 공백 3개 = %d. 눈으로는 절대 안 보입니다." % C7_LEN),
]


def build_ch7():
    wb = Workbook(); wb.remove(wb.active)
    sheet_guide(wb, u"7장 종합문제 — 한결상사 거래처 정산표",
                u"보이지 않는 공백, 글자로 된 날짜, 왼쪽에 있는 단가 — 셋을 차례로 풉니다.",
                C7_GUIDE, height=470)
    ws = wb.create_sheet(u"문제")
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 3, "B": 18, "C": 11, "D": 11, "E": 8, "F": 11, "G": 13,
                   "H": 14, "I": 13, "J": 3, "K": 11, "L": 10, "M": 3,
                   "N": 14, "O": 14}.items():
        ws.column_dimensions[col].width = w
    t = ws.cell(row=1, column=2, value=u"한결상사 거래처 정산표")
    t.font = Font(name=FONT, size=14, bold=True)
    put(ws, 3, 5, u"기준일", bold=True)
    put(ws, 3, 6, u"%04d-%02d-%02d" % C7_BASE)

    head(ws, 5, 2, [u"거래처명", u"날짜코드", u"제품코드", u"수량", u"단가", u"금액",
                    u"TRIM", u"진짜 날짜"])
    for i, (nm, pad, code, prod, qty) in enumerate(C7_ROWS):
        r = FIRST + i
        c = put(ws, r, 2, (C7_PAD + nm) if pad else nm)
        c.alignment = Alignment(horizontal="left", vertical="center")
        put(ws, r, 3, code)
        put(ws, r, 4, prod)
        put(ws, r, 5, qty)
        put(ws, r, 6, None, fill=YELLOW, fmt=u"#,##0")
        put(ws, r, 7, None, fill=YELLOW, fmt=u"#,##0")
        put(ws, r, 8, None, fill=YELLOW)
        put(ws, r, 9, None, fill=YELLOW, fmt=u"yyyy-mm-dd")

    head(ws, 5, 11, [u"제품코드", u"단가"])
    for i, (code, price) in enumerate(C7_PRICE):
        r = FIRST + i
        put(ws, r, 11, code); put(ws, r, 12, price, fmt=u"#,##0")

    head(ws, 5, 14, [u"거래처", u"합계"])
    for i, nm in enumerate(sorted(_by)):
        r = FIRST + i
        put(ws, r, 14, nm)
        put(ws, r, 15, None, fill=YELLOW, fmt=u"#,##0")

    g = ws.cell(row=3, column=2, value=u"노란 칸을 채우세요. [안내] 의 여덟 단계를 순서대로.")
    g.font = Font(name=FONT, size=10, color=GREY)

    sheet_check(wb, C7_Q)
    sheet_answer(wb, C7_STEPS, C7_Q)
    return wb, None


# ══════════ 8장 ══════════════════════════════════════════════════
C8_BAR = [(u"서울특별시", 698), (u"경기도", 301), (u"부산광역시", 142),
          (u"인천광역시", 118), (u"대구광역시", 96)]
C8_PIE = [(u"A제품", 34), (u"B제품", 22), (u"C제품", 15), (u"D제품", 9),
          (u"E제품", 7), (u"F제품", 5), (u"G제품", 4), (u"H제품", 3), (u"I제품", 1)]
C8_LINE = [(u"1분기", 102), (u"2분기", 112), (u"3분기", 118), (u"4분기(예상)", 124)]
C8_MIX = [(2023, 2_236_330, 0.0136), (2024, 2_305_271, 0.0208),
          (2025, 2_401_800, 0.0419), (2026, 2_468_500, 0.0278)]

C8_PIE_TOTAL = sum(v for _, v in C8_PIE)
_acc, C8_PIE_N = 0, 0
for _, v in sorted(C8_PIE, key=lambda x: -x[1]):
    _acc += v
    C8_PIE_N += 1
    if _acc / float(C8_PIE_TOTAL) >= 0.8:
        break
C8_RATIO = round(C8_BAR[0][1] / float(C8_BAR[-1][1]), 1)

C8_GUIDE = u"""■ 상황

  보고서에 들어갈 차트 **네 개**를 받았습니다.
  **넷 다 어딘가 거짓말을 하고 있습니다.**

  [문제] 시트에 각 차트의 자료와 **지금 어떻게 그려져 있는지**가 적혀 있습니다.
  직접 차트를 그려 보고, **무엇이 잘못됐는지 찾아 고치세요.**

■ 여섯 단계 — 순서대로

  ① 1번 — 막대   세로 축 최소값을 확인합니다                     (8-02)
  ② 2번 — 가로막대 항목 순서를 확인합니다                        (8-02)
  ③ 3번 — 꺾은선 예상치가 어떻게 그려져 있는지 봅니다             (8-02)
  ④ 4번 — 원형   조각이 몇 개인지 셉니다                        (8-01)
  ⑤ 색 점검      흑백으로 인쇄해도 구별되는지 봅니다              (8-06)
  ⑥ 최종확인 채우기 [최종확인] 다섯 칸                            —

■ 이 넷은 일부러 하지 않아도 저절로 그렇게 됩니다

  · 엑셀이 축을 「보기 좋게」 자동으로 잡아 줍니다 — 0 에서 시작 안 합니다
  · 가로 막대는 기본이 거꾸로입니다
  · 예상치라고 엑셀에게 알려 줄 방법이 없습니다
  · 원형에 조각 아홉 개를 넣어도 엑셀은 말리지 않습니다

  그래서 **만든 뒤에 한 번 확인하는 습관**이 필요합니다.

■ 채점

  주황 칸 = 답을 적는 곳 · 흰 칸은 자료입니다.
  ④는 숫자로, 나머지는 짧은 말로 적으세요."""

C8_STEPS = [
    (u"①", u"막대 축", u"8-02", u"막대는 길이가 값입니다. 축 서식 → 최소값을 0 으로. 꺾은선은 0 부터 안 해도 됩니다."),
    (u"②", u"가로 막대", u"8-02", u"세로 축 서식 → 「항목을 거꾸로」 체크. 표의 1등이 차트에서도 맨 위로 옵니다."),
    (u"③", u"예상치", u"8-02", u"4분기 구간만 골라 서식 → 선 → 대시 종류를 점선으로. 안 하면 실제 숫자로 읽힙니다."),
    (u"④", u"원형", u"8-01", u"조각이 다섯을 넘으면 가로 막대로 바꾸세요. 남기더라도 작은 것은 「기타」로 묶습니다."),
    (u"⑤", u"색", u"8-06", u"빨강·초록을 나란히 쓰지 않습니다. 흑백 인쇄를 생각해 선 모양·표식 모양도 다르게."),
    (u"⑥", u"최종확인", u"—", u"[최종확인] 주황 칸 다섯 개를 채웁니다."),
]

C8_Q = [
    (u"① 1번 막대의 세로 축 최소값을 얼마로?", 0, u"0",
     u"막대는 길이가 값입니다. 600 부터 시작하면 698 과 301 의 차이가 실제보다 커 보입니다."),
    (u"② 1번에서 1등과 5등의 실제 배수는? (소수 첫째까지)", C8_RATIO, u"0.0",
     u"%d ÷ %d 입니다. 축을 잘못 잡으면 이보다 훨씬 커 보입니다."
     % (C8_BAR[0][1], C8_BAR[-1][1])),
    (u"③ 2번 가로 막대에서 켜야 할 축 옵션은?", u"항목을 거꾸로", u"General",
     u"안 켜면 표의 1등이 차트에서는 맨 아래에 옵니다. 거의 항상 켜야 합니다."),
    (u"④ 4번 원형에서 상위 몇 조각이 80% 를 넘나?", C8_PIE_N, u"0",
     u"나머지 %d 조각은 합쳐도 20%% 가 안 됩니다. 「기타」로 묶는 편이 낫습니다."
     % (len(C8_PIE) - C8_PIE_N)),
    (u"⑤ 3번 꺾은선의 4분기를 어떻게 그려야 하나?", u"점선", u"General",
     u"실선으로 두면 아직 안 일어난 일을 이미 일어난 것처럼 보이게 합니다."),
]


def build_ch8():
    wb = Workbook(); wb.remove(wb.active)
    sheet_guide(wb, u"8장 종합문제 — 거짓말하는 차트 네 개",
                u"이미 그려진 차트 넷에서 잘못된 곳을 찾아 고칩니다.",
                C8_GUIDE, height=430)
    ws = wb.create_sheet(u"문제")
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 3, "B": 16, "C": 14, "D": 14, "E": 3, "F": 62}.items():
        ws.column_dimensions[col].width = w

    def block(row, no, title_text, cols, rows, symptom, fmts=None):
        t = ws.cell(row=row, column=2, value=u"%d번 — %s" % (no, title_text))
        t.font = Font(name=FONT, size=12, bold=True)
        head(ws, row + 1, 2, cols)
        for i, vals in enumerate(rows):
            r = row + 2 + i
            for j, v in enumerate(vals):
                put(ws, r, 2 + j, v, fmt=(fmts or {}).get(j))
        s = ws.cell(row=row + 1, column=6, value=u"지금 이렇게 그려져 있습니다 —\n" + symptom)
        s.font = Font(name=FONT, size=10, color=GREY)
        s.alignment = Alignment(vertical="top", wrap_text=True)
        return row + 2 + len(rows) + 1

    r = 3
    r = block(r, 1, u"지역별 합격자 수 (세로 막대)", [u"시도", u"인원"],
              [list(x) for x in C8_BAR],
              u"세로 축이 600 부터 시작합니다.\n서울 막대가 대구 막대의 열 배쯤 되어 보입니다.")
    r = block(r, 2, u"같은 자료 (가로 막대)", [u"시도", u"인원"],
              [list(x) for x in C8_BAR],
              u"1등 서울이 차트 맨 아래에 있습니다.\n표와 순서가 반대입니다.")
    r = block(r, 3, u"영업1팀 분기 실적 (꺾은선)", [u"분기", u"실적(억)"],
              [list(x) for x in C8_LINE],
              u"네 점이 같은 실선으로 이어져 있습니다.\n4분기는 아직 안 끝났습니다.")
    r = block(r, 4, u"제품별 점유율 (원형)", [u"제품", u"점유율(%)"],
              [list(x) for x in C8_PIE],
              u"조각이 아홉 개입니다. 범례가 오른쪽에 있고\n조각에는 아무 글자도 없습니다.")
    block(r, 5, u"곁들여 — GDP와 성장률 (혼합)", [u"년도", u"GDP(십억)", u"성장률"],
          [list(x) for x in C8_MIX],
          u"둘 다 막대로 그려져 있습니다.\n어느 축이 어느 막대인지 알 수 없습니다.",
          fmts={1: u"#,##0", 2: u"0.0%"})

    g = ws.cell(row=1, column=2,
                value=u"자료는 손대지 마세요. 차트를 그려 보고 무엇이 잘못됐는지 찾는 문제입니다.")
    g.font = Font(name=FONT, size=10, color=GREY)

    sheet_check(wb, C8_Q)
    sheet_answer(wb, C8_STEPS, C8_Q)
    return wb, None


BOOKS = [
    (u"1장_시작부터_남다른_실무자의_엑셀_활용", u"1장_종합문제.xlsx", build_ch1, C1_Q),
    (u"2장_실무자라면_반드시_알아야_할_엑셀_활용", u"2장_종합문제.xlsx", build_ch2, C2_Q),
    (u"3장_보고서가_달라지는_서식_활용법", u"3장_종합문제.xlsx", build_ch3, C3_Q),
    (u"4장_완성한_보고서_공유_및_출력하기", u"4장_종합문제.xlsx", build_ch4, C4_Q),
    (u"5장_데이터_정리부터_데이터_필터링까지", u"5장_종합문제.xlsx", build_ch5, C5_Q),
    (u"6장_자동화를_위한_표와_피벗_테이블", u"6장_종합문제.xlsx", build_ch6, C6_Q),
    (u"7장_기본과_필수_함수_익히기", u"7장_종합문제.xlsx", build_ch7, C7_Q),
    (u"8장_실무에서_필요한_데이터_시각화", u"8장_종합문제.xlsx", build_ch8, C8_Q),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args()
    for folder, name, build, qs in BOOKS:
        wb, _ = build()
        path = os.path.join(u"강의예제", folder, name)
        print(u"%s" % path)
        print(u"   시트 %d장 — %s" % (len(wb.sheetnames), u" · ".join(wb.sheetnames)))
        print(u"   답: %s" % u" / ".join(
            format(q[1], ',') if isinstance(q[1], (int, float)) else u"%s" % q[1]
            for q in qs))
        if a.write:
            wb.save(path)
    print(u"\n만들었습니다." if a.write else
          u"\n※ 보기만 했습니다. 실제로 만들려면 --write 를 붙이세요.")


if __name__ == "__main__":
    main()
