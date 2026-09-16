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


BOOKS = [
    (u"1장_시작부터_남다른_실무자의_엑셀_활용", u"1장_종합문제.xlsx", build_ch1, C1_Q),
    (u"2장_실무자라면_반드시_알아야_할_엑셀_활용", u"2장_종합문제.xlsx", build_ch2, C2_Q),
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
        print(u"   답: %s" % u" / ".join(format(q[1], ',') for q in qs))
        if a.write:
            wb.save(path)
    print(u"\n만들었습니다." if a.write else
          u"\n※ 보기만 했습니다. 실제로 만들려면 --write 를 붙이세요.")


if __name__ == "__main__":
    main()
