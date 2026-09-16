# -*- coding: utf-8 -*-
u"""
1·2장을 한 문제로 묶은 종합문제를 만듭니다.

  python tools/build-final-exam.py            # 보기만
  python tools/build-final-exam.py --write    # 실제로 만들기

왜 따로 만드나
  장별 복습(`N장_복습.xlsx`)은 **기술 하나씩** 확인하는 것입니다.
  실무는 그렇게 오지 않습니다 — **한 표에 문제가 여럿 섞여** 옵니다.

  그래서 **엉망인 표 한 장**을 주고, 여덟 단계를 **순서대로** 밟아
  쓸 수 있는 표로 만들게 합니다. 앞 단계를 건너뛰면 뒤가 안 됩니다.

여덟 단계 — 어느 강의에서 배운 것인지
  ① 병합 풀기            2-03
  ② 빈 칸 채우기          2-04
  ③ 글자로 된 숫자 고치기   1-11
  ④ 금액 = 수량 × 단가     1-10
  ⑤ 담당자 영문 뽑기       1-07
  ⑥ 제품코드 나누기        2-12
  ⑦ 커미션 — $ 로 비율 고정  1-09
  ⑧ 최종 확인 다섯 칸

순서가 곧 함정입니다
  · ②를 안 하면 ③·④가 **일부 줄만** 맞습니다.
  · ③을 안 하면 ④의 곱셈이 **글자 × 숫자**가 되어 틀립니다.
  · ⑦에서 `$` 를 안 붙이면 **아래로 갈수록 0** 이 됩니다.

채점
  [최종확인] 의 다섯 칸만 맞으면 앞 단계는 다 맞은 것입니다.
  틀리면 **몇 번 답이 틀렸는지로 어느 단계에서 샜는지** 알 수 있습니다.
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

OUT = os.path.join(u"강의예제", u"종합문제.xlsx")

# (분류, 지점, 수량, 단가, 담당자, 제품코드) — 분류가 None 인 줄은 병합으로 가려집니다
ROWS = [
    (u"선풍기", u"신촌점", 142, 38000,   u"김세민 (SM.Kim)",   u"FN-2026-001"),
    (None,     u"대방점", 111, 38000,   u"정다온 (DO.Jeong)", u"FN-2026-002"),
    (None,     u"구로점", 151, 38000,   u"김진선 (JS.Kim)",   u"FN-2026-003"),
    (u"TV",    u"신촌점", 163, 850000,  u"정희엘 (HL.Jeong)", u"TV-2026-001"),
    (None,     u"구로점", 125, 850000,  u"박단비 (DB.Park)",  u"TV-2026-002"),
    (u"에어컨", u"대방점", 133, 1200000, u"정진하 (JH.Jeong)", u"AC-2026-001"),
    (None,  u"영등포점", 135, 1200000, u"김병민 (BM.Kim)",   u"AC-2026-002"),
]
AS_TEXT = (1, 3, 5)        # 수량을 일부러 글자로 넣을 줄 (0부터)
RATE = 0.025
FIRST = 6                  # 자료 첫 줄


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


GUIDE = u"""■ 상황

  한빛상사 2026년 3월 판매현황입니다. 전임자가 만들어 두고 나갔습니다.
  보기에는 멀쩡한데 **계산이 하나도 안 됩니다.**

  [문제] 시트를 쓸 수 있는 표로 바꾸고, [최종확인] 의 다섯 칸을 채우세요.

■ 여덟 단계 — 순서대로 하세요

  ① 병합 풀기          B열 제품분류의 병합을 풉니다                (2-03)
  ② 빈 칸 채우기        드러난 빈 칸을 위 값으로 채웁니다            (2-04)
  ③ 숫자 고치기         D열 수량 중 글자로 된 것을 숫자로            (1-11)
  ④ 금액 내기          F열 = 수량 × 단가                        (1-10)
  ⑤ 영문 이름 뽑기      H열에 괄호 안 영문만                       (1-07)
  ⑥ 제품코드 나누기      J열 코드를 K·L·M 세 칸으로                 (2-12)
  ⑦ 커미션 내기         N열 = 금액 × 커미션비율(C3)                (1-09)
  ⑧ 최종확인 채우기      [최종확인] 시트 다섯 칸                     —

■ 순서를 지켜야 하는 이유

  · ②를 건너뛰면 ③·④가 일부 줄만 맞습니다.
  · ③을 건너뛰면 ④의 곱셈이 글자 × 숫자가 되어 틀립니다.
  · ⑦에서 $ 를 안 붙이면 아래로 갈수록 0 이 됩니다.

  막히면 그 단계 옆의 강의 번호를 보고 강의안을 다시 여세요.

■ 채점

  [최종확인] 다섯 칸만 맞으면 앞 단계는 다 맞은 것입니다.
  틀린 번호로 어느 단계에서 샜는지 알 수 있습니다. [정답] 시트에 적어 두었습니다.

  · 노란 칸 = 채울 곳     · 주황 칸 = 답을 적는 곳
  · 흰 칸은 자료입니다. 손대지 마세요."""


def sheet_guide(wb):
    ws = wb.create_sheet(u"안내")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 104
    t = ws.cell(row=1, column=2, value=u"종합문제 — 한빛상사 3월 판매현황")
    t.font = Font(name=FONT, size=18, bold=True, color=INK_HEAD)
    s = ws.cell(row=2, column=2, value=u"1장과 2장에서 배운 여덟 가지를 한 표에서 순서대로 씁니다.")
    s.font = Font(name=FONT, size=10, color=GREY)
    ws.merge_cells("B4:B32")
    c = ws.cell(row=4, column=2, value=GUIDE)
    c.font = Font(name=FONT, size=11)
    c.fill = PatternFill("solid", fgColor=NOTE)
    c.alignment = Alignment(vertical="top", wrap_text=True)
    c.border = THIN
    ws.row_dimensions[4].height = 450


def sheet_problem(wb):
    ws = wb.create_sheet(u"문제")
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 3, "B": 12, "C": 11, "D": 10, "E": 12, "F": 15,
                   "G": 20, "H": 14, "I": 3, "J": 15, "K": 8, "L": 8,
                   "M": 10, "N": 14}.items():
        ws.column_dimensions[col].width = w

    t = ws.cell(row=1, column=2, value=u"한빛상사 2026년 3월 판매현황")
    t.font = Font(name=FONT, size=14, bold=True)
    put(ws, 3, 2, u"커미션 비율", bold=True, align="left")
    put(ws, 3, 3, RATE, fmt=u"0.0%")

    head(ws, 5, 2, [u"제품분류", u"지점", u"수량", u"단가", u"금액",
                    u"담당자", u"담당자영문"])
    head(ws, 5, 10, [u"제품코드", u"분류", u"연도", u"일련", u"커미션"])

    for i, (cat, br, qty, price, nm, code) in enumerate(ROWS):
        r = FIRST + i
        if cat is not None:
            put(ws, r, 2, cat)
        else:
            put(ws, r, 2, None)                       # 병합으로 가려질 자리
        put(ws, r, 3, br)
        put(ws, r, 4, u"%d" % qty if i in AS_TEXT else qty, fmt=u"#,##0")
        put(ws, r, 5, price, fmt=u"#,##0")
        put(ws, r, 6, None, fill=YELLOW, fmt=u"#,##0")     # 금액
        put(ws, r, 7, nm, align="left")
        put(ws, r, 8, None, fill=YELLOW)                   # 담당자영문
        put(ws, r, 10, code)
        for c in (11, 12, 13):
            put(ws, r, c, None, fill=YELLOW)               # 코드 3조각
        put(ws, r, 14, None, fill=YELLOW, fmt=u"#,##0")    # 커미션

    # 제품분류 병합 — 이것이 ①단계
    ws.merge_cells(start_row=FIRST, start_column=2, end_row=FIRST + 2, end_column=2)
    ws.merge_cells(start_row=FIRST + 3, start_column=2, end_row=FIRST + 4, end_column=2)
    ws.merge_cells(start_row=FIRST + 5, start_column=2, end_row=FIRST + 6, end_column=2)

    last = FIRST + len(ROWS) - 1
    put(ws, last + 2, 2, u"합계", bold=True)
    put(ws, last + 2, 6, u"=SUM(F%d:F%d)" % (FIRST, last), bold=True, fmt=u"#,##0")
    put(ws, last + 2, 14, u"=SUM(N%d:N%d)" % (FIRST, last), bold=True, fmt=u"#,##0")

    g = ws.cell(row=3, column=6,
                value=u"노란 칸을 채우세요. [안내] 시트의 여덟 단계를 순서대로.")
    g.font = Font(name=FONT, size=10, color=GREY)
    return ws


QUESTIONS = [
    (u"① 선풍기 수량 합계는?", 404, u"0",
     u"②를 건너뛰면 142 만 나옵니다. 병합을 풀고 빈 칸을 채워야 404."),
    (u"② 선풍기 금액 합계는?", 15352000, u"#,##0",
     u"③을 건너뛰면 글자로 된 줄이 빠져 5,396,000 + 5,738,000 만 잡힙니다."),
    (u"③ 전체 금액 합계는?", 581752000, u"#,##0",
     u"④ 금액을 다 채웠으면 문제 시트 합계 줄과 같아야 합니다."),
    (u"④ 커미션 합계는?", 14543800, u"#,##0",
     u"⑦에서 $ 를 안 붙이면 아래로 갈수록 0 이 되어 훨씬 작게 나옵니다."),
    (u"⑤ 담당자 성(영문)이 Kim 인 사람은 몇 명?", 3, u"0",
     u"⑤를 해야 셀 수 있습니다. SM.Kim · JS.Kim · BM.Kim 셋."),
]


def sheet_check(wb):
    ws = wb.create_sheet(u"최종확인")
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 3, "B": 44, "C": 18, "D": 3, "E": 56}.items():
        ws.column_dimensions[col].width = w
    t = ws.cell(row=1, column=2, value=u"최종확인")
    t.font = Font(name=FONT, size=16, bold=True, color=INK_HEAD)
    s = ws.cell(row=2, column=2,
                value=u"주황 칸에 답을 적으세요. 다섯 칸이 맞으면 앞 단계는 다 맞은 것입니다.")
    s.font = Font(name=FONT, size=10, color=GREY)
    head(ws, 4, 2, [u"묻는 것", u"내 답"])
    for i, (q, _, fmt, _) in enumerate(QUESTIONS):
        r = 5 + i
        put(ws, r, 2, q, align="left")
        put(ws, r, 3, None, fill=ORANGE, fmt=fmt)
    n = ws.cell(row=11, column=2,
                value=u"다 적었으면 [정답] 시트를 여세요. 틀린 번호로 어느 단계에서 "
                      u"샜는지 알 수 있습니다.")
    n.font = Font(name=FONT, size=10, color=GREY)
    return ws


STEPS = [
    (u"①", u"병합 풀기", u"2-03", u"B6:B12 를 잡고 「병합하고 가운데 맞춤」 을 다시 눌러 끕니다."),
    (u"②", u"빈 칸 채우기", u"2-04",
     u"B6:B12 를 잡고 F5 → 옵션 → 빈 셀 → = 치고 ↑ → Ctrl+Enter. "
     u"그다음 복사 → 선택하여 붙여넣기 → 값 으로 굳힙니다."),
    (u"③", u"숫자 고치기", u"1-11",
     u"D7 · D9 · D11 이 글자입니다. D6:D12 를 잡고 느낌표 → 「숫자로 변환」."),
    (u"④", u"금액", u"1-10", u"F6 = D6*E6 을 넣고 F12 까지 끕니다."),
    (u"⑤", u"영문 이름", u"1-07",
     u"H6 에 SM.Kim 을 직접 치고 Ctrl+E. 괄호 안만 뽑힙니다."),
    (u"⑥", u"제품코드 나누기", u"2-12",
     u"J6:J12 를 잡고 데이터 → 텍스트 나누기 → 구분 기호 → 기타 에 - 를 넣습니다. "
     u"결과가 K·L·M 로 들어갑니다."),
    (u"⑦", u"커미션", u"1-09",
     u"N6 = F6*$C$3 을 넣고 N12 까지 끕니다. $ 를 빠뜨리면 아래가 0 이 됩니다."),
    (u"⑧", u"최종확인", u"—", u"[최종확인] 시트 주황 칸 다섯 개를 채웁니다."),
]


def sheet_answer(wb):
    ws = wb.create_sheet(u"정답")
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 3, "B": 5, "C": 18, "D": 9, "E": 74}.items():
        ws.column_dimensions[col].width = w
    t = ws.cell(row=1, column=2, value=u"정답")
    t.font = Font(name=FONT, size=16, bold=True, color=INK_HEAD)
    s = ws.cell(row=2, column=2, value=u"먼저 다 풀고 나서 보세요.")
    s.font = Font(name=FONT, size=10, color=GREY)

    head(ws, 4, 2, [u"단계", u"무엇", u"강의", u"어떻게"])
    for i, (no, what, lec, how) in enumerate(STEPS):
        r = 5 + i
        put(ws, r, 2, no)
        put(ws, r, 3, what, align="left")
        put(ws, r, 4, lec)
        c = put(ws, r, 5, how, align="left", wrap=True)
        ws.row_dimensions[r].height = 32

    r0 = 5 + len(STEPS) + 1
    h = ws.cell(row=r0, column=2, value=u"최종확인 답")
    h.font = Font(name=FONT, size=12, bold=True)
    head(ws, r0 + 1, 2, [u"번호", u"답", u"", u"틀렸다면"])
    for i, (q, a, fmt, why) in enumerate(QUESTIONS):
        r = r0 + 2 + i
        put(ws, r, 2, u"%d" % (i + 1))
        put(ws, r, 3, a, bold=True, fmt=fmt)
        put(ws, r, 4, None)
        c = put(ws, r, 5, why, align="left", wrap=True)
        ws.row_dimensions[r].height = 30
    return ws


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args()

    wb = Workbook()
    wb.remove(wb.active)
    sheet_guide(wb)
    sheet_problem(wb)
    sheet_check(wb)
    sheet_answer(wb)

    print(u"%s" % OUT)
    print(u"   시트 %d장 — %s" % (len(wb.sheetnames), u" · ".join(wb.sheetnames)))
    print(u"   자료 %d줄 · 단계 %d개 · 최종확인 %d문항"
          % (len(ROWS), len(STEPS), len(QUESTIONS)))
    print(u"   답: %s" % u" / ".join(format(q[1], ',') for q in QUESTIONS))
    if a.write:
        wb.save(OUT)
        print(u"\n만들었습니다.")
    else:
        print(u"\n※ 보기만 했습니다. 실제로 만들려면 --write 를 붙이세요.")


if __name__ == "__main__":
    main()
