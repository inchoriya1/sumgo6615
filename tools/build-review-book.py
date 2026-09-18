# -*- coding: utf-8 -*-
u"""
장마다 복습용 엑셀 파일을 한 개씩 만듭니다.

  python tools/build-review-book.py            # 보기만
  python tools/build-review-book.py --write    # 실제로 만들기

왜 필요한가
  예제는 **강의를 따라가며 한 번 해 보는 것**이라 파일이 열여섯 개로 흩어져
  있습니다. 나중에 다시 훑어보려면 **어느 파일을 열어야 할지부터** 막힙니다.

  그래서 장마다 **한 파일**로 모읍니다. 시트를 강의 순서로 놓고,
  **노란 칸만 채우면** 되게 했습니다.

구성
  [안내]   푸는 법 · 채점하는 법 · 막혔을 때 어디를 볼지
  [1-05]   강의별 문제 시트. 자료는 그대로 두고 **답 칸만 비워** 둡니다
    ⋮
  [정답]   시트마다 **답과 한 줄 설명**. 왜 그렇게 되는지까지

원칙
  · **정답 시트는 맨 뒤 한 장**입니다. 문제 옆에 두면 눈이 갑니다.
  · 노란 칸 = 채울 곳. 강의안의 색 규칙과 같습니다.
  · 문제는 **손이 기억할 것**만 냅니다. 메뉴 위치를 외우는 문제는 안 냅니다.
"""
import argparse
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as gl

FONT = u"맑은 고딕"
INK_HEAD = u"FF1F4E79"
YELLOW = u"FFFFE699"
NOTE = u"FFFFF9E6"
GREY = u"FF808080"
THIN = Border(*(Side(style="thin"),) * 4)

OUT = u"강의예제"


def head(ws, row, cols, labels):
    f = Font(name=FONT, size=9, bold=True, color="FFFFFFFF")
    fill = PatternFill("solid", fgColor=INK_HEAD)
    mid = Alignment(horizontal="center", vertical="center")
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=cols + i, value=lab)
        c.font, c.fill, c.alignment, c.border = f, fill, mid, THIN


def put(ws, row, col, value, fill=None, bold=False, fmt=None, wrap=False,
        align_left=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name=FONT, size=11, bold=bold)
    c.border = THIN
    c.alignment = Alignment(horizontal="left" if align_left else "center",
                            vertical="center", wrap_text=wrap)
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if fmt:
        c.number_format = fmt
    return c


def title(ws, text, sub=None):
    t = ws.cell(row=1, column=2, value=text)
    t.font = Font(name=FONT, size=14, bold=True)
    if sub:
        s = ws.cell(row=2, column=2, value=sub)
        s.font = Font(name=FONT, size=10, color=GREY)
    ws.sheet_view.showGridLines = False


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


# ── 안내 시트 ──────────────────────────────────────────────────────
GUIDE = u"""■ 이 파일은 복습용입니다

  강의를 한 번 본 뒤, 며칠 지나 다시 풀어 보는 용도입니다.
  강의예제 폴더의 원본은 그대로 두고 여기서만 푸세요.

■ 푸는 법

  · 시트는 강의 순서입니다. 앞에서부터 하나씩 하세요.
  · **노란 칸만 채우면 됩니다.** 흰 칸은 자료이니 손대지 마세요.
  · 한 시트가 한 강의입니다. 막히면 그 강의안을 다시 보세요.

■ 채점하는 법

  · **맨 뒤 [정답] 시트**에 시트별 답과 한 줄 설명이 있습니다.
  · 먼저 다 풀고 나서 보세요. 옆에 두고 보면 남지 않습니다.
  · 답이 달라도 **결과가 같으면 맞은 것**입니다.
    (수식은 여러 길이 있습니다)

■ 다시 풀고 싶으면

  노란 칸을 지우고 처음부터 하면 됩니다.
  자료는 안 건드렸으니 파일이 망가지지 않습니다."""


def sheet_guide(wb, chap):
    ws = wb.create_sheet(u"안내")
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 3, "B": 100})
    t = ws.cell(row=1, column=2, value=u"%d장 복습" % chap)
    t.font = Font(name=FONT, size=18, bold=True, color=INK_HEAD)
    ws.merge_cells("B3:B20")
    c = ws.cell(row=3, column=2, value=GUIDE)
    c.font = Font(name=FONT, size=11)
    c.fill = PatternFill("solid", fgColor=NOTE)
    c.alignment = Alignment(vertical="top", wrap_text=True)
    c.border = THIN
    ws.row_dimensions[3].height = 300
    return ws


# ── 1장 문제 ──────────────────────────────────────────────────────
def chapter1(wb):
    ans = []

    # 1-05 자동 채우기
    ws = wb.create_sheet(u"1-05 채우기")
    title(ws, u"1-05  자동 채우기", u"노란 칸을 채우세요. 끌어서 채워도 되고 직접 쳐도 됩니다.")
    widths(ws, {"A": 3, "B": 14, "C": 14, "D": 14, "E": 14, "F": 3, "G": 46})
    head(ws, 4, 2, [u"한 칸(사과)", u"두 칸(사과·배)", u"한 칸(1)", u"두 칸(1·3)"])
    put(ws, 5, 2, u"사과"); put(ws, 5, 3, u"사과"); put(ws, 5, 4, 1); put(ws, 5, 5, 1)
    put(ws, 6, 3, u"배"); put(ws, 6, 5, 3)
    for r in range(6, 11):
        for c in (2, 3, 4, 5):
            if ws.cell(r, c).value is None:
                put(ws, r, c, None, fill=YELLOW)
    g = ws.cell(row=4, column=7, value=u"왼쪽 두 열은 글자, 오른쪽 두 열은 숫자입니다.\n"
                                       u"한 칸만 잡고 끌 때와 두 칸을 잡고 끌 때가\n어떻게 다른지 채워 보세요.")
    g.font = Font(name=FONT, size=10, color=GREY)
    g.alignment = Alignment(vertical="top", wrap_text=True)
    ans.append((u"1-05 채우기",
                u"한 칸 = 복사 / 두 칸 = 규칙",
                u"B: 사과가 계속 · C: 사과·배 번갈아 · D: 1이 계속 · E: 5,7,9,11,13 (2씩)",
                u"숫자는 두 칸을 보여 줘야 규칙을 압니다. 요일·월은 한 칸이면 압니다."))

    # 1-07 빠른 채우기
    ws = wb.create_sheet(u"1-07 빠른채우기")
    title(ws, u"1-07  빠른 채우기", u"노란 칸을 Ctrl+E 로 채우세요. 첫 줄은 직접 쳐야 합니다.")
    widths(ws, {"A": 3, "B": 26, "C": 18, "D": 18, "E": 3, "F": 44})
    head(ws, 4, 2, [u"직원명", u"영문이름", u"지역번호"])
    rows = [u"최경민 (Stephen.Cho)", u"박재현 (JH.Park)", u"김인후 (John.Kim)",
            u"김민형 (MinHyung.Kim)", u"이승오 (Robert.Lee)", u"정수미 (Sara.Jeong)"]
    tel = [u"02-987-8182", u"02-987-7827", u"02-987-4453",
           u"02-987-8966", u"02-987-6745", u"031-654-6219"]
    for i, (nm, t) in enumerate(zip(rows, tel)):
        put(ws, 5 + i, 2, nm)
        put(ws, 5 + i, 3, None, fill=YELLOW)
        put(ws, 5 + i, 4, None, fill=YELLOW)
    put(ws, 12, 2, u"연락처 →", bold=True)
    for i, t in enumerate(tel):
        put(ws, 13 + i, 2, t)
    g = ws.cell(row=4, column=6, value=u"C열: 괄호 안 영문만 뽑으세요.\n"
                                       u"D열: 연락처(아래 12행부터)의 지역번호만 뽑으세요.\n\n"
                                       u"마지막 줄은 031 입니다 — 다 뽑고 나서\n맨 아래를 꼭 확인하세요.")
    g.font = Font(name=FONT, size=10, color=GREY)
    g.alignment = Alignment(vertical="top", wrap_text=True)
    ans.append((u"1-07 빠른채우기",
                u"C5 에 Stephen.Cho 를 치고 Ctrl+E",
                u"D열은 02·02·02·02·02·031",
                u"맨 아래가 031 입니다. 예를 02 로만 보여 주면 틀릴 수 있어 확인이 필요합니다."))

    # 1-09 셀 참조
    ws = wb.create_sheet(u"1-09 셀 참조")
    title(ws, u"1-09  셀 참조", u"노란 칸에 수식을 넣고 끌어 채우세요. F4 를 쓰세요.")
    widths(ws, {"A": 3, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 3, "H": 44})
    put(ws, 4, 2, u"커미션 비율", bold=True)
    put(ws, 4, 3, 0.1, fmt=u"0.0%")
    head(ws, 6, 2, [u"이름", u"상반기", u"하반기", u"매출액", u"커미션"])
    data = [(u"김세민", 192000, 167000), (u"정다온", 113000, 107000),
            (u"김진선", 136000, 167000), (u"정희엘", 161000, 125000),
            (u"박단비", 100000, 155000)]
    for i, (nm, a, b) in enumerate(data):
        r = 7 + i
        put(ws, r, 2, nm); put(ws, r, 3, a, fmt=u"#,##0"); put(ws, r, 4, b, fmt=u"#,##0")
        put(ws, r, 5, None, fill=YELLOW, fmt=u"#,##0")
        put(ws, r, 6, None, fill=YELLOW, fmt=u"#,##0")
    put(ws, 14, 2, u"혼합참조", bold=True)
    head(ws, 15, 2, [u"판매량\\비율", u"0.03", u"0.05", u"0.07"])
    for i, q in enumerate([1000, 3000, 5000, 7000]):
        r = 16 + i
        put(ws, r, 2, q, fmt=u"#,##0")
        for c in (3, 4, 5):
            put(ws, r, c, None, fill=YELLOW, fmt=u"#,##0")
    g = ws.cell(row=4, column=8, value=u"위: 매출액 = 상반기 + 하반기,\n"
                                       u"커미션 = 매출액 × C4.\n"
                                       u"끌어도 C4 가 안 움직여야 합니다.\n\n"
                                       u"아래: 수식 하나를 C16 에 넣고\n"
                                       u"오른쪽·아래로 끌어 12칸을 채우세요.")
    g.font = Font(name=FONT, size=10, color=GREY)
    g.alignment = Alignment(vertical="top", wrap_text=True)
    ans.append((u"1-09 셀 참조",
                u"E7 = C7+D7 · F7 = E7*$C$4",
                u"C16 = $B16*C$15 하나로 12칸 전부",
                u"세로로 읽는 값은 열 고정($B), 가로로 읽는 값은 행 고정(C$15)."))

    # 1-11 오류
    ws = wb.create_sheet(u"1-11 오류")
    title(ws, u"1-11  오류와 초록 삼각형", u"합계가 왜 0 인지 찾아 고치고, 오른쪽 표를 채우세요.")
    widths(ws, {"A": 3, "B": 16, "C": 12, "D": 12, "E": 12, "F": 3, "G": 16, "H": 30})
    head(ws, 4, 2, [u"소분류", u"1월", u"2월", u"3월"])
    for i, (nm, a, b, c) in enumerate([(u"정수기필터", 52, 91, 61), (u"슬로우쿠커", 93, 98, 85),
                                       (u"김치냉장고", 58, 64, 89), (u"에그마스터", 63, 90, 67)]):
        r = 5 + i
        put(ws, r, 2, nm)
        for j, v in enumerate((a, b, c)):
            put(ws, r, 3 + j, u"%d" % v)      # 일부러 글자로 넣습니다
    put(ws, 10, 2, u"합계", bold=True)
    for c in (3, 4, 5):
        put(ws, 10, c, None, fill=YELLOW)
    head(ws, 4, 7, [u"오류", u"무슨 뜻인가"])
    for i, e in enumerate([u"#DIV/0!", u"#VALUE!", u"#REF!", u"#NAME?", u"#N/A"]):
        put(ws, 5 + i, 7, e)
        put(ws, 5 + i, 8, None, fill=YELLOW)
    ans.append((u"1-11 오류",
                u"숫자가 글자로 들어 있어 SUM 이 0",
                u"범위를 잡고 느낌표 → 숫자로 변환. 합계 266 · 343 · 302",
                u"#DIV/0! 0으로 나눔 · #VALUE! 종류 안 맞음 · #REF! 참조 사라짐 · "
                u"#NAME? 함수 이름 오타 · #N/A 찾는 값 없음"))
    return ans


# ── 2장 문제 ──────────────────────────────────────────────────────
def chapter2(wb):
    ans = []

    ws = wb.create_sheet(u"2-01 블록쌓기")
    title(ws, u"2-01  세로방향 블록쌓기", u"왼쪽 표의 문제를 찾아 오른쪽에 다시 쌓으세요.")
    widths(ws, {"A": 3, "B": 12, "C": 12, "D": 10, "E": 10, "F": 3,
                "G": 12, "H": 12, "I": 10, "J": 8, "K": 10, "L": 3, "M": 40})
    put(ws, 4, 2, u"커미션", bold=True)
    put(ws, 4, 4, 0.03, fmt=u"0.00%"); put(ws, 4, 5, 0.035, fmt=u"0.00%")
    head(ws, 5, 2, [u"지사", u"영업점", u"1월", u"2월"])
    for i, (a, b, c, d) in enumerate([(u"강남지사", u"대치점", 2194, 1278),
                                      (u"강남지사", u"신사점", 1254, 2130),
                                      (u"서초지사", u"방배점", 1598, 1101)]):
        r = 6 + i
        put(ws, r, 2, a); put(ws, r, 3, b)
        put(ws, r, 4, c, fmt=u"#,##0"); put(ws, r, 5, d, fmt=u"#,##0")
    head(ws, 5, 7, [u"지사", u"영업점", u"커미션", u"월", u"매출"])
    for r in range(6, 12):
        for c in range(7, 12):
            put(ws, r, c, None, fill=YELLOW)
    g = ws.cell(row=4, column=13, value=u"왼쪽은 커미션이 표 위에 얹혀 있고\n"
                                        u"월이 가로로 벌어져 있습니다.\n\n"
                                        u"오른쪽에 한 줄 = 한 건 으로 다시 쌓으세요.\n"
                                        u"여섯 줄이 나옵니다.")
    g.font = Font(name=FONT, size=10, color=GREY)
    g.alignment = Alignment(vertical="top", wrap_text=True)
    ans.append((u"2-01 블록쌓기",
                u"커미션을 열로 내리고 월을 세로로 쌓습니다",
                u"강남지사·대치점·3.00%·1월·2194 … 이렇게 여섯 줄",
                u"머리글 한 줄 · 한 칸에 한 가지 · 아래로만 쌓기."))

    ws = wb.create_sheet(u"2-04 빈 칸")
    title(ws, u"2-04  빈 칸 찾고 채우기", u"노란 칸을 채우고, 합계가 어떻게 바뀌는지 보세요.")
    widths(ws, {"A": 3, "B": 14, "C": 12, "D": 12, "E": 3, "F": 44})
    head(ws, 4, 2, [u"제품명", u"지점", u"판매수량"])
    rows = [(u"선풍기", u"신촌점", 142), (None, u"대방점", 111), (None, u"구로점", 151),
            (u"TV", u"신촌점", 163), (None, u"구로점", 125),
            (u"에어컨", u"대방점", 133), (None, u"영등포점", 135)]
    for i, (a, b, c) in enumerate(rows):
        r = 5 + i
        if a is None:
            put(ws, r, 2, None, fill=YELLOW)
        else:
            put(ws, r, 2, a)
        put(ws, r, 3, b); put(ws, r, 4, c, fmt=u"#,##0")
    put(ws, 13, 2, u"선풍기 합계", bold=True)
    put(ws, 13, 4, u'=SUMIF($B$5:$B$11,"선풍기",$D$5:$D$11)', fmt=u"#,##0")
    g = ws.cell(row=4, column=6, value=u"노란 칸은 병합을 푼 자리입니다.\n"
                                       u"F5 → 빈 셀 → = ↑ → Ctrl+Enter 로\n한 번에 채우세요.\n\n"
                                       u"채우기 전과 후에 13행 합계가\n어떻게 다른지 보세요.")
    g.font = Font(name=FONT, size=10, color=GREY)
    g.alignment = Alignment(vertical="top", wrap_text=True)
    ans.append((u"2-04 빈 칸",
                u"F5 → 빈 셀 → = ↑ → Ctrl+Enter",
                u"합계가 142 에서 404 로 바뀝니다",
                u"채운 뒤 복사 → 선택하여 붙여넣기 → 값 으로 굳혀야 정렬에 안 무너집니다."))

    ws = wb.create_sheet(u"2-08 행열 변환")
    title(ws, u"2-08  행열 변환", u"위 표를 아래에 눕혀 붙이세요.")
    widths(ws, {"A": 3, "B": 14, "C": 12, "D": 12, "E": 12, "F": 3, "G": 40})
    head(ws, 4, 2, [u"항목", u"강남점", u"서초점", u"영등포점"])
    for i, (nm, a, b, c) in enumerate([(u"면적(평)", 250, 210, 320),
                                       (u"임대료", 3530, 3300, 2750),
                                       (u"유동인구", 75000, 68000, 55000)]):
        r = 5 + i
        put(ws, r, 2, nm)
        for j, v in enumerate((a, b, c)):
            put(ws, r, 3 + j, v, fmt=u"#,##0")
    put(ws, 10, 2, u"↓ 여기에 눕혀 붙이세요", bold=True)
    for r in range(11, 15):
        for c in range(2, 6):
            put(ws, r, c, None, fill=YELLOW)
    g = ws.cell(row=4, column=7, value=u"복사 → 오른쪽 클릭 →\n선택하여 붙여넣기 →\n"
                                       u"「행/열 바꿈」 체크.\n\n매장이 세로로 서야 합니다.")
    g.font = Font(name=FONT, size=10, color=GREY)
    g.alignment = Alignment(vertical="top", wrap_text=True)
    ans.append((u"2-08 행열 변환",
                u"선택하여 붙여넣기 → 「행/열 바꿈」",
                u"첫 줄이 매장 · 면적 · 임대료 · 유동인구 가 됩니다",
                u"붙여넣기는 한 번 하면 끝입니다. 연동하려면 TRANSPOSE 를 쓰세요."))

    ws = wb.create_sheet(u"2-12 텍스트 나누기")
    title(ws, u"2-12  텍스트 나누기", u"왼쪽 값을 오른쪽 노란 칸으로 나누세요.")
    widths(ws, {"A": 3, "B": 28, "C": 3, "D": 12, "E": 12, "F": 12, "G": 12, "H": 3, "I": 40})
    put(ws, 4, 2, u"세 글자씩 붙은 것", bold=True)
    for i, v in enumerate([u"무화과오렌지복숭아토마토", u"한라봉바나나두리안산딸기",
                           u"양상추양배추단무지양송이"]):
        put(ws, 5 + i, 2, v)
        for c in range(4, 8):
            put(ws, 5 + i, c, None, fill=YELLOW)
    put(ws, 10, 2, u"기호로 나뉜 것", bold=True)
    for i, v in enumerate([u"CA-2020-152156", u"CN-2018-115812", u"KR-2016-115812"]):
        put(ws, 11 + i, 2, v)
        for c in range(4, 7):
            put(ws, 11 + i, c, None, fill=YELLOW)
    g = ws.cell(row=4, column=9, value=u"위: 구분 기호가 없습니다 →\n「너비가 일정함」\n\n"
                                       u"아래: - 로 나뉩니다 →\n「구분 기호로 분리됨 → 기타」\n\n"
                                       u"3단계에서 서식을 정할 수 있습니다.")
    g.font = Font(name=FONT, size=10, color=GREY)
    g.alignment = Alignment(vertical="top", wrap_text=True)
    ans.append((u"2-12 텍스트 나누기",
                u"위는 「너비가 일정함」, 아래는 「구분 기호」",
                u"위: 무화과·오렌지·복숭아·토마토 / 아래: CA·2020·152156",
                u"나누면 오른쪽 칸을 덮어씁니다. 빈 열을 미리 만들어 두세요."))

    ws = wb.create_sheet(u"2-13 여러 줄")
    title(ws, u"2-13  여러 줄 합치기·나누기", u"왼쪽을 나누고, 오른쪽을 합치세요.")
    widths(ws, {"A": 3, "B": 16, "C": 14, "D": 14, "E": 3, "F": 12, "G": 12, "H": 18, "I": 3, "J": 38})
    put(ws, 4, 2, u"한 칸에 두 줄 → 나누기", bold=True)
    for i, v in enumerate([u"김세민\n정다온", u"박정화\n이서우", u"정진하\n김병민"]):
        c = put(ws, 5 + i, 2, v)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        put(ws, 5 + i, 3, None, fill=YELLOW)
        put(ws, 5 + i, 4, None, fill=YELLOW)
    put(ws, 10, 6, u"두 칸 → 한 칸으로 합치기", bold=True)
    head(ws, 11, 6, [u"성", u"이름", u"합친 것"])
    for i, (a, b) in enumerate([(u"김", u"수호"), (u"박", u"희선"), (u"정", u"도균")]):
        put(ws, 12 + i, 6, a); put(ws, 12 + i, 7, b)
        put(ws, 12 + i, 8, None, fill=YELLOW)
    g = ws.cell(row=4, column=10, value=u"왼쪽: 텍스트 나누기 → 기타 칸에서\nCtrl+J (아무것도 안 보입니다)\n\n"
                                        u"오른쪽: 두 줄로 합치려면\n=F12&CHAR(10)&G12 그리고\n"
                                        u"홈 → 맞춤 → 텍스트 줄 바꿈")
    g.font = Font(name=FONT, size=10, color=GREY)
    g.alignment = Alignment(vertical="top", wrap_text=True)
    ans.append((u"2-13 여러 줄",
                u"나누기는 Ctrl+J · 합치기는 CHAR(10)",
                u"Ctrl+J 는 화면에 안 보입니다. 점 하나가 깜빡일 뿐입니다",
                u"CHAR(10) 으로 합친 뒤 「텍스트 줄 바꿈」을 켜야 두 줄로 보입니다."))
    return ans


def guide_box(ws, row, col, text):
    u"""문제 옆에 붙이는 회색 설명 상자."""
    g = ws.cell(row=row, column=col, value=text)
    g.font = Font(name=FONT, size=10, color=GREY)
    g.alignment = Alignment(vertical="top", wrap_text=True)
    return g


# ── 3장 문제 ──────────────────────────────────────────────────────
def chapter3(wb):
    ans = []

    ws = wb.create_sheet(u"3-02 표시형식")
    title(ws, u"3-02  사용자 지정 표시 형식",
          u"노란 칸에 「쓸 서식」을 적으세요. 값은 왼쪽 그대로입니다.")
    widths(ws, {"A": 3, "B": 16, "C": 22, "D": 26, "E": 3, "F": 46})
    head(ws, 4, 2, [u"값", u"이렇게 보이게", u"쓸 서식"])
    rows = [(1500, u"1,500"), (-800, u"빨간 -800"), (0, u"-  (하이픈)"),
            (5930, u"005930"), (3500000, u"일금 삼백오십만원정")]
    for i, (v, want) in enumerate(rows):
        r = 5 + i
        put(ws, r, 2, v)
        put(ws, r, 3, want)
        put(ws, r, 4, None, fill=YELLOW)
    put(ws, 11, 2, u"날짜", bold=True)
    head(ws, 12, 2, [u"값", u"이렇게 보이게", u"쓸 서식"])
    put(ws, 13, 2, u"2021-01-01"); put(ws, 13, 3, u"2021년 01월")
    put(ws, 13, 4, None, fill=YELLOW)
    put(ws, 14, 2, u"2021-01-01"); put(ws, 14, 3, u"2021-01-01 (금)")
    put(ws, 14, 4, None, fill=YELLOW)
    guide_box(ws, 4, 6,
              u"세미콜론은 양수;음수;0;글자 순서입니다.\n"
              u"칸을 덜 써도 되고, 비워 두면 그 경우는 안 보입니다.\n\n"
              u"요일은 aaa, 앞자리 0은 0 을 자릿수만큼,\n"
              u"한글 금액은 [DBNum4] 입니다.")
    ans.append((u"3-02 표시형식",
                u"세미콜론 네 칸 = 양수·음수·0·글자",
                u"#,##0 / [빨강]-#,##0 / - / 000000 / \"일금 \"[DBNum4]G/표준\"원정\" "
                u"/ yyyy\"년\" mm\"월\" / yyyy-mm-dd (aaa)",
                u"값은 그대로입니다. 칸을 누르면 수식 입력줄에 원래 숫자가 보입니다."))

    ws = wb.create_sheet(u"3-03 보고서 규칙")
    title(ws, u"3-03  보고서 작성 규칙",
          u"표를 읽히게 고치세요. 고칠 곳마다 노란 칸에 「무엇을 했는지」 적습니다.")
    widths(ws, {"A": 3, "B": 18, "C": 16, "D": 16, "E": 3, "F": 30, "G": 3, "H": 46})
    put(ws, 3, 2, u"최근 3년간 매출", bold=True)
    head(ws, 4, 2, [u"항목", u"2023", u"2024"])
    for i, (nm, a, b) in enumerate([(u"매출액", 1444216203, 1393409033),
                                    (u"매출원가", 332915074, 146748600),
                                    (u"　음료", 125960000, 96208984),
                                    (u"　베이커리", 51821669, 36488800),
                                    (u"영업이익", 0, 0)]):
        r = 5 + i
        put(ws, r, 2, nm, align_left=True)
        put(ws, r, 3, a); put(ws, r, 4, b)
    head(ws, 11, 2, [u"고칠 곳", u"", u""])
    for i, q in enumerate([u"① 숫자 열 맞춤은?", u"② 천 단위 기호를 넣는 단축키는?",
                           u"③ 단위를 천으로 접는 서식은?", u"④ 들여쓰기는 무엇으로?",
                           u"⑤ 0 을 하이픈으로 보이는 서식은?"]):
        r = 12 + i
        put(ws, r, 2, q, align_left=True)
        c = ws.cell(row=r, column=3, value=None)
        c.fill = PatternFill("solid", fgColor=YELLOW)
        c.border = THIN
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    guide_box(ws, 4, 8,
              u"표의 「　음료」 「　베이커리」 앞에는 공백이 들어 있습니다.\n"
              u"이것이 왜 문제인지도 생각해 보세요.\n\n"
              u"(7장에서 TRIM 으로 떼어내게 됩니다)")
    ans.append((u"3-03 보고서 규칙",
                u"맞춤 · 천 단위 · 단위 · 들여쓰기 · 0",
                u"① 오른쪽 맞춤  ② Ctrl+Shift+1  ③ #,##0, (맨 뒤 쉼표)  "
                u"④ 홈 → 들여쓰기 늘림  ⑤ #,##0;-#,##0;-;@",
                u"공백으로 들여쓴 항목은 찾기·정렬·VLOOKUP 이 전부 어긋납니다."))

    ws = wb.create_sheet(u"3-04 조건부 서식")
    title(ws, u"3-04  조건부 서식",
          u"아래 표에 규칙을 걸고, 노란 칸에 「쓸 수식」을 적으세요.")
    widths(ws, {"A": 3, "B": 12, "C": 10, "D": 10, "E": 10, "F": 10, "G": 3, "H": 52})
    head(ws, 4, 2, [u"이름", u"1차", u"2차", u"3차", u"합계"])
    for i, (nm, a, b, c) in enumerate([(u"최효윤", 80, 81, 80), (u"이유림", 83, 76, 82),
                                       (u"정재현", 89, 85, 92), (u"김수호", 98, 98, 94),
                                       (u"박희선", 96, 73, 73), (u"이이서", 76, 73, 97)]):
        r = 5 + i
        put(ws, r, 2, nm); put(ws, r, 3, a); put(ws, r, 4, b); put(ws, r, 5, c)
        put(ws, r, 6, a + b + c)
    head(ws, 13, 2, [u"걸 규칙", u"", u"", u"", u""])
    for i, q in enumerate([u"① 합계가 240 미만인 줄 전체를 칠하려면?",
                           u"② 합계 상위 30% 만 칠하려면?",
                           u"③ 1차와 2차가 모두 80 이상인 줄을 칠하려면?"]):
        r = 14 + i
        put(ws, r, 2, q, align_left=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        c = ws.cell(row=r, column=5, value=None)
        c.fill = PatternFill("solid", fgColor=YELLOW)
        c.border = THIN
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
    guide_box(ws, 4, 8,
              u"줄 전체를 칠하려면 범위를 B5:F10 으로 <전부> 잡고\n"
              u"「수식을 사용하여 서식을 지정할 셀 결정」 을 씁니다.\n\n"
              u"$ 가 어디에 붙는지가 전부입니다.\n"
              u"열은 고정, 행은 안 고정입니다.")
    ans.append((u"3-04 조건부 서식",
                u"줄 전체는 수식 규칙 + $열",
                u"① =$F5<240   ② 상위/하위 규칙 → 상위 10% → 30 으로   ③ =AND($C5>=80,$D5>=80)",
                u"$F$5 로 쓰면 모든 칸이 한 칸만 봐서 전부 칠해지거나 전부 안 칠해집니다."))

    ws = wb.create_sheet(u"3-05 데이터 막대")
    title(ws, u"3-05  데이터 막대와 아이콘",
          u"규칙을 걸어 보고, 노란 칸에 답을 적으세요.")
    widths(ws, {"A": 3, "B": 12, "C": 14, "D": 14, "E": 3, "F": 52})
    head(ws, 4, 2, [u"직원명", u"실적매출", u"매출이익"])
    for i, (nm, a, b) in enumerate([(u"이연재", 2457000, 280000), (u"김고운", 1288000, 590000),
                                    (u"정소율", 1200000, -390000), (u"정준연", 2450000, 310000),
                                    (u"박소울", 2001000, -350000), (u"이하나", 2323000, 420000)]):
        r = 5 + i
        put(ws, r, 2, nm)
        put(ws, r, 3, a, fmt=u"#,##0"); put(ws, r, 4, b, fmt=u"#,##0")
    for i, q in enumerate([u"① 막대 범위에 넣으면 안 되는 줄은?",
                           u"② 음수 막대 색을 바꾸는 곳은?",
                           u"③ 아이콘 집합 기본 기준의 문제는?"]):
        r = 13 + i
        put(ws, r, 2, q, align_left=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=4, value=None)
        c.fill = PatternFill("solid", fgColor=YELLOW)
        c.border = THIN
    guide_box(ws, 4, 6,
              u"매출이익에는 음수가 셋 섞여 있습니다.\n"
              u"막대를 깔면 가운데 축이 생깁니다.\n\n"
              u"아이콘 집합은 기본이 「백분위」입니다.\n"
              u"모두 적자여도 상위 3분의 1에 초록불이 켜집니다.")
    ans.append((u"3-05 데이터 막대",
                u"합계 빼기 · 음수 축 · 기준을 숫자로",
                u"① 합계 줄  ② 규칙 편집 → 음수 값 및 축  ③ 백분위라 순위로만 나눔 → 「숫자」로 바꿈",
                u"합계를 넣으면 나머지 막대가 전부 짧아져 견줄 수 없습니다."))

    ws = wb.create_sheet(u"3-07 스파크라인")
    title(ws, u"3-07  스파크라인",
          u"트렌드 칸에 스파크라인을 그리고, 아래 물음에 답하세요.")
    widths(ws, {"A": 3, "B": 12, "C": 8, "D": 8, "E": 8, "F": 8, "G": 8, "H": 8,
                "I": 10, "J": 14, "K": 3, "L": 46})
    head(ws, 4, 2, [u"직원명", u"1월", u"2월", u"3월", u"4월", u"5월", u"6월", u"합계", u"트렌드"])
    data = [(u"정희엘", 146, 146, 83, 112, 146, 117),
            (u"박단비", 106, 94, 81, 114, 82, 102),
            (u"정진하", 90, 135, 105, 91, 131, 110),
            (u"김병민", 114, 124, 87, 113, 88, 139)]
    for i, row in enumerate(data):
        r = 5 + i
        put(ws, r, 2, row[0])
        for j, v in enumerate(row[1:]):
            put(ws, r, 3 + j, v)
        put(ws, r, 9, sum(row[1:]))
        put(ws, r, 10, None, fill=YELLOW)
    for i, q in enumerate([u"① 데이터 범위에 합계를 넣으면?",
                           u"② 줄끼리 견주려면 무엇을 맞춰야 하나?",
                           u"③ Delete 로 안 지워지는 이유는?"]):
        r = 11 + i
        put(ws, r, 2, q, align_left=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        c = ws.cell(row=r, column=6, value=None)
        c.fill = PatternFill("solid", fgColor=YELLOW)
        c.border = THIN
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=9)
    guide_box(ws, 4, 12,
              u"J5 를 누르고 삽입 → 스파크라인 → 꺾은선형.\n"
              u"데이터 범위는 C5:H5 입니다.\n\n"
              u"그린 뒤 스파크라인 탭에서\n높은 점·낮은 점을 켜 보세요.")
    ans.append((u"3-07 스파크라인",
                u"범위 · 축 · 지우기",
                u"① 합계가 커서 나머지가 눌립니다  ② 세로 축을 「모든 스파크라인에 대해 동일하게」  "
                u"③ 칸의 값이 아니라 배경 그림이라 스파크라인 탭 → 지우기 를 써야 함",
                u"합계가 같아도 흐름이 다른 사람을 가려내는 것이 이 기능의 쓰임입니다."))
    return ans


# ── 4장 문제 ──────────────────────────────────────────────────────
def chapter4(wb):
    ans = []

    ws = wb.create_sheet(u"4-01 파일 이름")
    title(ws, u"4-01  파일 이름 짓기",
          u"왼쪽 이름을 규칙에 맞게 고쳐 노란 칸에 적으세요.")
    widths(ws, {"A": 3, "B": 40, "C": 40, "D": 3, "E": 50})
    head(ws, 4, 2, [u"받은 이름", u"고친 이름"])
    for i, nm in enumerate([u"220101_예제파일_v1.xlsx", u"20220107_예제파일_v4.xlsx",
                            u"예제파일_정말진짜레알최종.xlsx",
                            u"3월실적(수정본)(2).xlsx", u"김대리_보고서_최종최종.xlsx"]):
        r = 5 + i
        put(ws, r, 2, nm, align_left=True)
        put(ws, r, 3, None, fill=YELLOW)
    guide_box(ws, 4, 5,
              u"규칙은 한 줄입니다 —\n"
              u"    YYYYMMDD_무엇_누가.xlsx\n\n"
              u"· 날짜는 여덟 자리로 맨 앞에\n"
              u"· 「최종」은 쓰지 않습니다\n"
              u"· 버전은 _v01 처럼 두 자리로")
    ans.append((u"4-01 파일 이름",
                u"YYYYMMDD_무엇_누가",
                u"20220101_예제파일_v01 / 20220107_예제파일_v04 / "
                u"20220101_예제파일_v06 / 20260301_3월실적_v02 / 20260301_보고서_김대리_v02",
                u"여덟 자리로 맞춰야 이름순이 곧 날짜순이 됩니다. v2 와 v10 은 v10 이 먼저 섭니다."))

    ws = wb.create_sheet(u"4-04 외부 참조")
    title(ws, u"4-04  외부 참조 찾기",
          u"아래 수식 중 「다른 파일을 보는 것」에 O 를 적으세요.")
    widths(ws, {"A": 3, "B": 48, "C": 10, "D": 3, "E": 50})
    head(ws, 4, 2, [u"수식", u"외부?"])
    forms = [(u"=SUM(D7:D14)", u""),
             (u"=SUM([1]Sheet1!$A$1:$A$7)", u"O"),
             (u"=VLOOKUP(B5,제품정보!A:C,2,0)", u""),
             (u"='C:\\실적\\[3월.xlsx]Sheet1'!$B$2", u"O"),
             (u"=INDIRECT(\"'\"&$D$4&\"'!A:C\")", u""),
             (u"=[예산.xlsx]요약!$C$9*1.1", u"O")]
    for i, (f, _) in enumerate(forms):
        r = 5 + i
        put(ws, r, 2, f, align_left=True)
        put(ws, r, 3, None, fill=YELLOW)
    guide_box(ws, 4, 5,
              u"대괄호 [ ] 가 「다른 파일」 표시입니다.\n"
              u"시트 이름 앞의 느낌표(!)만으로는 외부가 아닙니다.\n\n"
              u"확인하는 법 — 데이터 → 연결 편집.\n"
              u"단추가 회색이면 연결이 없는 것입니다.\n"
              u"또는 Ctrl+F 에서 범위=통합 문서, 찾는 위치=수식 으로 [ 찾기.")
    ans.append((u"4-04 외부 참조",
                u"대괄호 [ ] 가 표시",
                u"2·4·6 번이 외부 참조입니다",
                u"3번은 같은 파일의 다른 시트, 5번은 INDIRECT 로 같은 파일 안을 가리킵니다."))

    ws = wb.create_sheet(u"4-08 유효성 검사")
    title(ws, u"4-08  데이터 유효성 검사",
          u"노란 칸에 유효성 검사를 직접 걸어 보고, 아래 물음에 답하세요.")
    widths(ws, {"A": 3, "B": 16, "C": 10, "D": 12, "E": 3, "F": 14, "G": 10, "H": 3, "I": 48})
    head(ws, 4, 2, [u"제품명", u"수량", u"금액"])
    for i in range(5):
        r = 5 + i
        put(ws, r, 2, None, fill=YELLOW)
        put(ws, r, 3, None, fill=YELLOW)
        put(ws, r, 4, None, fill=YELLOW, fmt=u"#,##0")
    put(ws, 4, 6, u"제품 목록", bold=True)
    for i, (nm, p) in enumerate([(u"진라면", 780), (u"오라면", 830), (u"진짬뽕", 920),
                                 (u"진짜장", 920), (u"김치라면", 850)]):
        put(ws, 5 + i, 6, nm); put(ws, 5 + i, 7, p, fmt=u"#,##0")
    for i, q in enumerate([u"① 제품명 칸에 걸 제한 대상은?",
                           u"② 수량 칸에 걸 제한 대상과 조건은?",
                           u"③ 유효성 검사가 못 막는 것은?"]):
        r = 12 + i
        put(ws, r, 2, q, align_left=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=4, value=None)
        c.fill = PatternFill("solid", fgColor=YELLOW)
        c.border = THIN
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
    guide_box(ws, 4, 9,
              u"데이터 → 데이터 유효성 검사.\n"
              u"목록의 원본은 =$F$5:$F$9 입니다.\n\n"
              u"「오류 메시지」 탭에서 문구를 직접 쓰세요.\n"
              u"기본 문구는 「입력한 값이 잘못되었습니다」 뿐이라\n"
              u"받는 사람이 뭘 어떻게 하라는 건지 모릅니다.")
    ans.append((u"4-08 유효성 검사",
                u"목록 · 정수 · 붙여넣기",
                u"① 목록 (원본 =$F$5:$F$9)  ② 정수, >= 0  ③ 복사해서 붙여넣기 "
                u"— 규칙까지 덮어써 없어집니다",
                u"다 채운 뒤 데이터 → 데이터 유효성 검사 → 잘못된 데이터 로 한 번 확인하세요."))

    ws = wb.create_sheet(u"4-12 인쇄 설정")
    title(ws, u"4-12  인쇄 설정",
          u"아래 상황마다 「어디를 건드려야 하는지」 노란 칸에 적으세요.")
    widths(ws, {"A": 3, "B": 48, "C": 34, "D": 3, "E": 48})
    head(ws, 4, 2, [u"상황", u"어떻게"])
    for i, q in enumerate([u"① 옆의 빈 칸까지 한 장 더 나온다",
                           u"② 둘째 장에 머리글이 없다",
                           u"③ 열이 많아 오른쪽이 잘린다",
                           u"④ 눈금선이 같이 찍힌다",
                           u"⑤ 줄이 늘어도 다시 안 건드리고 싶다",
                           u"⑥ 세 시트를 한 PDF 로 만들고 싶다"]):
        r = 5 + i
        put(ws, r, 2, q, align_left=True)
        put(ws, r, 3, None, fill=YELLOW)
    guide_box(ws, 4, 5,
              u"화면과 종이는 따로입니다.\n"
              u"틀 고정은 화면, 인쇄 제목은 종이입니다.\n"
              u"둘 다 해 주셔야 합니다.\n\n"
              u"인쇄 전에는 보기 → 페이지 나누기 미리 보기.\n"
              u"파란 선이 장이 갈리는 자리이고, 끌어 옮길 수 있습니다.")
    ans.append((u"4-12 인쇄 설정",
                u"영역 · 제목 · 배율 · 눈금선",
                u"① 인쇄 영역 설정  ② 인쇄 제목 → 반복할 행  ③ 가로 방향 또는 너비 1페이지  "
                u"④ 페이지 레이아웃 → 눈금선 → 인쇄 체크 해제  ⑤ 크기 조정을 너비 1페이지·높이 자동  "
                u"⑥ 시트를 함께 잡고 내보내기 → 「선택한 시트」",
                u"너비·높이를 둘 다 1페이지로 두면 줄이 늘수록 글자가 개미만 해집니다."))
    return ans


# ── 5장 문제 ──────────────────────────────────────────────────────
def chapter5(wb):
    ans = []

    ws = wb.create_sheet(u"5-01 자료의 모양")
    title(ws, u"5-01  데이터 관리 기본 규칙",
          u"아래 표가 왜 못 쓰는지 노란 칸에 적으세요.")
    widths(ws, {"A": 3, "B": 24, "C": 12, "D": 12, "E": 12, "F": 3, "G": 46})
    put(ws, 3, 2, u"영업팀 실적", bold=True)
    put(ws, 4, 3, u"1분기"); put(ws, 4, 5, u"2분기")
    head(ws, 5, 2, [u"팀\n(인원)", u"1월", u"2월", u"3월"])
    for i, (nm, a, b, c) in enumerate([(u"영업1팀\n(3명)", 102, 449, 477),
                                       (u"영업2팀\n(2명)", 379, 478, 378),
                                       (u"소계", 481, 927, 855),
                                       (u"영업3팀\n(4명)", 102, 454, 310),
                                       (u"소계", 102, 454, 310)]):
        r = 6 + i
        c0 = put(ws, r, 2, nm, wrap=True)
        put(ws, r, 3, a); put(ws, r, 4, b); put(ws, r, 5, c)
    for i, q in enumerate([u"① 팀 칸의 문제는?", u"② 머리글의 문제는?",
                           u"③ 표 안의 「소계」 줄이 왜 문제인가?",
                           u"④ 월이 가로로 간 것이 왜 문제인가?"]):
        r = 12 + i
        put(ws, r, 2, q, align_left=True)
        c = ws.cell(row=r, column=3, value=None)
        c.fill = PatternFill("solid", fgColor=YELLOW)
        c.border = THIN
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
    guide_box(ws, 4, 7,
              u"자료 표의 다섯 규칙 —\n"
              u"  한 칸에 한 값 / 머리글은 한 줄 /\n"
              u"  소계를 섞지 않기 / 늘어나는 쪽을 세로로 /\n"
              u"  표 위에 아무것도 없이")
    ans.append((u"5-01 자료의 모양",
                u"다섯 규칙",
                u"① 팀 이름과 인원이 한 칸에 뭉침  ② 1분기/2분기 때문에 머리글이 두 줄  "
                u"③ 총합계를 잡으면 두 배가 됨  ④ 달이 늘 때마다 열을 더해야 함",
                u"소계는 피벗이 만들게 두고 원본에는 자료만 둡니다."))

    ws = wb.create_sheet(u"5-03 정렬")
    title(ws, u"5-03  정렬",
          u"아래 표를 정렬해 보고, 물음에 답하세요.")
    widths(ws, {"A": 3, "B": 12, "C": 14, "D": 12, "E": 14, "F": 3, "G": 46})
    head(ws, 4, 2, [u"지역", u"대분류", u"담당자", u"금액"])
    rows = [(u"경기", u"사무용품", u"박수혁", 805100), (u"서울", u"가구", u"이강은", 24600),
            (u"경상", u"가구", u"박예설", 288200), (u"서울", u"사무용품", u"최민슬", 82700),
            (u"전라", u"가구", u"이새은", 1053300), (u"경기", u"가구", u"정지산", 16100)]
    for i, (a, b, c, d) in enumerate(rows):
        r = 5 + i
        put(ws, r, 2, a); put(ws, r, 3, b); put(ws, r, 4, c)
        put(ws, r, 5, d, fmt=u"#,##0")
    for i, q in enumerate([u"① 한 열만 잡고 정렬하면?",
                           u"② 지역→대분류→금액(큰 순) 으로 세우려면 기준을 몇 개?",
                           u"③ 서울·경기·경상·전라 순으로 세우려면?"]):
        r = 12 + i
        put(ws, r, 2, q, align_left=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=4, value=None)
        c.fill = PatternFill("solid", fgColor=YELLOW)
        c.border = THIN
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
    guide_box(ws, 4, 7,
              u"표 안 아무 칸이나 한 칸만 누르고 정렬하세요.\n"
              u"그러면 엑셀이 표 전체를 알아서 잡습니다.\n\n"
              u"사용자 지정 목록은\n파일 → 옵션 → 고급 → 맨 아래에서 등록합니다.")
    ans.append((u"5-03 정렬",
                u"표 안 한 칸만 누르기",
                u"① 그 열만 움직여 이름과 금액이 어긋남  ② 세 개 (위가 먼저)  "
                u"③ 사용자 지정 목록에 등록하고 정렬 기준에서 고름",
                u"「현재 선택 영역으로 정렬」을 누르면 표가 깨집니다. 확장을 고르거나, 한 칸만 누르세요."))

    ws = wb.create_sheet(u"5-08 SUBTOTAL")
    title(ws, u"5-08  필터와 합계",
          u"노란 칸에 수식을 넣고, 필터를 걸어 값이 바뀌는지 보세요.")
    widths(ws, {"A": 3, "B": 12, "C": 14, "D": 14, "E": 3, "F": 48})
    put(ws, 3, 2, u"SUM 합계", bold=True); put(ws, 3, 3, None, fill=YELLOW, fmt=u"#,##0")
    put(ws, 4, 2, u"SUBTOTAL 합계", bold=True); put(ws, 4, 3, None, fill=YELLOW, fmt=u"#,##0")
    head(ws, 6, 2, [u"지역", u"담당자", u"금액"])
    for i, (a, b, c) in enumerate([(u"서울", u"이강은", 24600), (u"경기", u"박수혁", 805100),
                                   (u"서울", u"최민슬", 82700), (u"경상", u"박예설", 288200),
                                   (u"전라", u"이새은", 1053300), (u"경기", u"정지산", 16100)]):
        r = 7 + i
        put(ws, r, 2, a); put(ws, r, 3, b); put(ws, r, 4, c, fmt=u"#,##0")
    for i, q in enumerate([u"① 서울만 필터했을 때 SUM 은?",
                           u"② 서울만 필터했을 때 SUBTOTAL 은?",
                           u"③ 9 와 109 의 차이는?"]):
        r = 15 + i
        put(ws, r, 2, q, align_left=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=4, value=None)
        c.fill = PatternFill("solid", fgColor=YELLOW)
        c.border = THIN
    guide_box(ws, 4, 6,
              u"C3 에 =SUM(D7:D12),\nC4 에 =SUBTOTAL(9,D7:D12) 를 넣으세요.\n\n"
              u"그리고 Ctrl+Shift+L 로 필터를 켜고\n지역에서 서울만 골라 보세요.\n"
              u"두 칸의 값이 어떻게 달라지는지 보시면 됩니다.")
    ans.append((u"5-08 SUBTOTAL",
                u"보이는 줄만 더하기",
                u"① 2,270,000 (그대로)  ② 107,300 (서울 두 줄)  "
                u"③ 9 는 필터로 숨긴 줄만, 109 는 손으로 숨긴 줄도 뺍니다",
                u"SUBTOTAL 은 SUBTOTAL 을 안 셉니다. 소계와 총합계를 같이 둬도 두 배가 안 됩니다."))

    ws = wb.create_sheet(u"5-10 고급 필터")
    title(ws, u"5-10  고급 필터 조건 만들기",
          u"요구사항을 조건표로 옮겨 적으세요.")
    widths(ws, {"A": 3, "B": 16, "C": 14, "D": 12, "E": 12, "F": 3, "G": 50})
    put(ws, 3, 2, u"요구사항", bold=True)
    req = ws.cell(row=4, column=2,
                  value=u"「사무용품이면서 수량 100 이상」 또는 「가구이면서 수량 80 이상」 인 것만 뽑아 주세요.")
    req.font = Font(name=FONT, size=11)
    ws.merge_cells("B4:E4")
    put(ws, 6, 2, u"조건표 — 노란 칸을 채우세요", bold=True)
    head(ws, 7, 2, [u"고객명", u"대분류", u"수량", u"매출"])
    for r in (8, 9):
        for c in range(2, 6):
            put(ws, r, c, None, fill=YELLOW)
    for i, q in enumerate([u"① 조건표 첫 줄에 들어갈 것은?",
                           u"② 가로로 나란히 적으면 무슨 뜻?",
                           u"③ 줄을 바꿔 적으면 무슨 뜻?",
                           u"④ 다른 시트로 뽑으려면?"]):
        r = 12 + i
        put(ws, r, 2, q, align_left=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=4, value=None)
        c.fill = PatternFill("solid", fgColor=YELLOW)
        c.border = THIN
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
    guide_box(ws, 4, 7,
              u"규칙은 딱 둘입니다 —\n"
              u"  · 첫 줄은 원본 머리글과 글자가 똑같아야 합니다\n"
              u"    (손으로 치지 말고 복사해서 붙이세요)\n"
              u"  · 가로는 그리고, 세로는 또는")
    ans.append((u"5-10 고급 필터",
                u"가로는 그리고, 세로는 또는",
                u"8행: (고객명 비움) 사무용품 >=100 / 9행: (비움) 가구 >=80.  "
                u"① 원본과 똑같은 머리글  ② AND  ③ OR  "
                u"④ 결과를 놓을 시트에서 먼저 「데이터 → 고급」 을 누름",
                u"머리글이 한 글자라도 다르면 안 잡힙니다. 빈 줄을 조건 범위에 넣어도 안 됩니다."))
    return ans


# ── 6장 문제 ──────────────────────────────────────────────────────
def chapter6(wb):
    ans = []

    ws = wb.create_sheet(u"6-01 엑셀 표")
    title(ws, u"6-01  엑셀 표",
          u"표로 바꾸고 노란 칸에 답을 적으세요.")
    widths(ws, {"A": 3, "B": 16, "C": 14, "D": 14, "E": 3, "F": 48})
    put(ws, 3, 2, u"합계", bold=True); put(ws, 3, 4, None, fill=YELLOW, fmt=u"#,##0")
    head(ws, 5, 2, [u"날짜", u"배달업체", u"사용자수"])
    for i, (d, nm, v) in enumerate([(u"2021-05-12", u"여기요", 351),
                                    (u"2021-05-13", u"배달의혈통", 335),
                                    (u"2021-05-14", u"후팡잇츠", 393),
                                    (u"2021-05-15", u"신속배달", 446)]):
        r = 6 + i
        put(ws, r, 2, d); put(ws, r, 3, nm); put(ws, r, 4, v)
    for i, q in enumerate([u"① 표로 바꾸는 단축키는?",
                           u"② 표 이름을 「일별사용자수」로 바꾼 뒤 합계 수식은?",
                           u"③ 표 안에서 아예 안 되는 것은?",
                           u"④ 표를 되돌리려면?"]):
        r = 12 + i
        put(ws, r, 2, q, align_left=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=4, value=None)
        c.fill = PatternFill("solid", fgColor=YELLOW)
        c.border = THIN
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
    guide_box(ws, 4, 6,
              u"표로 바꾼 뒤 맨 아래 칸에서 Tab 을 눌러\n"
              u"줄을 하나 더해 보세요.\n"
              u"합계가 따라오는지 보시면 됩니다.\n\n"
              u"표 이름은 「표 디자인」 탭 왼쪽 끝에서 바꿉니다.")
    ans.append((u"6-01 엑셀 표",
                u"Ctrl+T · 구조적 참조",
                u"① Ctrl+T  ② =SUM(일별사용자수[사용자수])  ③ 셀 병합  ④ 표 디자인 → 범위로 변환",
                u"표는 범위가 저절로 늘어납니다. 수식·피벗·차트가 전부 따라옵니다."))

    ws = wb.create_sheet(u"6-03 자료 형태")
    title(ws, u"6-03  피벗을 위한 데이터 형태",
          u"왼쪽 표를 오른쪽에 세로로 옮겨 적으세요.")
    widths(ws, {"A": 3, "B": 12, "C": 10, "D": 10, "E": 10, "F": 3,
                "G": 12, "H": 10, "I": 12, "J": 3, "K": 46})
    head(ws, 4, 2, [u"지점", u"2021", u"2022", u"2023"])
    src = [(u"강남점", 30000, 8000, 26000), (u"명동점", 37000, 6000, 25000),
           (u"서초점", 33000, 11000, 33000)]
    for i, (nm, a, b, c) in enumerate(src):
        r = 5 + i
        put(ws, r, 2, nm); put(ws, r, 3, a); put(ws, r, 4, b); put(ws, r, 5, c)
    head(ws, 4, 7, [u"지점", u"년도", u"매출액"])
    for i in range(9):
        r = 5 + i
        for c in (7, 8, 9):
            put(ws, r, c, None, fill=YELLOW)
    for i, q in enumerate([u"① 옮기면 몇 줄이 되나?",
                           u"② 왼쪽 표로 피벗을 만들면 무엇이 없나?",
                           u"③ 손으로 안 옮기고 하려면?"]):
        r = 16 + i
        put(ws, r, 2, q, align_left=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        c = ws.cell(row=r, column=5, value=None)
        c.fill = PatternFill("solid", fgColor=YELLOW)
        c.border = THIN
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=9)
    guide_box(ws, 4, 11,
              u"한 줄 = 한 사건.\n"
              u"「강남점의 2021년 매출은 30000」 이 한 줄입니다.\n\n"
              u"3지점 × 3년이니 몇 줄이 될지 먼저 세어 보세요.")
    ans.append((u"6-03 자료 형태",
                u"한 줄 = 한 사건",
                u"① 9줄 (3지점 × 3년)  ② 「년도」라는 필드가 없어 년도별로 못 나눔  "
                u"③ 파워 쿼리 → 지점 열만 고르고 「다른 열 피벗 해제」",
                u"가로로도 피벗이 만들어지기는 합니다. 그래서 더 위험합니다."))

    ws = wb.create_sheet(u"6-04 피벗 네 칸")
    title(ws, u"6-04  피벗 테이블 네 칸",
          u"요구마다 어느 칸에 무엇을 넣는지 적으세요.")
    widths(ws, {"A": 3, "B": 52, "C": 14, "D": 14, "E": 14, "F": 14, "G": 3, "H": 44})
    head(ws, 4, 2, [u"이렇게 보고 싶다", u"행", u"열", u"값", u"필터"])
    for i, q in enumerate([u"① 금액을 지점별 제조사별로",
                           u"② 수량을 월별로, 무료배송 여부는 골라 가며",
                           u"③ 손님 수(영수증 고유 개수)를 브랜드별로"]):
        r = 5 + i
        put(ws, r, 2, q, align_left=True)
        for c in (3, 4, 5, 6):
            put(ws, r, c, None, fill=YELLOW)
    for i, q in enumerate([u"④ 금액을 넣었는데 「개수 : 금액」 이 나오는 이유는?",
                           u"⑤ 고유 개수를 쓰려면 만들 때 무엇을 체크하나?"]):
        r = 10 + i
        put(ws, r, 2, q, align_left=True)
        c = ws.cell(row=r, column=3, value=None)
        c.fill = PatternFill("solid", fgColor=YELLOW)
        c.border = THIN
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    guide_box(ws, 4, 8,
              u"한 줄로 외우면 —\n"
              u"「무엇을(값) 무엇별로(행·열) 보겠다」.\n\n"
              u"위에서 골라 볼 것은 필터 칸입니다.")
    ans.append((u"6-04 피벗 네 칸",
                u"값 / 행 / 열 / 필터",
                u"① 행=지점, 열=제조사, 값=금액  ② 행=월, 값=수량, 필터=무료배송  "
                u"③ 행=브랜드, 값=영수증번호(고유 개수)  ④ 그 열에 글자가 섞여 있음  "
                u"⑤ 「데이터 모델에 이 데이터 추가」",
                u"이미 만든 피벗은 데이터 모델로 못 바꿉니다. 새로 만들어야 합니다."))

    ws = wb.create_sheet(u"6-11 계산 필드")
    title(ws, u"6-11  계산 필드와 계산 항목",
          u"아래 값으로 이익률을 두 방식으로 내 보고, 차이를 적으세요.")
    widths(ws, {"A": 3, "B": 14, "C": 14, "D": 14, "E": 16, "F": 3, "G": 50})
    head(ws, 4, 2, [u"년도", u"매출액", u"매출이익", u"줄별 이익률"])
    vals = [(2021, 30000, 1488), (2022, 8000, 1357), (2023, 26000, 1575)]
    tot_sales = sum(v[1] for v in vals)
    tot_profit = sum(v[2] for v in vals)
    for i, (y, s, p) in enumerate(vals):
        r = 5 + i
        put(ws, r, 2, y); put(ws, r, 3, s, fmt=u"#,##0"); put(ws, r, 4, p, fmt=u"#,##0")
        put(ws, r, 5, None, fill=YELLOW, fmt=u"0.0%")
    put(ws, 9, 2, u"① 계산 필드 방식 (합계끼리 나눔)", align_left=True)
    ws.merge_cells("B9:D9")
    put(ws, 9, 5, None, fill=YELLOW, fmt=u"0.0%")
    put(ws, 10, 2, u"② 계산 항목 방식 (줄별 이익률을 더함)", align_left=True)
    ws.merge_cells("B10:D10")
    put(ws, 10, 5, None, fill=YELLOW, fmt=u"0.0%")
    put(ws, 11, 2, u"③ 어느 쪽이 맞나?", align_left=True)
    ws.merge_cells("B11:D11")
    put(ws, 11, 5, None, fill=YELLOW)
    guide_box(ws, 4, 7,
              u"계산 필드는 합계를 먼저 내고 나눕니다.\n"
              u"계산 항목은 줄마다 나눈 값을 더합니다.\n\n"
              u"둘이 얼마나 다른지 직접 계산해 보세요.\n"
              u"매출 합계 %s · 이익 합계 %s" % (format(tot_sales, ","), format(tot_profit, ",")))
    ans.append((u"6-11 계산 필드",
                u"비율은 반드시 계산 필드로",
                u"줄별: %.1f%% · %.1f%% · %.1f%%   ① %.1f%% (맞음)   ② %.1f%% (틀림)   ③ ①"
                % (vals[0][2] / vals[0][1] * 100, vals[1][2] / vals[1][1] * 100,
                   vals[2][2] / vals[2][1] * 100,
                   tot_profit / tot_sales * 100,
                   sum(p / s for _, s, p in vals) * 100),
                u"비율의 합은 아무 뜻이 없습니다. 그런데 피벗은 그 값을 자신 있게 내놓습니다."))
    return ans


# ── 7장 문제 ──────────────────────────────────────────────────────
def chapter7(wb):
    ans = []

    ws = wb.create_sheet(u"7-03 IF")
    title(ws, u"7-03  IF 와 중첩 IF",
          u"노란 칸에 수식을 넣으세요.")
    widths(ws, {"A": 3, "B": 12, "C": 10, "D": 10, "E": 10, "F": 12, "G": 10, "H": 3, "I": 46})
    head(ws, 4, 2, [u"학생명", u"듣기", u"쓰기", u"평균", u"합격여부", u"등급"])
    data = [(u"김세민", 92, 88), (u"정다온", 55, 91), (u"김진선", 78, 82),
            (u"정희엘", 64, 59), (u"박단비", 71, 73)]
    for i, (nm, a, b) in enumerate(data):
        r = 5 + i
        put(ws, r, 2, nm); put(ws, r, 3, a); put(ws, r, 4, b)
        put(ws, r, 5, None, fill=YELLOW, fmt=u"0.0")
        put(ws, r, 6, None, fill=YELLOW)
        put(ws, r, 7, None, fill=YELLOW)
    guide_box(ws, 4, 9,
              u"합격: 듣기와 쓰기가 모두 60 이상\n"
              u"등급: 90↑ A · 80↑ B · 70↑ C · 나머지 D\n\n"
              u"중첩 IF 는 큰 것부터 씁니다.\n"
              u"거꾸로 쓰면 95점도 C 가 됩니다.")
    ans.append((u"7-03 IF",
                u"AND · 중첩은 큰 것부터",
                u"E5 = AVERAGE(C5:D5) / F5 = IF(AND(C5>=60,D5>=60),\"합격\",\"불합격\") / "
                u"G5 = IF(E5>=90,\"A\",IF(E5>=80,\"B\",IF(E5>=70,\"C\",\"D\")))",
                u"정다온·정희엘이 불합격입니다. 한 과목만 60 미만이어도 불합격입니다."))

    ws = wb.create_sheet(u"7-05 SUMIFS")
    title(ws, u"7-05  SUMIF 와 SUMIFS",
          u"노란 칸에 수식을 넣으세요.")
    widths(ws, {"A": 3, "B": 12, "C": 14, "D": 12, "E": 3, "F": 14, "G": 14, "H": 3, "I": 46})
    head(ws, 4, 2, [u"지역", u"대분류", u"금액"])
    rows = [(u"서울", u"가구", 24600), (u"경기", u"사무용품", 805100),
            (u"서울", u"사무용품", 82700), (u"경상", u"가구", 288200),
            (u"서울", u"가구", 153300), (u"경기", u"가구", 16100)]
    for i, (a, b, c) in enumerate(rows):
        r = 5 + i
        put(ws, r, 2, a); put(ws, r, 3, b); put(ws, r, 4, c, fmt=u"#,##0")
    put(ws, 4, 6, u"서울 합계", bold=True)
    put(ws, 4, 7, None, fill=YELLOW, fmt=u"#,##0")
    put(ws, 5, 6, u"서울·가구 합계", bold=True)
    put(ws, 5, 7, None, fill=YELLOW, fmt=u"#,##0")
    put(ws, 6, 6, u"10만 이상 합계", bold=True)
    put(ws, 6, 7, None, fill=YELLOW, fmt=u"#,##0")
    guide_box(ws, 4, 9,
              u"SUMIF  는 더할 범위가 맨 뒤,\n"
              u"SUMIFS 는 더할 범위가 맨 앞입니다.\n\n"
              u"헷갈리면 SUMIFS 만 쓰세요.\n조건이 하나여도 됩니다.\n\n"
              u"부등호는 따옴표 안에 — \">=100000\"")
    seoul = sum(c for a, b, c in rows if a == u"서울")
    sg = sum(c for a, b, c in rows if a == u"서울" and b == u"가구")
    big = sum(c for a, b, c in rows if c >= 100000)
    ans.append((u"7-05 SUMIFS",
                u"SUMIFS 만 쓰면 안 헷갈립니다",
                u"=SUMIFS($D$5:$D$10,$B$5:$B$10,\"서울\") → %s / "
                u"조건 둘 → %s / =SUMIFS($D$5:$D$10,$D$5:$D$10,\">=100000\") → %s"
                % (format(seoul, ","), format(sg, ","), format(big, ",")),
                u"\">=K5\" 라고 쓰면 K5 라는 글자와 견주게 되어 0 이 나옵니다. \">=\"&K5 로 써야 합니다."))

    ws = wb.create_sheet(u"7-07 TRIM")
    title(ws, u"7-07  LEN 과 TRIM",
          u"왜 VLOOKUP 이 안 되는지 찾아 고치세요.")
    widths(ws, {"A": 3, "B": 20, "C": 14, "D": 10, "E": 16, "F": 3,
                "G": 14, "H": 14, "I": 3, "J": 46})
    head(ws, 4, 2, [u"지역", u"인구수", u"글자 수", u"공백 뗀 것"])
    src = [(u"　　　종로구", 151291), (u"용산구", 217194), (u"　　　성동구", 282550),
           (u"광진구", 341417), (u"　　　동대문구", 337215)]
    for i, (nm, v) in enumerate(src):
        r = 5 + i
        put(ws, r, 2, nm, align_left=True); put(ws, r, 3, v, fmt=u"#,##0")
        put(ws, r, 4, None, fill=YELLOW)
        put(ws, r, 5, None, fill=YELLOW)
    put(ws, 4, 7, u"찾을 지역", bold=True); put(ws, 4, 8, u"인구수", bold=True)
    for i, nm in enumerate([u"종로구", u"성동구", u"동대문구"]):
        put(ws, 5 + i, 7, nm)
        put(ws, 5 + i, 8, None, fill=YELLOW, fmt=u"#,##0")
    guide_box(ws, 4, 10,
              u"왼쪽 지역 이름 앞에 보이지 않는 공백이 있습니다.\n"
              u"D열에 =LEN(B5) 를 넣어 글자 수를 세어 보세요.\n\n"
              u"E열에 =TRIM(B5) 를 넣어 떼어낸 뒤,\n"
              u"H열 VLOOKUP 이 E열을 보게 하면 됩니다.")
    ans.append((u"7-07 TRIM",
                u"LEN 으로 재고 TRIM 으로 뗌",
                u"D5 = LEN(B5) → 6 (종로구는 3이어야 함) / E5 = TRIM(B5) / "
                u"H5 = VLOOKUP(G5,$E$5:$C$9…) → TRIM 한 열을 찾을 열로 두고 INDEX/MATCH 가 편합니다",
                u"웹에서 복사한 자료에는 CHAR(160) 이 섞여 TRIM 으로 안 떨어집니다. "
                u"=TRIM(SUBSTITUTE(B5,CHAR(160),\" \")) 로 먼저 바꾸세요."))

    ws = wb.create_sheet(u"7-12 날짜")
    title(ws, u"7-12  날짜 되살리기와 DATEDIF",
          u"글자로 된 날짜를 살리고 기간을 재세요.")
    widths(ws, {"A": 3, "B": 18, "C": 12, "D": 14, "E": 12, "F": 12, "G": 3, "H": 46})
    put(ws, 3, 2, u"기준일", bold=True)
    put(ws, 3, 3, u"2026-09-16")
    head(ws, 5, 2, [u"등록코드", u"날짜코드", u"진짜 날짜", u"구입년도", u"사용개월"])
    for i, code in enumerate([u"KM-A03-210415", u"KM-B01-220930", u"JD-C02-231201",
                              u"FQ-A07-240718"]):
        r = 6 + i
        put(ws, r, 2, code)
        put(ws, r, 3, None, fill=YELLOW)
        put(ws, r, 4, None, fill=YELLOW, fmt=u"yyyy-mm-dd")
        put(ws, r, 5, None, fill=YELLOW)
        put(ws, r, 6, None, fill=YELLOW)
    guide_box(ws, 4, 8,
              u"코드 맨 뒤 여섯 자리가 YYMMDD 입니다.\n\n"
              u"RIGHT 로 여섯 자리를 뽑고,\n"
              u"DATE(LEFT()+2000, MID(), RIGHT()) 로 날짜를 만듭니다.\n\n"
              u"DATEDIF 는 목록에 안 뜹니다. 손으로 전부 치세요.")
    ans.append((u"7-12 날짜",
                u"DATE 로 만들고 DATEDIF 로 잼",
                u"C6 = RIGHT(B6,6) / D6 = DATE(LEFT(C6,2)+2000,MID(C6,3,2),RIGHT(C6,2)) / "
                u"E6 = YEAR(D6) / F6 = DATEDIF(D6,$C$3,\"M\")",
                u"+2000 을 빼먹으면 서기 21년이 됩니다. 기준일은 TODAY() 대신 칸 하나에 적어 두세요."))
    return ans


# ── 8장 문제 ──────────────────────────────────────────────────────
def chapter8(wb):
    ans = []

    ws = wb.create_sheet(u"8-01 차트 고르기")
    title(ws, u"8-01  어느 차트를 고르나",
          u"하려는 말에 맞는 차트를 노란 칸에 적으세요.")
    widths(ws, {"A": 3, "B": 52, "C": 22, "D": 3, "E": 48})
    head(ws, 4, 2, [u"하려는 말", u"고를 차트"])
    for i, q in enumerate([u"① 3년간 매출이 어떻게 변했나",
                           u"② 17개 시도 중 어디가 많은가",
                           u"③ 상위 4개 지점만 견주고 싶다",
                           u"④ 전체가 얼마이고 안이 어떻게 나뉘나",
                           u"⑤ 매출(억)과 이익률(%)을 함께",
                           u"⑥ 일정이 언제부터 언제까지인가"]):
        r = 5 + i
        put(ws, r, 2, q, align_left=True)
        put(ws, r, 3, None, fill=YELLOW)
    guide_box(ws, 4, 5,
              u"자료가 정하는 것이 아니라\n하려는 말이 정합니다.\n\n"
              u"항목이 많거나 이름이 길면 가로 막대입니다.\n"
              u"단위가 다르면 혼합 + 보조 축입니다.")
    ans.append((u"8-01 차트 고르기",
                u"하려는 말이 차트를 정합니다",
                u"① 꺾은선  ② 가로 막대  ③ 세로 막대  ④ 누적 막대  ⑤ 혼합 + 보조 축  ⑥ 간트",
                u"원형은 조각이 다섯을 넘으면 아무도 못 읽습니다. 가로 막대로 바꾸세요."))

    ws = wb.create_sheet(u"8-02 축")
    title(ws, u"8-02  축과 눈금",
          u"각 상황에서 무엇이 잘못됐는지 적으세요.")
    widths(ws, {"A": 3, "B": 54, "C": 26, "D": 3, "E": 48})
    head(ws, 4, 2, [u"상황", u"무엇이 잘못됐나"])
    for i, q in enumerate([u"① 막대 차트의 세로 축이 250부터 시작한다",
                           u"② 가로 막대에서 1등이 맨 아래에 있다",
                           u"③ 보조 축 눈금을 보기 좋게 맞췄다",
                           u"④ 4분기 예상치가 실선으로 이어져 있다",
                           u"⑤ 3차원 막대로 그렸다"]):
        r = 5 + i
        put(ws, r, 2, q, align_left=True)
        put(ws, r, 3, None, fill=YELLOW)
    guide_box(ws, 4, 5,
              u"차트는 일부러 하지 않아도\n저절로 거짓말을 하게 됩니다.\n\n"
              u"그래서 만든 뒤에\n한 번 확인하는 습관이 필요합니다.")
    ans.append((u"8-02 축",
                u"막대는 0부터",
                u"① 작은 차이가 거대해 보임  ② 축 서식 → 「항목을 거꾸로」 를 안 함  "
                u"③ 상관없는 둘이 같이 움직이는 것처럼 보임  ④ 안 일어난 일이 일어난 것처럼 보임 (점선으로)  "
                u"⑤ 원근 때문에 길이가 왜곡됨",
                u"꺾은선은 0부터 안 해도 됩니다. 거기서는 기울기가 값이니까요."))

    ws = wb.create_sheet(u"8-04 혼합")
    title(ws, u"8-04  혼합 차트와 보조 축",
          u"아래 자료로 혼합 차트를 만들고 물음에 답하세요.")
    widths(ws, {"A": 3, "B": 12, "C": 16, "D": 14, "E": 3, "F": 48})
    head(ws, 4, 2, [u"년도", u"GDP(십억)", u"성장률"])
    for i, (y, g, r2) in enumerate([(2021, 2071658, 0.0410), (2022, 2161774, 0.0262),
                                    (2023, 2236330, 0.0136), (2024, 2305271, 0.0208)]):
        r = 5 + i
        put(ws, r, 2, y); put(ws, r, 3, g, fmt=u"#,##0")
        put(ws, r, 4, r2, fmt=u"0.0%")
    for i, q in enumerate([u"① 그냥 그리면 성장률이 왜 안 보이나?",
                           u"② 어느 쪽을 막대로 하나?",
                           u"③ 보조 축을 쓸 때 꼭 해야 할 것은?"]):
        r = 11 + i
        put(ws, r, 2, q, align_left=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=4, value=None)
        c.fill = PatternFill("solid", fgColor=YELLOW)
        c.border = THIN
    guide_box(ws, 4, 6,
              u"차트 디자인 → 차트 종류 변경 → 맨 아래 「콤보」.\n"
              u"계열마다 종류를 정하고\n성장률 쪽에 「보조 축」을 체크합니다.")
    ans.append((u"8-04 혼합",
                u"막대와 선으로 축을 구별",
                u"① 200만과 0.04 라 한 축에서는 선이 바닥에 깔림  "
                u"② 크기를 말하는 GDP 가 막대, 흐름을 말하는 성장률이 선  "
                u"③ 축 두 개에 단위를 적기 (+ → 축 제목)",
                u"둘 다 막대로 하면 어느 축이 어느 막대인지 알 수 없습니다."))

    ws = wb.create_sheet(u"8-05 간트")
    title(ws, u"8-05  간트 차트",
          u"만드는 순서를 적으세요.")
    widths(ws, {"A": 3, "B": 16, "C": 14, "D": 10, "E": 14, "F": 3, "G": 50})
    head(ws, 4, 2, [u"일정명", u"시작일", u"소요일", u"마감일"])
    for i, (nm, d, n) in enumerate([(u"기획", u"2026-10-05", 10), (u"설계", u"2026-10-15", 14),
                                    (u"개발", u"2026-10-29", 21), (u"검수", u"2026-11-19", 7)]):
        r = 5 + i
        put(ws, r, 2, nm); put(ws, r, 3, d); put(ws, r, 4, n)
        put(ws, r, 5, None, fill=YELLOW)
    for i, q in enumerate([u"① 어떤 차트에서 출발하나?",
                           u"② 첫 막대(시작일)를 어떻게 하나?",
                           u"③ 세로 축에 무엇을 해야 하나?",
                           u"④ 가로 축 최소값에 무엇을 넣나?"]):
        r = 11 + i
        put(ws, r, 2, q, align_left=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=4, value=None)
        c.fill = PatternFill("solid", fgColor=YELLOW)
        c.border = THIN
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
    guide_box(ws, 4, 7,
              u"엑셀에 간트 차트라는 종류는 없습니다.\n"
              u"누적 가로 막대에서 앞 막대를 투명하게 만든 것입니다.\n\n"
              u"마감일 = 시작일 + 소요일 입니다.")
    ans.append((u"8-05 간트",
                u"누적 가로 막대 + 투명",
                u"① 누적 가로 막대  ② 채우기 없음 · 선 없음  ③ 축 서식 → 「항목을 거꾸로」  "
                u"④ 날짜 일련번호 (버전에 따라 날짜를 못 받습니다)",
                u"일정이 많으면 조건부 서식 수식 규칙으로 칸을 칠하는 편이 더 간단합니다."))
    return ans


def sheet_answer(wb, ans):
    ws = wb.create_sheet(u"정답")
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 3, "B": 18, "C": 34, "D": 44, "E": 52})
    t = ws.cell(row=1, column=2, value=u"정답")
    t.font = Font(name=FONT, size=16, bold=True, color=INK_HEAD)
    s = ws.cell(row=2, column=2, value=u"먼저 다 풀고 나서 보세요. 답이 달라도 결과가 같으면 맞은 것입니다.")
    s.font = Font(name=FONT, size=10, color=GREY)
    head(ws, 4, 2, [u"시트", u"핵심", u"답", u"왜 그런가"])
    for i, (a, b, c, d) in enumerate(ans):
        r = 5 + i
        for j, v in enumerate((a, b, c, d)):
            cell = put(ws, r, 2 + j, v, wrap=True)
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        ws.row_dimensions[r].height = 46
    return ws


BOOKS = [
    (1, u"1장_시작부터_남다른_실무자의_엑셀_활용", u"1장_복습.xlsx", chapter1),
    (2, u"2장_실무자라면_반드시_알아야_할_엑셀_활용", u"2장_복습.xlsx", chapter2),
    (3, u"3장_보고서가_달라지는_서식_활용법", u"3장_복습.xlsx", chapter3),
    (4, u"4장_완성한_보고서_공유_및_출력하기", u"4장_복습.xlsx", chapter4),
    (5, u"5장_데이터_정리부터_데이터_필터링까지", u"5장_복습.xlsx", chapter5),
    (6, u"6장_자동화를_위한_표와_피벗_테이블", u"6장_복습.xlsx", chapter6),
    (7, u"7장_기본과_필수_함수_익히기", u"7장_복습.xlsx", chapter7),
    (8, u"8장_실무에서_필요한_데이터_시각화", u"8장_복습.xlsx", chapter8),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args()

    for chap, folder, name, build in BOOKS:
        wb = Workbook()
        wb.remove(wb.active)
        sheet_guide(wb, chap)
        ans = build(wb)
        sheet_answer(wb, ans)
        path = os.path.join(OUT, folder, name)
        print(u"%s" % path)
        print(u"   시트 %d장 — %s" % (len(wb.sheetnames), u" · ".join(wb.sheetnames)))
        if a.write:
            wb.save(path)
    if not a.write:
        print(u"\n※ 보기만 했습니다. 실제로 만들려면 --write 를 붙이세요.")
    else:
        print(u"\n만들었습니다.")


if __name__ == "__main__":
    main()
