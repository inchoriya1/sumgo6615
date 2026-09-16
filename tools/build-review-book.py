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


def put(ws, row, col, value, fill=None, bold=False, fmt=None, wrap=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name=FONT, size=11, bold=bold)
    c.border = THIN
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
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
