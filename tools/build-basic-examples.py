# -*- coding: utf-8 -*-
u"""
엑셀 **기초 10장** 전용 예제 파일을 만듭니다.

  python tools/build-basic-examples.py            # 보기만
  python tools/build-basic-examples.py --write    # 실제로 만들기

왜 따로 만드나
  마스터 클래스 예제(`강의예제/N장_…`)는 **실무 자료가 그대로** 들어 있어
  줄이 많고 열이 넓습니다. 처음 켜보는 분에게는 그 자체가 벽입니다.
  기초 예제는 **한 화면에 다 보이는 크기**로, 한 장에 하나씩만 연습하도록 새로 만듭니다.

이름 규칙
  `B-02_입력_연습.xlsx` 처럼 **B-장번호** 를 앞에 붙입니다.
  (마스터는 `1-05_…`, 부록은 `A-01_…`, 기초는 `B-02_…`)
  1장은 빈 통합 문서만 쓰므로 예제가 없습니다. 그래서 B-02 부터입니다.

  답이 파일에 남는 것은 `_완성.xlsx` 를 함께 만듭니다.
  5장(저장)은 결과가 **새 파일**로 나오므로 완성본이 없습니다.
"""
import argparse
import os

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

DST = os.path.join(u"강의예제", u"기초_예제")

HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(color="FFFFFF", bold=True)
TODO     = PatternFill("solid", fgColor="FFE699")   # 채울 곳(노랑)
HINT     = PatternFill("solid", fgColor="FFF9E6")   # 설명(연노랑)
_thin    = Side(style="thin", color="BFBFBF")
BOX      = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
CTR      = Alignment(horizontal="center", vertical="center")
TOPWRAP  = Alignment(vertical="top", wrap_text=True)


def head(ws, row, col, names, width=None):
    u"""머리글 한 줄을 깝니다."""
    for i, n in enumerate(names):
        c = ws.cell(row, col + i, n)
        c.fill, c.font, c.alignment, c.border = HDR_FILL, HDR_FONT, CTR, BOX
    if width:
        for i, w in enumerate(width):
            ws.column_dimensions[ws.cell(row, col + i).column_letter].width = w


def body(ws, row, col, rows, todo_cols=(), fmt=None, done=False):
    u"""자료 줄을 깝니다.

    todo_cols 에 든 열은 **연습본에서는 비우고, 완성본에서는 값을 채웁니다.**
    어느 쪽이든 노란 칠은 그대로 둡니다 — 완성본을 열었을 때
    「내가 채울 곳이 여기였구나」 가 한눈에 보여야 하기 때문입니다.
    """
    for r, vals in enumerate(rows):
        for i, v in enumerate(vals):
            c = ws.cell(row + r, col + i)
            if i in todo_cols:
                c.fill = TODO
                if done:
                    c.value = v
            else:
                c.value = v
            c.border = BOX
            if fmt and i in fmt:
                c.number_format = fmt[i]
            if isinstance(v, (int, float)):
                c.alignment = Alignment(horizontal="right")


def title(ws, text, sub=None):
    ws["B2"] = text
    ws["B2"].font = Font(bold=True, size=14)
    if sub:
        ws["B3"] = sub
        ws["B3"].font = Font(size=10, color="666666")


def hint(ws, anchor, text, span="E:I", rows=6):
    c = ws[anchor]
    c.value = text
    c.fill, c.alignment, c.font = HINT, TOPWRAP, Font(size=10)
    a, b = span.split(":")
    r = int("".join(ch for ch in anchor if ch.isdigit()))
    ws.merge_cells("%s%d:%s%d" % (a, r, b, r + rows - 1))


# ────────────────────────────── 2장 · 입력하기 ──────────────────────────────
def b02(done):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = u"따라 치기"
    title(ws, u"과일 가게 하루 장부", u"머리글만 있습니다. 아래 네 줄을 직접 쳐 보세요.")
    head(ws, 5, 2, [u"품목", u"수량", u"단가", u"금액"], [12, 8, 10, 12])
    rows = [[u"사과", 3, 1000, None], [u"배", 2, 2000, None],
            [u"감", 5, 800, None], [u"귤", 10, 500, None]]
    # 금액 칸(맨 오른쪽)은 7장에서 수식으로 채웁니다 — 여기서는 노랗게 칠하지 않습니다
    body(ws, 6, 2, rows, todo_cols=(0, 1, 2), fmt={1: "#,##0", 2: "#,##0"}, done=done)
    hint(ws, "F5", u"· 숫자에는 쉼표를 직접 치지 마세요 — 1,000 이 아니라 1000\n"
                   u"· 다 치면 숫자는 오른쪽, 글자는 왼쪽에 붙습니다\n"
                   u"· 금액 칸은 비워 두세요. 7장에서 계산으로 채웁니다", "F:J")

    ws = wb.create_sheet(u"숫자인가 글자인가")
    title(ws, u"어느 쪽이 진짜 숫자일까요", u"오른쪽에 붙으면 숫자, 왼쪽에 붙으면 글자입니다.")
    head(ws, 5, 2, [u"보이는 것", u"내 판단", u"정답"], [16, 12, 12])
    pairs = [(1000, u"숫자"), (u"2000", u"글자"), (3000, u"숫자"),
             (u"4,000", u"글자"), (5000, u"숫자"), (u" 6000", u"글자")]
    for r, (v, ans) in enumerate(pairs):
        c = ws.cell(6 + r, 2, v); c.border = BOX
        if isinstance(v, str):
            c.number_format = "@"
        ws.cell(6 + r, 3).fill = TODO; ws.cell(6 + r, 3).border = BOX
        ws.cell(6 + r, 4).border = BOX
        if done:
            ws.cell(6 + r, 3, ans).alignment = CTR
            ws.cell(6 + r, 4, ans).alignment = CTR
    hint(ws, "F5", u"· 글자로 들어간 숫자는 더해지지 않습니다\n"
                   u"· 쉼표를 직접 치거나, 앞에 빈칸이 붙으면 글자가 됩니다\n"
                   u"· 왼쪽 위 초록 삼각형이 뜨면 「글자로 된 숫자」라는 뜻입니다", "F:J")

    ws = wb.create_sheet(u"날짜와 전화번호")
    title(ws, u"조심해야 하는 두 가지")
    head(ws, 5, 2, [u"치려는 것", u"그냥 치면", u"어떻게 해야 하나"], [18, 16, 30])
    rows = [[u"2026년 3월 5일", u"2026-3-5", u"그냥 치면 됩니다. 날짜로 들어갑니다"],
            [u"010-1234-5678", u"010-1234-5678", u"그냥 치면 됩니다. - 가 있어 글자로 남습니다"],
            [u"01012345678", u"1012345678", u"앞의 0 이 날아갑니다 → 칸을 「텍스트」로 먼저 바꾸세요"],
            [u"3-5 (제품 번호)", u"3월 5일", u"날짜로 바뀝니다 → 칸을 「텍스트」로 먼저 바꾸세요"]]
    body(ws, 6, 2, rows)
    for r in range(6, 10):
        ws.cell(r, 4).alignment = TOPWRAP
    return wb


# ─────────────────────────── 3장 · 고치고 되돌리기 ───────────────────────────
def b03(done):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = u"오타 고치기"
    title(ws, u"동호회 회원 명단", u"틀린 곳이 넷 있습니다. 일부 글자만 고쳐 보세요.")
    head(ws, 5, 2, [u"이름", u"부서", u"연락처"], [12, 14, 16])
    bad = [[u"김세민", u"영어팀", u"010-2540-7960"],
           [u"박정화", u"개발팀", u"010-5530-4901"],
           [u"이서우", u"생상팀", u"010-6603-6733"],
           [u"정다온", u"총무팀", u"010-2921-4409"],
           [u"김준용", u"인사팁", u"010-4874-3255"],
           [u"최효윤", u"기획팀", u"010-3596-2881"],
           [u"박단비", u"IT팀",  u"010-3533-9391"]]
    good = [r[:] for r in bad]
    good[0][1] = u"영업팀"; good[2][1] = u"생산팀"
    good[4][1] = u"인사팀"; good[6][2] = u"010-3533-9381"
    body(ws, 6, 2, good if done else bad)
    hint(ws, "F5", u"· 통째로 바꿀 때는 그냥 다시 칩니다\n"
                   u"· 한 글자만 고칠 때는 F2 를 눌러 칸 안으로 들어갑니다\n"
                   u"· 잘못 고쳤으면 Ctrl+Z — 여러 번 눌러도 됩니다\n"
                   u"· 틀린 곳: 영어팀 · 생상팀 · 인사팁 · 마지막 연락처 끝자리", "F:J", 7)

    ws = wb.create_sheet(u"지우기 두 가지")
    title(ws, u"Delete 와 Backspace 는 다릅니다")
    head(ws, 5, 2, [u"누르면", u"이렇게 됩니다"], [16, 40])
    body(ws, 6, 2, [[u"Delete", u"칸 내용만 지웁니다. 여러 칸을 한 번에 지울 수 있습니다"],
                    [u"Backspace", u"한 칸만, 지우고 바로 칸 안으로 들어갑니다"],
                    [u"Ctrl+Z", u"방금 한 일을 되돌립니다"]])
    head(ws, 11, 2, [u"여기를 지워 보세요", u"", u""], [20, 12, 12])
    body(ws, 12, 2, [[u"지워도 됩니다", 111, 222], [u"이것도", 333, 444]])
    return wb


# ────────────────────────────── 4장 · 행·열·시트 ──────────────────────────────
def b04(done):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = u"줄과 칸"
    title(ws, u"칸이 좁아 안 보이는 표", u"C열이 좁아 ##### 이 뜹니다. 넓혀 보세요.")
    head(ws, 5, 2, [u"품목", u"금액", u"들어온 날"], [12, 8, 12])
    ws.column_dimensions["C"].width = 60 if done else 5
    ws.column_dimensions["D"].width = 14
    body(ws, 6, 2, [[u"복사용지", 1250000, u"2026-03-05"],
                    [u"토너",     890000,  u"2026-03-07"],
                    [u"의자",    2400000,  u"2026-03-11"],
                    [u"책상",    3150000,  u"2026-03-12"]], fmt={1: "#,##0"})
    hint(ws, "F5", u"· ##### 은 오류가 아닙니다. 칸이 좁을 뿐입니다\n"
                   u"· 열 이름 경계선을 두 번 누르면 알아서 맞춰집니다\n"
                   u"· 줄을 끼워 넣으려면 행 번호를 오른쪽 눌러 「삽입」\n"
                   u"· 내용만 지우기(Delete)와 줄 없애기(삭제)는 다릅니다", "F:J")

    ws = wb.create_sheet(u"1월")
    title(ws, u"1월 매출"); head(ws, 5, 2, [u"지점", u"매출"], [12, 12])
    body(ws, 6, 2, [[u"강남점", 4200], [u"신촌점", 3100]], fmt={1: "#,##0"})
    ws = wb.create_sheet(u"2월")
    title(ws, u"2월 매출"); head(ws, 5, 2, [u"지점", u"매출"], [12, 12])
    body(ws, 6, 2, [[u"강남점", 3800], [u"신촌점", 3300]], fmt={1: "#,##0"})
    if done:
        ws = wb.create_sheet(u"3월")
        title(ws, u"3월 매출"); head(ws, 5, 2, [u"지점", u"매출"], [12, 12])
        body(ws, 6, 2, [[u"강남점", 4500], [u"신촌점", 2900]], fmt={1: "#,##0"})
    return wb


# ─────────────────────────────── 5장 · 저장 ───────────────────────────────
def b05(done):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = u"저장 연습"
    title(ws, u"3월 회비 장부", u"이 파일로 「다른 이름으로 저장」과 PDF 를 연습합니다.")
    head(ws, 5, 2, [u"이름", u"낸 날", u"금액"], [12, 14, 12])
    body(ws, 6, 2, [[u"김세민", u"2026-03-02", 30000],
                    [u"박정화", u"2026-03-02", 30000],
                    [u"이서우", u"2026-03-05", 30000],
                    [u"정다온", u"2026-03-09", 30000],
                    [u"김준용", u"2026-03-11", 30000]], fmt={2: "#,##0"})
    hint(ws, "F5", u"· Ctrl+S 는 덮어쓰기 — 원본이 사라집니다\n"
                   u"· 원본을 남기려면 F12(다른 이름으로 저장)\n"
                   u"· 남에게 보낼 때는 PDF — 상대가 못 고칩니다\n"
                   u"· 파일 이름에 날짜를 넣으면 나중에 찾기 쉽습니다", "F:J")
    return wb


# ───────────────────────────── 6장 · 자동 채우기 ─────────────────────────────
def b06(done):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = u"끌어 보기"
    title(ws, u"씨앗을 하나 놓고 아래로 끕니다", u"각 열의 첫 칸(또는 두 칸)만 있습니다.")
    seeds = [(u"번호", [1], [1, 2, 3, 4, 5, 6]),
             (u"홀수", [1, 3], [1, 3, 5, 7, 9, 11]),
             (u"요일", [u"월요일"], [u"월요일", u"화요일", u"수요일", u"목요일", u"금요일", u"토요일"]),
             (u"달",   [u"1월"], [u"1월", u"2월", u"3월", u"4월", u"5월", u"6월"]),
             (u"분기", [u"1분기"], [u"1분기", u"2분기", u"3분기", u"4분기", u"1분기", u"2분기"]),
             (u"제품", [u"제품-001"], [u"제품-001", u"제품-002", u"제품-003",
                                       u"제품-004", u"제품-005", u"제품-006"])]
    head(ws, 5, 2, [s[0] for s in seeds], [10] * 6)
    for i, (_, seed, full) in enumerate(seeds):
        vals = full if done else seed
        for r in range(6):
            c = ws.cell(6 + r, 2 + i); c.border = BOX; c.alignment = CTR
            if r >= len(seed):
                c.fill = TODO
            if r < len(vals):
                c.value = vals[r]
    hint(ws, "J5", u"· 한 칸만 잡고 끌면 → 그대로 복사됩니다\n"
                   u"· 두 칸을 잡고 끌면 → 그 간격으로 늘어납니다\n"
                   u"· 요일·달·분기는 한 칸만 잡아도 늘어납니다\n"
                   u"· 끌고 나면 오른쪽 아래 작은 단추에서 고를 수 있습니다", "J:N")

    ws = wb.create_sheet(u"빠른 채우기")
    title(ws, u"Ctrl+E — 예를 보여 주면 따라 합니다")
    head(ws, 5, 2, [u"이름(영문)", u"영문만"], [24, 16])
    src = [u"최경민(Stephen)", u"박재현(JH.Park)", u"김인후(John.Kim)",
           u"김민형(MinHyung.Kim)", u"이승오(Robert.Lee)", u"정수미(Sara.Jeong)"]
    out = [s[s.index(u"(") + 1:-1] for s in src]
    for r, s in enumerate(src):
        ws.cell(6 + r, 2, s).border = BOX
        c = ws.cell(6 + r, 3); c.border = BOX
        if r:
            c.fill = TODO
        if done or r == 0:
            c.value = out[r]
    return wb


# ───────────────────────────── 7장 · 계산 시키기 ─────────────────────────────
def b07(done):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = u"사칙연산"
    title(ws, u"금액 = 수량 × 단가", u"E열에 수식을 넣어 보세요. = 로 시작합니다.")
    head(ws, 5, 2, [u"품목", u"수량", u"단가", u"금액"], [12, 8, 10, 12])
    items = [(u"사과", 3, 1000), (u"배", 2, 2000), (u"감", 5, 800),
             (u"귤", 10, 500), (u"포도", 4, 3000)]
    for r, (n, q, p) in enumerate(items):
        ws.cell(6 + r, 2, n).border = BOX
        ws.cell(6 + r, 3, q).border = BOX
        ws.cell(6 + r, 4, p).border = BOX; ws.cell(6 + r, 4).number_format = "#,##0"
        c = ws.cell(6 + r, 5); c.border = BOX; c.number_format = "#,##0"; c.fill = TODO
        if done:
            c.value = "=C%d*D%d" % (6 + r, 6 + r)
    ws.cell(11, 2, u"합계").font = Font(bold=True)
    c = ws.cell(11, 5); c.border = BOX; c.number_format = "#,##0"
    c.font = Font(bold=True); c.fill = TODO
    if done:
        c.value = "=SUM(E6:E10)"
    hint(ws, "G5", u"· 수식은 반드시 = 로 시작합니다\n"
                   u"· 곱하기는 * (별표), 나누기는 / (빗금)\n"
                   u"· 숫자를 직접 치지 말고 칸 주소를 쓰세요 → =C6*D6\n"
                   u"· 한 줄 만들고 아래로 끌면 나머지는 저절로 됩니다", "G:K")

    ws = wb.create_sheet(u"함수 다섯")
    title(ws, u"단추 하나로 되는 다섯 가지")
    head(ws, 5, 2, [u"이름", u"국어", u"영어", u"수학"], [12, 8, 8, 8])
    scores = [(u"김세민", 88, 92, 79), (u"박정화", 75, 68, 91),
              (u"이서우", 94, 85, 88), (u"정다온", 61, 77, 70),
              (u"김준용", 82, 90, 95)]
    body(ws, 6, 2, [list(s) for s in scores])
    asks = [(u"합계", "=SUM(C6:C10)"), (u"평균", "=AVERAGE(C6:C10)"),
            (u"개수", "=COUNT(C6:C10)"), (u"최고", "=MAX(C6:C10)"),
            (u"최저", "=MIN(C6:C10)")]
    for r, (label, f) in enumerate(asks):
        ws.cell(12 + r, 2, label).font = Font(bold=True)
        ws.cell(12 + r, 2).border = BOX
        for i in range(3):
            c = ws.cell(12 + r, 3 + i); c.border = BOX; c.fill = TODO
            if done:
                c.value = f.replace("C6:C10", "%s6:%s10" % ("CDE"[i], "CDE"[i]))
                if label == u"평균":
                    c.number_format = "0.0"
    return wb


# ──────────────────────────── 8장 · 참조와 조건 ────────────────────────────
def b08(done):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = u"$ 고정"
    title(ws, u"할인율은 한 칸에만 있습니다", u"끌어내려도 그 칸을 보게 해야 합니다.")
    ws["E3"] = u"할인율"; ws["E3"].font = Font(bold=True)
    ws["F3"] = 0.15; ws["F3"].number_format = "0%"
    ws["F3"].fill = PatternFill("solid", fgColor="FFF2CC"); ws["F3"].border = BOX
    head(ws, 5, 2, [u"품목", u"정가", u"할인액"], [12, 12, 12])
    items = [(u"의자", 240000), (u"책상", 315000), (u"선반", 98000),
             (u"조명", 52000), (u"서랍", 176000)]
    for r, (n, p) in enumerate(items):
        ws.cell(6 + r, 2, n).border = BOX
        ws.cell(6 + r, 3, p).border = BOX; ws.cell(6 + r, 3).number_format = "#,##0"
        c = ws.cell(6 + r, 4); c.border = BOX; c.number_format = "#,##0"; c.fill = TODO
        if done:
            c.value = "=C%d*$F$3" % (6 + r)
    hint(ws, "F6", u"· 그냥 =C6*F3 으로 하고 끌면 F4·F5… 로 밀려 0 이 나옵니다\n"
                   u"· $F$3 으로 쓰면 그 칸에 붙박여 안 움직입니다\n"
                   u"· 수식에서 F4 키를 누르면 $ 가 자동으로 붙습니다", "F:J", 5)

    ws = wb.create_sheet(u"IF 조건")
    title(ws, u"70점이 넘으면 합격", u"쓰는 모양 — IF(조건, 맞을 때, 틀릴 때)")
    head(ws, 5, 2, [u"이름", u"점수", u"판정"], [12, 8, 12])
    rows = [(u"김세민", 88), (u"박정화", 65), (u"이서우", 94),
            (u"정다온", 70), (u"김준용", 58), (u"최효윤", 77)]
    for r, (n, s) in enumerate(rows):
        ws.cell(6 + r, 2, n).border = BOX
        ws.cell(6 + r, 3, s).border = BOX; ws.cell(6 + r, 3).alignment = CTR
        c = ws.cell(6 + r, 4); c.border = BOX; c.alignment = CTR; c.fill = TODO
        if done:
            c.value = '=IF(C%d>=70,"합격","불합격")' % (6 + r)
    hint(ws, "F5", u'· =IF(C6>=70,"합격","불합격")\n'
                   u"· 글자는 따옴표로 감쌉니다\n"
                   u"· >= 는 「크거나 같다」 — 정다온(70)은 합격입니다", "F:J", 5)
    return wb


# ───────────────────────────── 9장 · 보기 좋게 ─────────────────────────────
def b09(done):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = u"꾸미기 전"
    ws["B2"] = u"분기별 지점 매출"
    if done:
        ws["B2"].font = Font(bold=True, size=14)
    rows = [(u"강남점", 4200000, 3800000, 4500000),
            (u"신촌점", 3100000, 3300000, 2900000),
            (u"구로점", 2800000, 2600000, 3000000),
            (u"용산점", 3600000, 3900000, 3700000),
            (u"서면점", 2400000, 2500000, 2200000)]
    names = [u"지점", u"1월", u"2월", u"3월"]
    for i, n in enumerate(names):
        c = ws.cell(4, 2 + i, n)
        if done:
            c.fill, c.font, c.alignment, c.border = HDR_FILL, HDR_FONT, CTR, BOX
    for r, vals in enumerate(rows):
        for i, v in enumerate(vals):
            c = ws.cell(5 + r, 2 + i, v)
            if done:
                c.border = BOX
                if i:
                    c.number_format = "#,##0"
    if done:
        for col, w in zip("BCDE", [12, 12, 12, 12]):
            ws.column_dimensions[col].width = w
    hint(ws, "G4", u"· 숫자에 쉼표: Ctrl+1 → 표시 형식 → 숫자 → 1000 단위 구분\n"
                   u"· 머리글: 굵게 + 바탕색 + 가운데\n"
                   u"· 테두리: 홈 → 테두리 → 모든 테두리\n"
                   u"· 겉모습만 바뀌는 것입니다. 속 값은 그대로입니다", "G:K")

    ws = wb.create_sheet(u"조건부 서식")
    title(ws, u"큰 값에 저절로 색이 들어가게", u"300만 원이 넘는 칸을 칠해 보세요.")
    head(ws, 5, 2, names, [12, 12, 12, 12])
    body(ws, 6, 2, [list(r) for r in rows], fmt={1: "#,##0", 2: "#,##0", 3: "#,##0"})
    return wb


# ────────────────────── 10장 · 정리하고 내보내기 ──────────────────────
def b10(done):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = u"명단"
    title(ws, u"사원 명단", u"정렬 · 필터 · 틀 고정 · 차트를 모두 여기서 합니다.")
    head(ws, 5, 2, [u"이름", u"부서", u"입사년도", u"평가점수"], [12, 12, 10, 10])
    people = [(u"김세민", u"영업팀", 2019, 88), (u"박정화", u"개발팀", 2021, 75),
              (u"이서우", u"생산팀", 2018, 94), (u"정다온", u"총무팀", 2022, 61),
              (u"김준용", u"인사팀", 2020, 82), (u"최효윤", u"영업팀", 2023, 77),
              (u"박단비", u"IT팀",   2017, 90), (u"이유림", u"기획팀", 2021, 68),
              (u"정진하", u"생산팀", 2019, 85), (u"김병민", u"개발팀", 2020, 72),
              (u"최리",   u"재무팀", 2022, 79), (u"이제우", u"영업팀", 2018, 93)]
    body(ws, 6, 2, [list(p) for p in people])
    for r in range(6, 18):
        ws.cell(r, 4).alignment = CTR
        ws.cell(r, 5).alignment = CTR
    if done:
        ws.freeze_panes = "B6"
        ws.auto_filter.ref = "B5:E17"
    hint(ws, "G5", u"· 정렬은 표 안 아무 칸을 누르고 데이터 → 정렬\n"
                   u"· 필터는 Ctrl+Shift+L — 한 번 더 누르면 꺼집니다\n"
                   u"· 틀 고정은 B6 을 누르고 보기 → 틀 고정\n"
                   u"· 차트는 범위를 잡고 삽입 → 세로 막대형", "G:K")

    ws = wb.create_sheet(u"차트용")
    title(ws, u"부서별 인원", u"이 표로 막대 그래프를 만들어 보세요.")
    head(ws, 5, 2, [u"부서", u"인원"], [12, 10])
    counts = {}
    for p in people:
        counts[p[1]] = counts.get(p[1], 0) + 1
    body(ws, 6, 2, [[k, v] for k, v in sorted(counts.items(), key=lambda x: -x[1])])
    return wb


BUILD = [
    (u"B-02_입력_연습",        b02, True),
    (u"B-03_고치기_연습",      b03, True),
    (u"B-04_행열_연습",        b04, True),
    (u"B-05_저장_연습",        b05, False),   # 결과가 새 파일이라 완성본 없음
    (u"B-06_자동채우기",       b06, True),
    (u"B-07_계산_연습",        b07, True),
    (u"B-08_참조와_조건",      b08, True),
    (u"B-09_서식_연습",        b09, True),
    (u"B-10_정리하고_내보내기", b10, True),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--write", action="store_true", help=u"실제로 파일을 만듭니다")
    a = ap.parse_args()

    if a.write and not os.path.isdir(DST):
        os.makedirs(DST)

    n = 0
    for name, fn, has_done in BUILD:
        jobs = [(name + u".xlsx", False)]
        if has_done:
            jobs.append((name + u"_완성.xlsx", True))
        for fname, done in jobs:
            path = os.path.join(DST, fname)
            if a.write:
                fn(done).save(path)
            print(u"%s %s" % (u"만듦 " if a.write else u"만들 것", path))
            n += 1
    print(u"\n모두 %d개%s" % (n, u"" if a.write else u" — 실제로 만들려면 --write"))


if __name__ == "__main__":
    main()
