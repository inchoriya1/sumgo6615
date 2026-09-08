# -*- coding: utf-8 -*-
"""
근태를 넣으면 근무일수가 저절로 채워지게 합니다.

  python tools/auto-workdays.py 급여대장.xlsx            # 보기만
  python tools/auto-workdays.py 급여대장.xlsx --write    # 새 파일 만들기

지금까지
  매달 [월별입력] 에 여섯 사람의 근무일수를 손으로 적었습니다.

바꾼 뒤
  **근태를 넣으면 그 달 그 사람이 켜지고, 근무일수가 저절로 들어갑니다.**
  다른 값을 줄 달만 월별입력에 적으면 됩니다(적으면 그 값이 이깁니다).

기본값을 어떻게 내나 — 자료에서 찾아낸 규칙
  · 209H · 일수 → **그 달 날수 − 일요일 수**.
    6·7·8월 13건 중 10건이 정확히 일치했고, 안 맞는 3건은 전부
    비고에 이유가 적혀 있었습니다(연차 1일, 중도 입사, 대장 기준).
  · 일급제 → 달력과 무관한 **합의된 고정값**(21일 + 주차 5일)이라,
    [직원설정] 에 `기본근무일수`·`기본주차일수` 칸을 만들어 거기서 가져옵니다.

'그 달을 켜는' 신호는 근태입니다
  근태원본에 **출퇴근이 하나라도 들어온 달**만 켭니다.
  그래서 아직 근태를 안 넣은 달은 0 그대로이고,
  **연간집계에 미래 달이 미리 찍히지 않습니다.**

  다만 **고정월급(대표)은 근태기록이 없어서** 이 신호가 안 옵니다.
  그 한 사람만 지금처럼 월별입력에 적습니다.

입사일·퇴사일이 이제 진짜로 일합니다
  전에는 퇴사일이 **표시용**이라, 적어도 급여가 안 멈췄습니다(함정).
  이제 **재직 기간 밖의 달은 무조건 0** 입니다.
  퇴사일만 적으면 그다음 달부터 알아서 멈춥니다.
"""
import argparse
import re

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ENGINE, INPUT, AGG = u"기본정보", u"월별입력", u"근태집계"
SET_EMP, RAW = u"직원설정", u"근태원본"
ROW0, ROW1 = 2, 193          # 기본정보 24개월 × 8명
SLOTS = 8
EMP0 = 2                     # 직원설정 첫 줄
RAW_FIRST, RAW_LAST = 2, 5953
COL_DAYS, COL_WEEK = 12, 13  # 기본정보 L 근무일수 · M 주차일수
AGG_TOTAL, AGG_SUN, AGG_HAS = 10, 11, 12   # 근태집계에 새로 만들 J·K·L
EMP_DAYS, EMP_WEEK = 25, 26                # 직원설정에 새로 만들 Y·Z

FONT = u"맑은 고딕"
FILL_HEAD = PatternFill("solid", fgColor=u"FF1F4E79")
FILL_IN = PatternFill("solid", fgColor=u"FFFFE699")
FILL_AUTO = PatternFill("solid", fgColor=u"FFF2F2F2")
THIN = Side(style="thin", color=u"FFD9D9D9")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# 일급제처럼 달력으로 안 나오는 사람의 고정값. 자료에서 확인된 값입니다.
FIXED_DEFAULTS = {u"일급제": (21, 5)}


def head(ws, col, text):
    c = ws.cell(1, col)
    c.value = text
    c.font = Font(name=FONT, size=9, bold=True, color=u"FFFFFFFF")
    c.fill = FILL_HEAD
    c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center")


def slot_of(row):
    """기본정보 row → 직원설정 row (여덟 줄마다 되풀이)."""
    return EMP0 + (row - ROW0) % SLOTS


def month_first(row):
    return u'DATE(LEFT($BB{r},4),MID($BB{r},6,2),1)'.format(r=row)


def working(row):
    """재직 중인 달인가 — 입사일 이후이고, 퇴사했다면 퇴사월까지."""
    first = month_first(row)
    return (u'AND($E{r}<=EOMONTH({f},0),OR(NOT(ISNUMBER($F{r})),$F{r}>={f}))'
            .format(r=row, f=first))


def run(path, out_path, do_write):
    wb = openpyxl.load_workbook(path, data_only=False)
    for n in (ENGINE, INPUT, AGG, SET_EMP, RAW):
        if n not in wb.sheetnames:
            raise SystemExit(u"[%s] 시트가 없습니다." % n)
    st = wb[SET_EMP]

    print(u"■ 지금  근무일수 = %s" % wb[ENGINE].cell(ROW0, COL_DAYS).value)
    print(u"■ 바꾼 뒤")
    print(u"   ① 월별입력에 적혀 있으면 → 그 값")
    print(u"   ② 재직 기간 밖이면 → 0  (입사일·퇴사일이 이제 실제로 판단합니다)")
    print(u"   ③ 그 달 근태가 들어왔으면 → 직원설정 기본값, 없으면 (그 달 날수 − 일요일 수)")
    print(u"   ④ 근태가 아직 없으면 → 0  (미래 달이 미리 계산되지 않습니다)")
    print(u"\n   ※ 고정월급인 분은 근태기록이 없어 ③ 신호가 안 옵니다 — 지금처럼 손으로 적습니다.")

    ways = {st.cell(EMP0 + i, 9).value: EMP0 + i for i in range(SLOTS)}
    print(u"\n■ 직원설정에 넣을 기본값")
    for i in range(SLOTS):
        nm, way = st.cell(EMP0 + i, 2).value, st.cell(EMP0 + i, 9).value
        if not nm:
            continue
        if way in FIXED_DEFAULTS:
            d, w = FIXED_DEFAULTS[way]
            print(u"   %-6s %-7s 기본근무일수 %s · 기본주차일수 %s" % (nm, way, d, w))
        elif way == u"고정월급":
            print(u"   %-6s %-7s (근태가 없어 자동 대상이 아닙니다 — 월별입력에 적으세요)" % (nm, way))
        else:
            print(u"   %-6s %-7s 달력에서 (그 달 날수 − 일요일 수)" % (nm, way))

    if not do_write:
        print(u"\n※ 보기만 했습니다. 실제로 만들려면 --write 를 붙이세요.")
        return

    # ── 근태집계에 도우미 세 열 ──
    ag = wb[AGG]
    head(ag, AGG_TOTAL, u"그달날수")
    head(ag, AGG_SUN, u"일요일수")
    head(ag, AGG_HAS, u"근태들어옴")
    for r in range(ROW0, ROW1 + 1):
        ag.cell(r, AGG_TOTAL).value = (
            u'=IF($C{r}="",0,COUNTIF({R}!$V${a}:$V${b},$C{r}))'
            .format(r=r, R=RAW, a=RAW_FIRST, b=RAW_LAST))
        ag.cell(r, AGG_SUN).value = (
            u'=IF($C{r}="",0,COUNTIFS({R}!$V${a}:$V${b},$C{r},'
            u'{R}!$R${a}:$R${b},"일"))'.format(r=r, R=RAW, a=RAW_FIRST, b=RAW_LAST))
        ag.cell(r, AGG_HAS).value = (
            u'=IF($C{r}="",0,COUNTIFS({R}!$V${a}:$V${b},$C{r},'
            u'{R}!$D${a}:$D${b},"<>"))'.format(r=r, R=RAW, a=RAW_FIRST, b=RAW_LAST))
        for c in (AGG_TOTAL, AGG_SUN, AGG_HAS):
            ag.cell(r, c).fill = FILL_AUTO
            ag.cell(r, c).border = BOX
            ag.cell(r, c).alignment = Alignment(horizontal="center")

    # ── 직원설정에 기본값 두 열 ──
    head(st, EMP_DAYS, u"기본근무일수")
    head(st, EMP_WEEK, u"기본주차일수")
    st.column_dimensions["Y"].width = 13
    st.column_dimensions["Z"].width = 13
    for i in range(SLOTS):
        r = EMP0 + i
        for c in (EMP_DAYS, EMP_WEEK):
            st.cell(r, c).fill = FILL_IN
            st.cell(r, c).border = BOX
            st.cell(r, c).alignment = Alignment(horizontal="center")
        way = st.cell(r, 9).value
        if way in FIXED_DEFAULTS:
            st.cell(r, EMP_DAYS).value, st.cell(r, EMP_WEEK).value = FIXED_DEFAULTS[way]
    st.cell(11, 25).value = (u"※ 비워 두면 '그 달 날수 − 일요일 수' 를 씁니다. "
                             u"달력으로 안 나오는 사람(일급제 등)만 적으세요.")
    st.cell(11, 25).font = Font(name=FONT, size=10, color=u"FFC00000")

    # ── 기본정보 근무일수·주차일수 ──
    gi = wb[ENGINE]
    for r in range(ROW0, ROW1 + 1):
        s = slot_of(r)
        # '그 달을 켜는' 신호는 오직 근태입니다.
        # 고정월급도 예외를 두지 않습니다 — 예외를 두면 근태가 없는 대표만
        # 2027년 12월까지 매달 급여가 미리 계산되어, 연간집계가 '아직 안 준 돈'
        # 까지 더해 버립니다. 대표는 월별입력에 손으로 적습니다.
        on = u'{A}!$L{r}>0'.format(A=AGG, r=r)
        cal = u'({A}!$J{r}-{A}!$K{r})'.format(A=AGG, r=r)
        base_d = (u'IF({E}!$Y${s}<>"",{E}!$Y${s},{c})'.format(E=SET_EMP, s=s, c=cal))
        gi.cell(r, COL_DAYS).value = (
            u'=IF($B{r}="",0,IF(NOT({w}),0,'
            u'IF({I}!D{r}<>"",{I}!D{r},IF({on},{bd},0))))'
            .format(r=r, w=working(r), I=INPUT, on=on, bd=base_d))
        gi.cell(r, COL_WEEK).value = (
            u'=IF($B{r}="",0,IF(NOT({w}),0,'
            u'IF({I}!E{r}<>"",{I}!E{r},IF({on},N({E}!$Z${s}),0))))'
            .format(r=r, w=working(r), I=INPUT, on=on, E=SET_EMP, s=s))

    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)
    print(u"\n근태집계에 도우미 3열 · 직원설정에 기본값 2열을 만들고,")
    print(u"기본정보의 근무일수·주차일수 %d줄을 고쳤습니다." % (ROW1 - ROW0 + 1))
    print(u"만들었습니다: %s" % out_path)
    print(u"원본은 그대로 두었습니다: %s" % path)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("book")
    ap.add_argument("--out", default=None)
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args()
    out = a.out or re.sub(r"\.xlsx$", "", a.book) + u"_근무일수자동.xlsx"
    run(a.book, out, a.write)
