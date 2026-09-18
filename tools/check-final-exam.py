# -*- coding: utf-8 -*-
u"""종합문제 1~7장을 **진짜 엑셀로** 풀어 봅니다.

  python tools/check-final-exam.py

엑셀(COM)과 pywin32 가 있는 윈도에서만 돕니다.
사본을 임시 폴더에 떠서 [정답] 시트의 단계를 엑셀 명령으로 따라 한 뒤,
재계산된 값을 최종확인 정답과 견줍니다. 원본 파일은 건드리지 않습니다.

build-final-exam.py 가 답을 **파이썬으로** 계산하므로 숫자는 맞게 나옵니다.
이 도구가 잡는 것은 그 너머입니다 — 엑셀에서 **그 단계가 실제로 되는가**.
2026-09-18 에 이것으로 찾은 것: 5·7장 파일이 아예 안 열림(설명문이 수식으로
저장됨), 4장 ⑤ 에 IFERROR 가 빠져 합계가 #N/A, 6장 ⑤ 계산 항목은 엑셀이 거부,
7장 단가표에서 VLOOKUP 이 그냥 됨(함정이 성립 안 함). 8장은 차트라 빠집니다.
"""
import glob
import os
import shutil
import sys
import tempfile

import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.abspath(u".")
TMP = tempfile.mkdtemp(prefix=u"exam_")

xl = win32.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

results = []


def open_copy(chap):
    src = glob.glob(os.path.join(ROOT, u"강의예제", u"%d장_*" % chap, u"%d장_종합문제.xlsx" % chap))[0]
    dst = os.path.join(TMP, u"%d.xlsx" % chap)
    shutil.copy(src, dst)
    return xl.Workbooks.Open(dst)


def answers(wb):
    ws = wb.Worksheets(u"정답")
    out = {}
    r = 5
    while r < 60:
        b = ws.Cells(r, 2).Value
        c = ws.Cells(r, 3).Value
        if b in (u"1", u"2", u"3", u"4", u"5") and c is not None:
            out[int(b)] = c
        r += 1
    return out


def check(chap, no, got, want, note=u""):
    if isinstance(want, float) and isinstance(got, (int, float)):
        ok = abs(got - want) < 0.006
    elif isinstance(want, (int, float)) and isinstance(got, (int, float)):
        ok = abs(got - want) < 1e-6
    else:
        ok = (u"%s" % got).strip() == (u"%s" % want).strip()
    results.append((chap, no, ok, got, want, note))


def ch1():
    wb = open_copy(1); ws = wb.Worksheets(u"문제"); ans = answers(wb)
    ws.Range("B6:B7").AutoFill(ws.Range("B6:B11"), 0)
    ws.Range("C6:C7").AutoFill(ws.Range("C6:C11"), 0)
    ws.Range("F6").Value = "HB.Kim"
    ws.Range("F6:F11").FlashFill()
    for r in range(6, 12):
        v = ws.Cells(r, 7).Value
        if isinstance(v, str):
            ws.Cells(r, 7).Value = float(v)
    ws.Range("I6").Formula = "=G6*H6"; ws.Range("I6").AutoFill(ws.Range("I6:I11"), 0)
    ws.Range("J6").Formula = "=I6>=$C$4"; ws.Range("J6").AutoFill(ws.Range("J6:J11"), 0)
    ws.Range("K6").Formula = "=I6*$C$3"; ws.Range("K6").AutoFill(ws.Range("K6:K11"), 0)
    xl.CalculateFull()
    f = xl.WorksheetFunction
    check(1, 1, f.Sum(ws.Range("G6:G11")), ans[1])
    check(1, 2, f.Sum(ws.Range("I6:I11")), ans[2])
    check(1, 3, f.Sum(ws.Range("K6:K11")), ans[3])
    check(1, 4, f.CountIf(ws.Range("J6:J11"), True), ans[4])
    check(1, 5, f.CountIf(ws.Range("F6:F11"), "*.Kim"), ans[5], u"빠른 채우기 결과: %s" % [ws.Cells(r, 6).Value for r in range(6, 12)])
    wb.Close(False)


def ch2():
    wb = open_copy(2); ws = wb.Worksheets(u"문제"); ans = answers(wb)
    ws.Range("B6:B12").UnMerge()
    try:
        ws.Range("B6:B12").SpecialCells(4).FormulaR1C1 = "=R[-1]C"
    except Exception:
        pass
    ws.Range("B6:B12").Value = ws.Range("B6:B12").Value
    # 여러 줄 나누기 (Ctrl+J → 줄바꿈 문자)
    for r in range(6, 13):
        v = ws.Cells(r, 4).Value or u""
        parts = v.split(u"\n")
        ws.Cells(r, 5).Value = parts[0]
        ws.Cells(r, 6).Value = parts[1] if len(parts) > 1 else None
        code = ws.Cells(r, 8).Value.split(u"-")
        ws.Cells(r, 9).Value, ws.Cells(r, 10).Value, ws.Cells(r, 11).Value = code[0], code[1], code[2]
    wb.Names.Add(Name=u"매출", RefersTo="=문제!$L$6:$N$12")
    ws.Range("O6").Formula = "=SUM(L6:N6)"; ws.Range("O6").AutoFill(ws.Range("O6:O12"), 0)
    xl.CalculateFull()
    f = xl.WorksheetFunction
    check(2, 1, f.SumIfs(ws.Range("O6:O12"), ws.Range("B6:B12"), u"수도권"), ans[1])
    check(2, 2, xl.Evaluate(u"=SUM(매출)"), ans[2], u"O열 합계 %s" % f.Sum(ws.Range("O6:O12")))
    codes = [ws.Cells(r, 8).Value for r in range(6, 13)]
    check(2, 3, len(codes) - len(set(codes)), ans[3], u"코드 %s" % codes)
    check(2, 4, f.CountA(ws.Range("E6:F12")), ans[4])
    check(2, 5, len(set(ws.Cells(r, 9).Value for r in range(6, 13))), ans[5])
    wb.Close(False)


def ch3():
    wb = open_copy(3); ws = wb.Worksheets(u"문제"); ans = answers(wb)
    wb.Close(False)
    wb = open_copy(3); ws = wb.Worksheets(u"문제")
    ws.Range("C6:J10").NumberFormat = "#,##0,;-#,##0,;-;@"
    ws.Range("C3").Value = u"(단위: 천원)"
    ws.Range("K6").Formula = "=J6/I6"; ws.Range("K6").AutoFill(ws.Range("K6:K10"), 0)
    xl.CalculateFull()
    f = xl.WorksheetFunction
    check(3, 1, f.Sum(ws.Range("J6:J10")), ans[1])
    check(3, 2, f.CountIf(ws.Range("K6:K10"), "<1"), ans[2])
    check(3, 3, ws.Cells(5 + int(f.Match(f.Max(ws.Range("J6:J10")), ws.Range("J6:J10"), 0)), 2).Value, ans[3])
    check(3, 4, f.CountIf(ws.Range("C6:H10"), 0), ans[4], u"0 칸의 화면 표시: '%s'" % ws.Range("E6").Text)
    check(3, 5, ws.Range("C3").Value, ans[5])
    wb.Close(False)


def ch4():
    wb = open_copy(4); ws = wb.Worksheets(u"문제"); ans = answers(wb)
    last = ws.Cells.SpecialCells(11)
    check(4, 4, last.Row, ans[4], u"Ctrl+End = %s" % last.Address)
    n = ws.UsedRange.Rows.Count
    ws.Range("A15:A1000").EntireRow.Delete()
    for r in range(6, 14):
        v = ws.Cells(r, 5).Value
        if isinstance(v, str):
            ws.Cells(r, 5).Value = float(v)
    ws.Range("F6").Formula = "=INDEX($M$6:$M$10,MATCH(C6,$K$6:$K$10,0))"
    ws.Range("F6").AutoFill(ws.Range("F6:F13"), 0)
    ws.Range("G6").Formula = "=E6*F6"; ws.Range("G6").AutoFill(ws.Range("G6:G13"), 0)
    ws.Range("E15").Formula = "=SUM(E6:E13)"
    ws.Range("G15").Formula = "=SUM(G6:G13)"
    xl.CalculateFull()
    na = [r for r in range(6, 14) if ws.Cells(r, 6).Text == "#N/A"]
    check(4, 1, ws.Range("E15").Value, ans[1])
    check(4, 3, len(na), ans[3], u"#N/A 줄: %s" % na)
    note = u"단계 ⑤ 문구대로 하면 금액 합계 칸 = %s" % ws.Range("G15").Text
    ws.Range("F6").Formula = "=IFERROR(INDEX($M$6:$M$10,MATCH(C6,$K$6:$K$10,0)),0)"
    ws.Range("F6").AutoFill(ws.Range("F6:F13"), 0)
    xl.CalculateFull()
    check(4, 2, ws.Range("G15").Value, ans[2], note + u" → IFERROR(…,0) 을 씌우면 %s" % ws.Range("G15").Text)
    check(4, 5, u"$5:$5", ans[5])
    wb.Close(False)


def ch5():
    wb = open_copy(5); ws = wb.Worksheets(u"문제"); ans = answers(wb)
    last = ws.Cells.SpecialCells(11).Row
    ws.Range("B5:F%d" % last).AutoFilter(3, u"소계")
    vis = ws.Range("B6:F%d" % last).SpecialCells(12)
    vis.EntireRow.Delete()
    ws.AutoFilterMode = False
    last = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
    rows = last - 5
    check(5, 1, rows, ans[1], u"데이터 끝 행 %d" % last)
    ws.Range("J3").Formula = "=SUBTOTAL(9,F6:F%d)" % last
    xl.CalculateFull()
    check(5, 2, ws.Range("J3").Value, ans[2], u"단계 ④ 의 범위는 F6:F16 — 실제 끝 행 %d" % last)
    ws.Range("B5:F%d" % last).AutoFilter(1, u"서울")
    xl.CalculateFull()
    check(5, 3, ws.Range("J3").Value, ans[3])
    ws.AutoFilterMode = False
    ex = wb.Worksheets(u"추출")
    ex.Range("B3").Value = u"대분류"; ex.Range("C3").Value = u"수량"
    ex.Range("B4").Value = u"사무용품"; ex.Range("C4").Value = ">=100"
    ex.Range("B5").Value = u"가구"; ex.Range("C5").Value = ">=80"
    for c in "DEF":
        for r in (3, 4, 5):
            ex.Range("%s%d" % (c, r)).Value = None
    ws.Range("B5:F%d" % last).AdvancedFilter(2, ex.Range("B3:C5"), ex.Range("B8"), False)
    out_last = ex.Cells(ex.Rows.Count, 2).End(-4162).Row
    check(5, 4, out_last - 8, ans[4], u"추출 결과 B8~B%d" % out_last)
    f = xl.WorksheetFunction
    check(5, 5, f.Large(ws.Range("F6:F%d" % last), 1) + f.Large(ws.Range("F6:F%d" % last), 2) + f.Large(ws.Range("F6:F%d" % last), 3), ans[5])
    wb.Close(False)


def ch6():
    wb = open_copy(6); src = wb.Worksheets(u"원본"); dt = wb.Worksheets(u"자료"); ans = answers(wb)
    r = 4
    for i in range(6, 11):
        for k in range(5):
            dt.Cells(r, 2).Value = src.Cells(i, 2).Value
            dt.Cells(r, 3).Value = 2021 + k
            dt.Cells(r, 4).Value = src.Cells(i, 3 + 2 * k).Value
            dt.Cells(r, 5).Value = src.Cells(i, 4 + 2 * k).Value
            r += 1
    check(6, 1, r - 4, ans[1])
    lo = dt.ListObjects.Add(1, dt.Range("B3:E%d" % (r - 1)), None, 1)
    lo.Name = u"지점매출"
    pv_ws = wb.Worksheets.Add()
    pc = wb.PivotCaches().Create(1, u"지점매출")
    pt = pc.CreatePivotTable(pv_ws.Range("A3"), u"피벗")
    pt.PivotFields(u"지점").Orientation = 1
    pt.PivotFields(u"연도").Orientation = 2
    pt.AddDataField(pt.PivotFields(u"매출"), u"합계: 매출", -4157)
    xl.CalculateFull()
    y2023 = pt.GetPivotData(u"합계: 매출", u"연도", 2023).Value
    check(6, 2, y2023, ans[2])
    best = max(((pt.GetPivotData(u"합계: 매출", u"지점", n).Value, n) for n in [u"강남점", u"명동점", u"서초점", u"일산점", u"용산점"]))
    check(6, 3, best[1], ans[3])
    pt.CalculatedFields().Add(u"이익률", u"= 이익 / 매출", True)
    df = pt.AddDataField(pt.PivotFields(u"이익률"), u"합계: 이익률", -4157)
    xl.CalculateFull()
    field_total = pt.GetPivotData(u"합계: 이익률").Value
    check(6, 4, round(field_total * 100, 2), ans[4])
    note5 = []
    try:
        pt.PivotFields(u"연도").CalculatedItems().Add(u"이익률항목", u"= 이익 / 매출", True)
        note5.append(u"계산 항목 만들어짐")
    except Exception as e:
        note5.append(u"계산 항목을 「= 이익 / 매출」 로 만들 수 없음 — %s" % (u"%s" % e)[:80])
    # 원본에 비율 열을 만들어 합계로 넣었을 때
    dt.Range("F3").Value = u"이익률"
    lo.Resize(dt.Range("B3:F%d" % (r - 1)))
    dt.Range("F4").Formula = u"=[@이익]/[@매출]"
    pt.RefreshTable()
    pt.AddDataField(pt.PivotFields(u"이익률"), u"합계: 줄별이익률", -4157)
    xl.CalculateFull()
    try:
        summed = pt.GetPivotData(u"합계: 줄별이익률").Value
    except Exception as e:
        summed = u"오류 %s" % e
    check(6, 5, round(summed * 100, 2) if isinstance(summed, float) else summed, ans[5], u" / ".join(note5) + u" / 원본 비율 열을 합계로 넣으면 %s" % summed)
    wb.Close(False)


def ch7():
    wb = open_copy(7); ws = wb.Worksheets(u"문제"); ans = answers(wb)
    f = xl.WorksheetFunction
    ws.Range("R6").Formula = "=LEN(B6)"; xl.Calculate(); len6 = ws.Range("R6").Value; ws.Range("R6").Value = None
    raw_hits = sum(1 for r in range(6, 14) if ws.Cells(r, 2).Value in [ws.Cells(k, 15).Value for k in range(6, 10)])
    ws.Range("H6").Formula = "=TRIM(B6)"; ws.Range("H6").AutoFill(ws.Range("H6:H13"), 0)
    ws.Range("I6").Formula = "=DATE(LEFT(C6,2)+2000,MID(C6,3,2),RIGHT(C6,2))"; ws.Range("I6").AutoFill(ws.Range("I6:I13"), 0)
    ws.Range("F6").Formula = "=INDEX($L$6:$L$10,MATCH(D6,$M$6:$M$10,0))"; ws.Range("F6").AutoFill(ws.Range("F6:F13"), 0)
    ws.Range("G6").Formula = "=E6*F6"; ws.Range("G6").AutoFill(ws.Range("G6:G13"), 0)
    ws.Range("P6").Formula = "=SUMIFS($G$6:$G$13,$H$6:$H$13,O6)"; ws.Range("P6").AutoFill(ws.Range("P6:P9"), 0)
    ws.Range("J6").Formula = '=DATEDIF(I6,$F$3,"M")'; ws.Range("J6").AutoFill(ws.Range("J6:J13"), 0)
    # VLOOKUP 이 정말 안 되는지
    ws.Range("R7").Formula = "=VLOOKUP(D6,$L$6:$M$10,1,0)"
    xl.CalculateFull()
    months = [ws.Cells(r, 10).Value for r in range(6, 14)]
    check(7, 1, raw_hits, ans[1], u"TRIM 전 B열이 O열 이름과 같은 줄")
    check(7, 2, f.Sum(ws.Range("G6:G13")), ans[2], u"TRIM 결과 %s" % [ws.Cells(r, 8).Value for r in range(6, 10)])
    best = max(((ws.Cells(k, 16).Value, ws.Cells(k, 15).Value) for k in range(6, 10)))
    check(7, 3, best[1], ans[3], u"%s" % [(ws.Cells(k, 15).Value, ws.Cells(k, 16).Value) for k in range(6, 10)])
    check(7, 4, sum(1 for m in months if isinstance(m, float) and m >= 6), ans[4], u"개월 %s" % months)
    check(7, 5, len6, ans[5], u"VLOOKUP 시도 → %s" % ws.Range("R7").Text)
    wb.Close(False)


try:
    for fn in (ch1, ch3, ch4, ch5, ch6, ch7):
        try:
            fn()
        except Exception as e:
            print(u"★ %s 실행 실패: %s" % (fn.__name__, e))
    ch2()
finally:
    xl.Quit()

print()
bad = 0
for chap, no, ok, got, want, note in results:
    got = getattr(got, "Value", got)
    bad += 0 if ok else 1
    print(u"%s %d장 %d번  계산 %-14s 정답 %-14s %s" % (u"○" if ok else u"✗", chap, no, got, want, note))
print()
print(u"맞지 않는 문항 %d개" % bad)
sys.exit(1 if bad else 0)
