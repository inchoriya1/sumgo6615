# -*- coding: utf-8 -*-
"""
근태원본이 사람 눈에 제대로 보이게 서식과 단위를 고칩니다.

  python tools/fix-formats.py 급여대장.xlsx            # 보기만
  python tools/fix-formats.py 급여대장.xlsx --write    # 새 파일 만들기

무엇이 문제였나
  ① **근무일자가 `46235` 로 보입니다.**
     엑셀은 날짜를 '1900년부터 며칠째'라는 숫자로 담습니다.
     날짜 서식을 입혀야 사람이 읽는 모양으로 보입니다.
  ② **기본·연장이 `0.3333` 으로 보입니다.**
     원래 파일이 시간을 '하루의 몇 분의 몇'(8시간 = 8/24 = 0.3333)으로
     담고 있었고, 그것을 그대로 따라 만들었기 때문입니다.
  ③ **출근·퇴근도 붙여넣으면 소수로 보입니다.** 같은 이유입니다.

어떻게 고치나
  · 근무일자 → `yyyy-mm-dd (aaa)` — 날짜 옆에 요일까지 보입니다.
  · 출근·퇴근 등 시각 칸 → `h:mm`
  · 곁들여, **비어 있던 파생 열(P~W)도 채웁니다.** 원본이 4001행까지만
    가지고 있어서, 자리를 5,952줄로 늘린 뒤 **2027년 5월부터 날짜·월키가 없어
    급여가 통째로 0** 이 되는 상태였습니다.
  · **기본·연장은 서식만 바꾸지 않고 값 자체를 '시간'으로** 바꿉니다.
    `8/24` 대신 그냥 `8`. 눈으로 보기도 쉽고, 다른 데 옮겨 쓰기도 쉽습니다.
    이 두 열을 쓰는 곳은 같은 시트의 S·T(기본(H)·연장(H)) 뿐이어서,
    거기서 ×24 하던 것만 빼면 됩니다. **나머지는 하나도 안 바뀝니다.**
"""
import argparse
import re

import openpyxl

RAW = u"근태원본"
FIRST, LAST = 2, 5953
DATE_FMT = u"yyyy-mm-dd (aaa)"
TIME_FMT = u"h:mm"
HOUR_FMT = u"General"     # 0.## 로 하면 8 이 "8." 로 보입니다

TIME_COLS = (4, 5, 6, 7, 8, 9, 10, 13)      # D 출근 · E 퇴근 · F~J · M 휴일
HOUR_COLS = (11, 12, 15, 19, 20)            # K 기본 · L 연장 · O 지각차감 · S · T

# P~W 파생 열. 원본이 4001행까지만 가지고 있어서, 늘어난 줄에 채워 넣어야 합니다.
DERIVED = [
    (16, u'=IF($B{r}="","",IF(ISNUMBER($B{r}),$B{r},'
         u'IFERROR(DATEVALUE(SUBSTITUTE($B{r},"/","-")),"")))'),
    (17, u'=IF($P{r}="","",TEXT($P{r},"yyyy-mm"))'),
    (18, u'=IF($P{r}="","",TEXT($P{r},"aaa"))'),
    (21, u'=IF(N($T{r})=0,"",IF($C{r}="평일","연장",'
         u'IF(AND($R{r}="토",$C{r}="휴일"),"토요","특근")))'),
    (22, u'=IF($P{r}="","",$Q{r}&"|"&$A{r})'),
    (23, u'=IF($P{r}="","",$A{r}&"|"&TEXT($P{r},"yyyy-mm-dd"))'),
]


def run(path, out_path, do_write):
    wb = openpyxl.load_workbook(path, data_only=False)
    if RAW not in wb.sheetnames:
        raise SystemExit(u"[%s] 시트가 없습니다." % RAW)
    raw = wb[RAW]

    r = FIRST
    print(u"■ 지금 (%d행 기준)" % r)
    for c, nm in ((2, u"근무일자"), (4, u"출근"), (5, u"퇴근"),
                  (11, u"기본"), (12, u"연장"), (19, u"기본(H)")):
        print(u"   %-8s 서식 %-16s" % (nm, raw.cell(r, c).number_format))
    print(u"\n■ 바꾼 뒤")
    print(u"   근무일자      %s   ← 46235 대신 날짜와 요일" % DATE_FMT)
    print(u"   출근·퇴근 등   %s" % TIME_FMT)
    print(u"   기본·연장      값 자체를 '시간' 으로 (8/24 → 8)")
    print(u"   기본(H)·연장(H) 에서 ×24 를 뺍니다 — 결과 값은 그대로입니다")

    if not do_write:
        print(u"\n※ 보기만 했습니다. 실제로 만들려면 --write 를 붙이세요.")
        return

    n_fmt = n_k = n_s = n_d = 0
    raw.column_dimensions["B"].width = 17      # yyyy-mm-dd (요일) 이 다 보이게
    for row in range(FIRST, LAST + 1):
        raw.cell(row, 2).number_format = DATE_FMT
        raw.cell(row, 16).number_format = u"yyyy-mm-dd"
        for c in TIME_COLS:
            raw.cell(row, c).number_format = TIME_FMT
        for c in HOUR_COLS:
            raw.cell(row, c).number_format = HOUR_FMT
        n_fmt += 1

        # K 기본 · L 연장 — '하루의 몇 분의 몇' 이 아니라 '시간' 으로.
        # 문자열을 잘라 붙이면 어긋나기 쉬워서, 수식을 통째로 다시 씁니다.
        if isinstance(raw.cell(row, 11).value, str):
            raw.cell(row, 11).value = (
                u'=IF(OR($B{r}="",$D{r}="",$E{r}=""),"",'
                u'MAX(0,IF($C{r}="평일",8,0)-N($O{r})))'.format(r=row))
            raw.cell(row, 12).value = (
                u'=IF(OR($B{r}="",$D{r}="",$E{r}=""),"",'
                u'FLOOR(MAX(0,$E{r}-IF($C{r}="평일",TIME(17,30,0),TIME(8,30,0)))*48,1)/2)'
                .format(r=row))
            n_k += 2

        # P~W 파생 열 — 원래 파일이 4001행까지만 있었는데 자리를 5953줄로
        # 늘리면서 여기를 안 늘렸습니다. 그대로 두면 2027년 5월부터
        # 날짜·월키가 없어 **급여가 통째로 0** 이 됩니다. 채워 넣습니다.
        for c, f in DERIVED:
            if raw.cell(row, c).value in (None, ""):
                raw.cell(row, c).value = f.format(r=row)
                n_d += 1

        # S 기본(H) · T 연장(H) — K·L 이 이미 시간이므로 ×24 가 필요 없습니다
        for c, src in ((19, "K"), (20, "L")):
            raw.cell(row, c).value = (
                u'=IF($P%d="","",ROUND(N($%s%d),2))' % (row, src, row))
            n_s += 1

    print(u"\n서식 %d줄 · 기본·연장 %d칸 · 기본(H)·연장(H) %d칸을 고쳤습니다."
          % (n_fmt, n_k, n_s))
    if n_d:
        print(u"★ 비어 있던 파생 열(날짜·연월·요일·구분·월키·일키) %d칸을 채웠습니다." % n_d)
        print(u"   원본이 4001행까지만 가지고 있어서, 그대로 두면")
        print(u"   2027년 5월부터 급여가 통째로 0 이 되는 상태였습니다.")
    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)
    print(u"만들었습니다: %s" % out_path)
    print(u"원본은 그대로 두었습니다: %s" % path)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("book")
    ap.add_argument("--out", default=None)
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args()
    out = a.out or re.sub(r"\.xlsx$", "", a.book) + u"_서식.xlsx"
    run(a.book, out, a.write)
