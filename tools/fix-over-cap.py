# -*- coding: utf-8 -*-
"""
간이세액표 상한(1천만원)을 넘는 급여의 소득세를 '조용히 틀리지 않게' 고칩니다.

  # 무엇이 바뀌는지 보기만
  python tools/fix-over-cap.py 급여대장.xlsx

  # 실제로 만들기 (원본은 그대로, 새 파일이 생깁니다)
  python tools/fix-over-cap.py 급여대장.xlsx --write

무엇이 문제인가
  소득세는 간이세액표에서 VLOOKUP(..., TRUE) 로 찾아옵니다.
  표는 월급여액 1천만원에서 끝나므로, 그보다 많이 받으면
  **마지막 줄 값에서 멈춥니다.** 오류도 안 뜨고 빨간 표시도 없습니다.
  1,200만원을 받아도 1천만원일 때와 똑같은 세금을 뗍니다.

왜 산식을 넣지 않았나
  국세청은 1천만원 초과분을 **별도 산식**으로 정합니다.
  그 산식을 공식 자료로 확인하지 못했고, 세금 산식은 짐작으로 넣으면
  안 되는 것이라 **비워 두었습니다.**

그래서 무엇을 하나 — 두 가지
  ① [월별입력] 에 **소득세직접입력** 칸을 만듭니다.
     세무사무소에서 받은 정확한 세액을 그 달 그 사람 칸에 적으면
     표 대신 그 값을 씁니다. (이 파일이 이미 여덟 군데에서 쓰고 있는
     "비었으면 기본값, 적혀 있으면 그 값" 과 똑같은 방식입니다.)
  ② [기본정보] 맨 오른쪽 **성명확인** 열에 경고를 띄웁니다.
     월급여액이 표 상한을 넘었는데 직접 입력이 비어 있으면
     "★소득세 확인필요" 가 뜹니다. 이 열은 매달 검산할 때 보는 자리입니다.

  즉 **모르는 채로 넘어가는 일이 없게** 만드는 것이 이 도구가 하는 일입니다.
"""
import argparse
import re

import openpyxl

ENGINE = u"기본정보"          # 계산 엔진
INPUT = u"월별입력"           # 그 달만 다른 값
TABLE = u"간이세액표1"        # 간이세액표
ROW_FROM, ROW_TO = 2, 193     # 24개월 x 8명
COL_TAX = 43                  # AQ 소득세
COL_PAY = 36                  # AJ 월급여액
COL_NAME = 2                  # B 성명
COL_CHECK = 56                # BD 성명확인
COL_NEW = 24                  # 월별입력 X — 새로 만들 칸
NEW_HEADER = u"소득세직접입력"
CAP_NAME = u"세액표상한"


def col_letter(n):
    return openpyxl.utils.get_column_letter(n)


def inspect(wb):
    """지금 상태를 확인하고, 고칠 수 있는 파일인지 봅니다."""
    for name in (ENGINE, INPUT, TABLE):
        if name not in wb.sheetnames:
            raise SystemExit(u"[%s] 시트가 없습니다." % name)
    mi = wb[INPUT]
    if mi.cell(1, COL_NEW).value not in (None, "", NEW_HEADER):
        raise SystemExit(u"[%s] %s1 칸이 이미 쓰이고 있습니다: %s"
                         % (INPUT, col_letter(COL_NEW), mi.cell(1, COL_NEW).value))
    ws = wb[TABLE]
    cap = None
    for r in range(3, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, (int, float)):
            cap = v
    return cap


def build(path, out_path, do_write):
    wb = openpyxl.load_workbook(path, data_only=False)
    cap = inspect(wb)
    eng, mi = wb[ENGINE], wb[INPUT]

    print(u"■ 간이세액표 상한 : %s 천원 (= %s 원)"
          % (cap, format(int(cap * 1000), ",")))
    print(u"■ 지금 소득세 수식 : %s" % eng.cell(ROW_FROM, COL_TAX).value)
    print(u"■ 지금 확인 열     : %s" % eng.cell(ROW_FROM, COL_CHECK).value)

    # 지금 값으로 상한을 넘는 줄이 있는지 (계산값 파일에서 확인)
    vals = openpyxl.load_workbook(path, data_only=True)[ENGINE]
    over, near = [], []
    for r in range(ROW_FROM, ROW_TO + 1):
        nm = vals.cell(r, COL_NAME).value
        pay = vals.cell(r, COL_PAY).value
        ym = vals.cell(r, 54).value
        if not nm or not isinstance(pay, (int, float)):
            continue
        if pay / 1000.0 > cap:
            over.append((ym, nm, pay))
        elif pay / 1000.0 == cap:
            near.append((ym, nm, pay))
    print(u"\n■ 지금 상한을 **넘는** 줄 : %d개" % len(over))
    for ym, nm, pay in over[:10]:
        print(u"     %s %s %s원" % (ym, nm, format(int(pay), ",")))
    print(u"■ 지금 상한에 **딱 걸린** 줄 : %d개  (표 마지막 줄이 정확히 맞는 자리)" % len(near))
    for ym, nm, pay in near[:10]:
        print(u"     %s %s %s원  ← 상여가 한 번 붙으면 바로 넘어갑니다" % (ym, nm, format(int(pay), ",")))

    tax_ref = u"%s!$%s%d" % (INPUT, col_letter(COL_NEW), ROW_FROM)
    print(u"\n■ 고친 뒤 소득세 수식 (2행 기준)")
    print(u"     =IF(%s<>\"\", %s, 지금수식)" % (tax_ref, tax_ref))
    print(u"■ 고친 뒤 확인 열 (2행 기준)")
    print(u"     빈 슬롯이면 '(빈 슬롯)', 상한 초과인데 직접입력이 비면 '★소득세 확인필요'")

    if not do_write:
        print(u"\n※ 보기만 했습니다. 실제로 만들려면 --write 를 붙이세요.")
        return

    # ── ① 월별입력에 소득세직접입력 칸 ──
    mi.cell(1, COL_NEW).value = NEW_HEADER

    # ── ② 이름 정의: 표 상한 (표를 갈아 끼워도 따라옵니다) ──
    defn = openpyxl.workbook.defined_name.DefinedName(
        CAP_NAME, attr_text=u"MAX(%s!$A$3:$A$100000)" % TABLE)
    if CAP_NAME in wb.defined_names:
        del wb.defined_names[CAP_NAME]
    wb.defined_names[CAP_NAME] = defn

    # ── ③ 소득세 수식 · 확인 열 ──
    changed_tax = changed_chk = 0
    for r in range(ROW_FROM, ROW_TO + 1):
        cell = eng.cell(r, COL_TAX)
        cur = cell.value
        if isinstance(cur, str) and cur.startswith("=") and NEW_HEADER not in cur:
            ref = u"%s!$%s%d" % (INPUT, col_letter(COL_NEW), r)
            cell.value = u'=IF(%s<>"",%s,%s)' % (ref, ref, cur[1:])
            changed_tax += 1

        chk = eng.cell(r, COL_CHECK)
        ref = u"%s!$%s%d" % (INPUT, col_letter(COL_NEW), r)
        chk.value = (
            u'=IF($B{r}="","(빈 슬롯)",'
            u'IF(AND($AJ{r}/1000>{cap},{ref}=""),'
            u'"★소득세 확인필요 — 월급여가 간이세액표 상한을 넘었습니다",""))'
        ).format(r=r, cap=CAP_NAME, ref=ref)
        changed_chk += 1

    print(u"\n소득세 수식 %d곳, 확인 열 %d곳을 고쳤습니다." % (changed_tax, changed_chk))
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
    out = a.out or re.sub(r"\.xlsx$", "", a.book) + u"_상한초과보완.xlsx"
    build(a.book, out, a.write)
