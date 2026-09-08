# -*- coding: utf-8 -*-
"""
급여대장의 [간이세액표1] 시트를 새 간이세액표로 갈아 끼웁니다.

  # 먼저 무엇이 바뀌는지 보기만 (파일은 건드리지 않습니다)
  python tools/swap-tax-table.py 급여대장.xlsx 새간이세액표.xlsx

  # 실제로 만들기 (원본은 그대로 두고 새 파일을 만듭니다)
  python tools/swap-tax-table.py 급여대장.xlsx 새간이세액표.xlsx --write --label "2026년 개정"

왜 도구가 필요한가
  표만 붙여넣으면 안 됩니다. 이 파일의 소득세 수식은
      =VLOOKUP( 월급여액/1000, 간이세액표1!$A$3:$N$591, 3+공제대상가족수, TRUE )
  처럼 **행 번호가 박혀** 있습니다(기본정보 AQ열, 192줄).
  새 표의 줄 수가 다르면 범위를 함께 고쳐야 하는데,
  이 도구가 그것까지 같이 합니다.

무엇을 검사하나
  · 구간이 오름차순인가          → 아니면 VLOOKUP(TRUE) 이 엉뚱한 값을 가져옵니다
  · 앞 구간의 '미만' = 다음 '이상' 인가  → 빈틈이 있으면 그 구간이 통째로 잘못 잡힙니다
  · 공제대상가족 열이 11개인가   → 수식이 3+가족수 로 열을 세므로 어긋나면 안 됩니다
  · 가족 수가 늘면 세액이 줄어드는가 → 열 순서가 뒤집혔는지 보는 검사
  · 세액이 음수가 아닌가

무엇을 못 하나
  · **월급여 1천만원 초과분**은 표에 없습니다. 국세청은 그 구간을 별도 산식으로 정합니다.
    지금 수식은 마지막 줄 값에서 멈춥니다(고소득자의 소득세가 실제보다 적게 잡힙니다).
    이 도구는 그것을 고치지 않습니다 — 고치려면 수식 자체를 바꿔야 합니다.
"""
import argparse
import io
import os
import re
import shutil
import sys

import openpyxl
from openpyxl.utils import get_column_letter

SHEET = u"간이세액표1"
ENGINE = u"기본정보"
FAMILY_COLS = 11          # 공제대상가족 1~11명
FIRST_DATA_ROW = 3        # A3 부터 표가 시작합니다


# ─────────────────────────────────────────────────────────────
# 1. 새 표 읽기
# ─────────────────────────────────────────────────────────────
def read_new_table(path):
    """어떤 모양으로 오든 (이상, 미만, 세액 11개) 로 추려 냅니다.

    국세청 파일은 위에 제목·머리글이 여러 줄 붙어 있고 열 위치도 판마다
    조금씩 다릅니다. 그래서 '숫자가 규칙적으로 늘어선 줄'만 골라 씁니다.
    """
    if path.lower().endswith(".csv"):
        rows = []
        for line in io.open(path, encoding="utf-8-sig"):
            rows.append([c.strip() for c in line.rstrip("\n").split(",")])
    else:
        ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
        rows = [[c.value for c in r] for r in ws.iter_rows()]

    def num(v):
        if v is None:
            return None
        if isinstance(v, (int, float)):
            return float(v)
        s = re.sub(r"[,\s원]", "", str(v))
        if not s or not re.match(r"^-?\d+(\.\d+)?$", s):
            return None
        return float(s)

    table = []
    for raw in rows:
        nums = [num(v) for v in raw]
        # 앞쪽 두 개가 숫자이고, 그 뒤로 숫자가 11개 이상 이어지는 줄만
        # '이상 <= 미만' 인 자리를 표의 시작으로 봅니다.
        # 마지막 줄은 '10000 이상 10000 미만' 처럼 둘이 같습니다(맨 위 구간을
        # 받아 주는 줄). 그래서 '<' 이 아니라 '<=' 로 봐야 그 줄이 안 빠집니다.
        start = None
        for i in range(len(nums) - 1):
            if nums[i] is not None and nums[i + 1] is not None and nums[i] <= nums[i + 1]:
                start = i
                break
        if start is None:
            continue
        tail = [x for x in nums[start + 2:] if x is not None]
        # '이상×1000' 같은 보조 열이 끼어 있으면 버립니다
        if tail and abs(tail[0] - nums[start] * 1000) < 1:
            tail = tail[1:]
        if len(tail) < FAMILY_COLS:
            continue
        table.append((nums[start], nums[start + 1], [int(round(x)) for x in tail[:FAMILY_COLS]]))

    if not table:
        raise SystemExit(u"새 표에서 숫자 표를 찾지 못했습니다: " + path)
    return table


# ─────────────────────────────────────────────────────────────
# 2. 검사
# ─────────────────────────────────────────────────────────────
def validate(table):
    bad = []
    prev_hi = None
    for i, (lo, hi, tax) in enumerate(table):
        where = u"%d번째 줄(%s이상 %s미만)" % (i + 1, lo, hi)
        if hi < lo:
            bad.append(u"%s — 미만이 이상보다 작습니다" % where)
        if prev_hi is not None:
            if lo < prev_hi:
                bad.append(u"%s — 앞 구간과 겹칩니다" % where)
            elif lo > prev_hi:
                bad.append(u"%s — 앞 구간과 사이가 비었습니다(앞 미만 %s)" % (where, prev_hi))
        if any(t < 0 for t in tax):
            bad.append(u"%s — 세액에 음수가 있습니다" % where)
        # 가족이 늘면 세액은 줄거나 같아야 합니다
        for k in range(len(tax) - 1):
            if tax[k + 1] > tax[k]:
                bad.append(u"%s — 가족 %d명 세액이 %d명보다 큽니다(열 순서 확인)"
                           % (where, k + 2, k + 1))
                break
        prev_hi = hi
        if len(bad) > 12:
            bad.append(u"… (그 밖에도 더 있습니다)")
            break
    return bad


def sample(wb_or_table, label):
    """눈으로 견줄 표본을 뽑습니다."""
    out = []
    for target in (2000, 3000, 4000, 5000, 6000):
        for lo, hi, tax in wb_or_table:
            if lo <= target < hi:
                out.append((target, lo, hi, tax[:5]))
                break
    return out


def read_old_table(ws):
    old = []
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        lo, hi = ws.cell(r, 1).value, ws.cell(r, 2).value
        if not isinstance(lo, (int, float)) or not isinstance(hi, (int, float)):
            continue
        tax = [ws.cell(r, 4 + k).value or 0 for k in range(FAMILY_COLS)]
        old.append((lo, hi, tax))
    return old


# ─────────────────────────────────────────────────────────────
# 3. 갈아 끼우기
# ─────────────────────────────────────────────────────────────
def swap(book_path, table_path, out_path, label, do_write):
    table = read_new_table(table_path)
    bad = validate(table)

    wb = openpyxl.load_workbook(book_path, data_only=False)
    if SHEET not in wb.sheetnames:
        raise SystemExit(u"[%s] 시트가 없습니다." % SHEET)
    ws = wb[SHEET]
    old = read_old_table(ws)

    old_last = FIRST_DATA_ROW + len(old) - 1
    new_last = FIRST_DATA_ROW + len(table) - 1

    print(u"■ 지금 표      %d줄  (A%d:N%d)" % (len(old), FIRST_DATA_ROW, old_last))
    print(u"   제목 칸    %s" % ws.cell(1, 1).value)
    print(u"■ 새 표        %d줄  (A%d:N%d)" % (len(table), FIRST_DATA_ROW, new_last))
    print(u"   구간        %s ~ %s 천원" % (table[0][0], table[-1][1]))

    # 맨 위 구간을 받아 주는 줄(이상 == 미만)이 있어야, 표 끝을 넘는 월급여도
    # 마지막 값으로 잡힙니다. 없으면 그 위 구간 값이 잡혀 세액이 달라집니다.
    top_lo, top_hi, top_tax = table[-1]
    if top_lo == top_hi:
        print(u"   맨 윗줄     %s 이상은 모두 %s원(가족 1명)" % (top_lo, format(top_tax[0], ",")))
    else:
        print(u"   ⚠ 맨 윗줄이 '이상 == 미만' 형태가 아닙니다.")
        print(u"     %s 천원을 넘는 월급여는 그 아래 구간 값으로 잡힙니다 — 확인하세요." % top_hi)

    print(u"\n■ 표본 대조 (월급여 천원 → 가족수 1~5 세액)")
    o = dict((t, (lo, hi, tax)) for t, lo, hi, tax in sample(old, "old"))
    n = dict((t, (lo, hi, tax)) for t, lo, hi, tax in sample(table, "new"))
    for t in (2000, 3000, 4000, 5000, 6000):
        print(u"   %5d  지금 %s" % (t, o.get(t, ("-",))[2] if t in o else "-"))
        print(u"          새로 %s" % (n.get(t, ("-",))[2] if t in n else "-",))

    if bad:
        print(u"\n■ 검사에서 걸린 것 — 이대로 쓰면 안 됩니다")
        for b in bad:
            print(u"   ✗ " + b)
        raise SystemExit(u"\n중단했습니다. 새 표 파일을 확인하세요.")
    print(u"\n■ 검사 통과 — 오름차순·빈틈·열 수·가족수별 감소 모두 정상")

    # 소득세 수식의 범위가 몇 곳에 박혀 있는지
    eng = wb[ENGINE]
    pat = re.compile(re.escape(SHEET) + r"!\$A\$" + str(FIRST_DATA_ROW) + r":\$N\$(\d+)")
    hits = []
    for row in eng.iter_rows():
        for c in row:
            if isinstance(c.value, str) and pat.search(c.value):
                hits.append(c)
    print(u"■ 소득세 수식 %d곳이 이 표의 범위를 가리킵니다 (%s 시트)" % (len(hits), ENGINE))
    if len(old) != len(table):
        print(u"   줄 수가 달라졌으므로 $N$%d → $N$%d 로 함께 고칩니다."
              % (old_last, new_last))

    if not do_write:
        print(u"\n※ 보기만 했습니다. 실제로 만들려면 --write 를 붙이세요.")
        return

    # ── 여기서부터 실제 수정 ──
    for r in range(FIRST_DATA_ROW, max(old_last, new_last) + 1):
        for c in range(1, 15):
            ws.cell(r, c).value = None
    for i, (lo, hi, tax) in enumerate(table):
        r = FIRST_DATA_ROW + i
        ws.cell(r, 1).value = lo
        ws.cell(r, 2).value = hi
        ws.cell(r, 3).value = "=A%d*1000" % r
        for k, t in enumerate(tax):
            ws.cell(r, 4 + k).value = t

    if label:
        base = re.sub(r"\s*\d{4}\s*$", "", str(ws.cell(1, 1).value or u"월급여액(천원)"))
        ws.cell(1, 1).value = u"%s %s" % (base.rstrip(), label)

    for c in hits:
        c.value = pat.sub(
            lambda m: u"%s!$A$%d:$N$%d" % (SHEET, FIRST_DATA_ROW, new_last), c.value)

    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)
    print(u"\n만들었습니다: %s" % out_path)
    print(u"원본은 그대로 두었습니다: %s" % book_path)
    print(u"\n※ 엑셀에서 열어 [전체임금대장]의 소득세가 달라졌는지 꼭 확인하세요.")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument("book", help=u"급여대장 .xlsx")
    ap.add_argument("table", help=u"새 간이세액표 (.xlsx 또는 .csv)")
    ap.add_argument("--out", default=None, help=u"만들 파일 이름")
    ap.add_argument("--label", default=None, help=u'제목 칸에 넣을 판 이름 (예: "2026년 개정")')
    ap.add_argument("--write", action="store_true", help=u"실제로 만들기")
    a = ap.parse_args()
    out = a.out or re.sub(r"\.xlsx$", "", a.book) + u"_세액표교체.xlsx"
    swap(a.book, a.table, out, a.label, a.write)
