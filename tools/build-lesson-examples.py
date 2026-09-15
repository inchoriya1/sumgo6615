# -*- coding: utf-8 -*-
u"""
예제 파일을 강의 순서대로 다시 깔아 줍니다 (마스터 클래스 1·2장).

  python tools/build-lesson-examples.py            # 보기만
  python tools/build-lesson-examples.py --write    # 실제로 만들기

왜 필요한가
  원본 폴더는 `01-001.xlsx` · `01-A08.xlsx` 처럼 **책 기준 번호**입니다.
  강의를 따라가며 쓰기에는 **어느 파일이 몇 강 것인지 알 수가 없습니다.**
  그래서 강의 번호를 앞에 붙여 새 폴더에 복사합니다.

  · **원본은 그대로 둡니다.** 복사만 합니다.
  · 이름만 바꾸므로 **서식·매크로·시트 구조가 하나도 안 깨집니다.**
  · 한 강의에 파일이 여럿이면 `_1` `_2` 로 나눠 붙입니다.
  · `-완성` 파일은 `_완성` 으로 따라옵니다.

매핑 근거
  파일을 열어 **시트 이름과 원본↔완성 차이**를 보고 정했습니다. 예를 들어 —
  · `02-002` 는 완성본에서 병합이 풀리고 빈 칸이 채워집니다 → 2-4 빈 셀
  · `02-007` 은 완성본이 8열에서 16열로 눕습니다 → 2-8 행열 변환
  · `01-A09` 는 시트가 '상대참조&절대참조'·'혼합참조' 입니다 → 1-9 셀 참조
  확신이 덜한 것은 목록에 **(추정)** 으로 적어 두었습니다. 고치려면 MAP 만 고치면 됩니다.
"""
import argparse
import os
import shutil

SRC = u"진짜쓰는 실무엑셀 예제파일"
DST = u"강의예제"

CHAPTERS = [
    (u"Chapter01", u"1장_시작부터_남다른_실무자의_엑셀_활용", [
        (u"1-01", u"엑셀_필수_설정",        [u"01-A02", u"자동교정"],   False),
        (u"1-02", u"빠른실행_도구모음",      [u"01-A08"],               True),
        (u"1-03", u"엑셀_기본기",           [u"01-A04", u"01-A03"],     False),
        (u"1-04", u"찾기_및_바꾸기",        [u"01-A05"],               False),
        (u"1-05", u"자동_채우기",           [u"01-001"],               False),
        (u"1-06", u"사용자_목록",           [u"01-006"],               False),
        (u"1-07", u"빠른_채우기",           [u"01-002"],               False),
        (u"1-08", u"날짜_시간_데이터",       [u"01-A07", u"01-008"],     False),
        (u"1-09", u"셀_참조_방식",          [u"01-A09"],               False),
        (u"1-10", u"연산자_네_가지",        [u"01-A10"],               False),
        (u"1-11", u"오류와_초록_삼각형",     [u"01-003"],               False),
        (u"1-12", u"필수_단축키_20개",      [u"01-A08"],               True),
    ]),
    (u"Chapter02", u"2장_실무자라면_반드시_알아야_할_엑셀_활용", [
        (u"2-01", u"세로방향_블록쌓기",      [u"02-001"],               False),
        (u"2-02", u"그룹_기능",             [u"02-006", u"02-A02"],     False),
        (u"2-03", u"셀_병합_오류",          [u"02-A01"],               False),
        (u"2-04", u"빈_셀_찾고_채우기",      [u"02-002", u"02-003"],     False),
        (u"2-05", u"한영_자동변환",          [u"02-004"],               False),
        (u"2-06", u"단어_찾아_강조하기",     [u"02-005"],               False),
        (u"2-07", u"숫자_단위_변경",         [u"02-009"],               True),
        (u"2-08", u"행열_변환",             [u"02-007"],               False),
        (u"2-09", u"중복_데이터_제한",       [u"02-008"],               False),
        (u"2-10", u"이름관리자",            [u"02-009"],               True),
        (u"2-11", u"여러_시트_비교",        [u"02-010"],               False),
        (u"2-12", u"텍스트_나누기",         [u"02-011"],               False),
        (u"2-13", u"여러_줄_합치기_나누기",  [u"02-012", u"02-013"],     False),
    ]),
]

# 강의에 붙이지 않고 '추가예제' 로 모아 두는 것들
EXTRA = {
    u"Chapter01": [u"01-004", u"01-005", u"01-006-1", u"01-007", u"01-009",
                   u"01-A01", u"01-A06"],
    u"Chapter02": [],
}


def find(src_dir, stem):
    u"""`01-A08` → 01-A08.xlsx · 01-A08-완성.xlsx 를 찾아 (경로, 완성여부) 로 돌려줍니다."""
    out = []
    for name in sorted(os.listdir(src_dir)):
        base, ext = os.path.splitext(name)
        if ext.lower() not in (".xlsx", ".xlsm", ".xls"):
            continue
        if base == stem:
            out.append((os.path.join(src_dir, name), False, ext))
        elif base == stem + u"-완성":
            out.append((os.path.join(src_dir, name), True, ext))
    return out


def plan():
    jobs, missing = [], []
    for ch_dir, ch_name, lessons in CHAPTERS:
        src_dir = os.path.join(SRC, ch_dir)
        for no, title, stems, guessed in lessons:
            found_any = False
            multi = len(stems) > 1
            for i, stem in enumerate(stems, 1):
                hits = find(src_dir, stem)
                if not hits:
                    continue
                found_any = True
                for path, done, ext in hits:
                    tag = (u"_%d" % i) if multi else u""
                    suffix = u"_완성" if done else u""
                    new = u"%s_%s%s%s%s" % (no, title, tag, suffix, ext)
                    jobs.append((path, os.path.join(DST, ch_name, new), guessed))
            if not found_any:
                missing.append((no, title, stems))
        for stem in EXTRA.get(ch_dir, []):
            for path, done, ext in find(src_dir, stem):
                new = u"%s%s%s" % (stem, u"_완성" if done else u"", ext)
                jobs.append((path, os.path.join(DST, ch_name, u"추가예제", new), False))
    return jobs, missing


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args()

    jobs, missing = plan()
    cur = None
    for src, dst, guessed in jobs:
        folder = os.path.dirname(dst)
        if folder != cur:
            cur = folder
            print(u"\n[%s]" % folder)
        print(u"  %-52s ← %s%s"
              % (os.path.basename(dst), os.path.basename(src),
                 u"   (추정)" if guessed else u""))
        if a.write:
            if not os.path.isdir(folder):
                os.makedirs(folder)
            shutil.copy2(src, dst)

    print(u"\n파일 %d개" % len(jobs))
    if missing:
        print(u"★ 짝이 될 예제가 없는 강의:")
        for no, title, stems in missing:
            print(u"   %s %s (찾은 이름: %s)" % (no, title, u", ".join(stems)))
    if not a.write:
        print(u"\n※ 보기만 했습니다. 실제로 만들려면 --write 를 붙이세요.")


if __name__ == "__main__":
    main()
