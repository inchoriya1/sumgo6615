# -*- coding: utf-8 -*-
u"""
예제 파일을 강의 순서대로 다시 깔아 줍니다 (마스터 클래스 1~8장 · 부록).

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
    (u"Chapter03", u"3장_보고서가_달라지는_서식_활용법", [
        (u"3-01", u"셀_서식의_원리",         [u"03-001"],               False),
        (u"3-02", u"사용자_지정_표시_형식",   [u"03-002"],               False),
        (u"3-03", u"보고서_작성_규칙",       [u"03-003"],               False),
        (u"3-04", u"조건부_서식",           [u"03-004"],               False),
        (u"3-05", u"데이터_막대",           [u"03-005"],               False),
        (u"3-06", u"아이콘_집합",           [u"03-006"],               False),
        (u"3-07", u"스파크라인",            [u"03-007"],               False),
    ]),
    (u"Chapter04", u"4장_완성한_보고서_공유_및_출력하기", [
        (u"4-01", u"틀_고정과_창_나누기",     [u"04-001"],               False),
        (u"4-02", u"인쇄_영역과_눈금선",      [u"04-002"],               False),
        (u"4-03", u"시트_복사와_수식",       [u"04-003"],               True),
        (u"4-04", u"외부_통합문서_참조",      [u"04-004"],               False),
        (u"4-05", u"공공데이터_받아_쓰기",     [u"04-005"],               True),
        (u"4-06", u"열_순서_바꾸기",         [u"04-006"],               False),
        (u"4-07", u"파일_크기_줄이기",        [u"04-007"],               False),
        (u"4-08", u"데이터_유효성_검사",      [u"04-008"],               False),
        (u"4-09", u"그림_다루기",           [u"04-009"],               True),
        (u"4-10", u"통합_문서_보호",         [u"04-010"],               False),
        (u"4-11", u"인쇄_배율과_용지_방향",    [u"04-011"],               False),
        (u"4-12", u"인쇄_제목_반복",         [u"04-012"],               False),
        (u"4-13", u"긴_표_인쇄",            [u"04-013"],               False),
        (u"4-14", u"머리글과_바닥글",        [u"04-014"],               False),
        (u"4-15", u"여러_시트_한번에",        [u"04-015"],               False),
    ]),
    (u"Chapter05", u"5장_데이터_정리부터_데이터_필터링까지", [
        (u"5-01", u"데이터_관리_기본_규칙",    [u"05-001"],               False),
        (u"5-02", u"여러_시트_취합",         [u"05-002"],               False),
        (u"5-03", u"정렬_기초",             [u"05-003"],               False),
        (u"5-04", u"사용자_지정_목록_정렬",    [u"05-004"],               False),
        (u"5-05", u"한_칸에_뭉친_값",        [u"05-005"],               True),
        (u"5-06", u"자동_필터",             [u"05-006"],               False),
        (u"5-07", u"숫자_필터와_색_필터",     [u"05-007"],               True),
        (u"5-08", u"SUBTOTAL_함수",        [u"05-008"],               False),
        (u"5-09", u"고급_필터_기초",         [u"05-009"],               False),
        (u"5-10", u"고급_필터_조건_만들기",    [u"05-010"],               False),
        (u"5-11", u"고급_필터_다른_시트로",    [u"05-011"],               False),
        (u"5-12", u"파워_쿼리_열_피벗_해제",   [u"05-012"],               False),
    ]),
    (u"Chapter06", u"6장_자동화를_위한_표와_피벗_테이블", [
        (u"6-01", u"엑셀_표_기능",          [u"06-001"],               False),
        (u"6-02", u"표의_구조적_참조",        [u"06-002"],               True),
        (u"6-03", u"피벗을_위한_데이터_형태",   [u"06-003"],               False),
        (u"6-04", u"피벗_테이블_만들기",       [u"06-004"],               False),
        (u"6-05", u"값_요약과_고유_개수",      [u"06-005"],               False),
        (u"6-06", u"피벗_새로_고침",         [u"06-006"],               True),
        (u"6-07", u"그룹_숫자로_묶기",        [u"06-007"],               False),
        (u"6-08", u"그룹_날짜로_묶기",        [u"06-008"],               False),
        (u"6-09", u"보고서_레이아웃과_부분합",  [u"06-009"],               False),
        (u"6-10", u"보고서_필터_페이지_표시",   [u"06-010"],               False),
        (u"6-11", u"계산_필드와_계산_항목",    [u"06-011"],               False),
        (u"6-12", u"슬라이서",              [u"06-012"],               False),
        (u"6-13", u"시간_표시_막대",         [u"06-013"],               False),
        (u"6-14", u"피벗_차트",             [u"06-014"],               False),
        (u"6-15", u"피벗_대시보드",          [u"06-015"],               False),
    ]),
    (u"Chapter07", u"7장_기본과_필수_함수_익히기", [
        (u"7-01", u"함수의_기본",           [u"07-001"],               False),
        (u"7-02", u"최대_최소와_순위",        [u"07-002"],               False),
        (u"7-03", u"IF_함수",              [u"07-003"],               False),
        (u"7-04", u"VLOOKUP과_IFERROR",    [u"07-004"],               False),
        (u"7-05", u"SUMIF와_SUMIFS",       [u"07-005"],               False),
        (u"7-06", u"LEFT_MID_RIGHT",       [u"07-006"],               False),
        (u"7-07", u"LEN과_TRIM",           [u"07-007"],               False),
        (u"7-08", u"FIND로_잘라내기",        [u"07-008"],               False),
        (u"7-09", u"SUBSTITUTE와_줄바꿈",    [u"07-009"],               False),
        (u"7-10", u"TEXT_함수",            [u"07-010"],               False),
        (u"7-11", u"DATE와_날짜_계산",       [u"07-011"],               False),
        (u"7-12", u"DATEDIF_사용기간",       [u"07-012"],               False),
        (u"7-13", u"INDIRECT_시트_바꿔가며",  [u"07-013"],               False),
        (u"7-14", u"배열_이해하기",          [u"07-014"],               False),
        (u"7-15", u"FILTER_함수",          [u"07-015"],               False),
        (u"7-16", u"XLOOKUP_함수",         [u"07-016"],               False),
        (u"7-17", u"UNIQUE와_SORT",        [u"07-017"],               False),
        (u"7-18", u"AND_OR과_중첩_IF",      [u"07-018"],               False),
        (u"7-19", u"여러_단어_포함_검사",      [u"07-019"],               False),
        (u"7-20", u"INDEX와_MATCH",        [u"07-020"],               False),
        (u"7-21", u"다중_조건_조회",         [u"07-021", u"07-023"],     False),
        (u"7-22", u"OFFSET_동적_범위",       [u"07-022"],               False),
    ]),
    (u"Chapter08", u"8장_실무에서_필요한_데이터_시각화", [
        (u"8-01", u"막대로_만드는_표",        [u"08-001"],               False),
        (u"8-02", u"실무_차트_다섯_가지",      [u"08-002"],               False),
        (u"8-03", u"혼합_차트",             [u"08-003"],               False),
        (u"8-04", u"보조_축",               [u"08-004"],               False),
        (u"8-05", u"간트_차트",             [u"08-005"],               False),
    ]),
    (u"한 걸음 더(부록)", u"부록_한_걸음_더", [
        (u"A-01", u"그림판으로_만드는_차트",   [u"A-001"],                False),
        (u"A-02", u"데이터_모델_피벗",        [u"A-002"],                False),
        (u"A-03", u"시계열_예측",           [u"A-003"],                False),
    ]),
]

# 강의에 붙이지 않고 '추가예제' 로 모아 두는 것들
EXTRA = {
    u"Chapter01": [u"01-004", u"01-005", u"01-006-1", u"01-007", u"01-009",
                   u"01-A01", u"01-A06"],
    u"Chapter02": [],
    u"Chapter03": [],
    u"Chapter04": [u"04-A01", u"04-A02", u"04-A03", u"04-A04", u"04-A05"],
    u"Chapter05": [u"05-A01", u"05-A02", u"05-A03"],
    u"Chapter06": [u"06-A01"],
    u"Chapter07": [u"07-A01", u"07-A02"],
    u"Chapter08": [u"08-A01"],
}

# `04-007-1.xlsb` 처럼 **번호가 더 붙은 짝**이 있는 강의.
# 여기 적은 것만 `-1` `-2` 파일을 함께 가져와 이름 끝에 `_1` `_2` 를 붙입니다.
# (안 적으면 `01-006` 이 `01-006-1` 까지 끌어와 추가예제와 겹칩니다.)
NUMBERED = frozenset([u"04-007", u"04-014", u"06-006"])


EXTS = (".xlsx", ".xlsm", ".xls", ".xlsb", ".png")


def find(src_dir, stem):
    u"""`01-A08` → 01-A08.xlsx · 01-A08-완성.xlsx 를 찾아
    (경로, 완성여부, 확장자, 꼬리) 로 돌려줍니다.

    `NUMBERED` 에 적은 강의는 `04-007-1.xlsb` 같은 **번호 짝**도 함께 가져오고,
    꼬리에 `_1` `_2` 가 붙습니다.
    """
    out = []
    numbered = stem in NUMBERED
    for name in sorted(os.listdir(src_dir)):
        base, ext = os.path.splitext(name)
        if ext.lower() not in EXTS:
            continue
        if base == stem:
            out.append((os.path.join(src_dir, name), False, ext, u""))
        elif base == stem + u"-완성":
            out.append((os.path.join(src_dir, name), True, ext, u""))
        elif numbered and base.startswith(stem + u"-") \
                and base[len(stem) + 1:].isdigit():
            out.append((os.path.join(src_dir, name), False, ext,
                        u"_" + base[len(stem) + 1:]))
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
                for path, done, ext, tail in hits:
                    tag = (u"_%d" % i) if multi else u""
                    suffix = u"_완성" if done else u""
                    new = u"%s_%s%s%s%s%s" % (no, title, tag, tail, suffix, ext)
                    jobs.append((path, os.path.join(DST, ch_name, new), guessed))
            if not found_any:
                missing.append((no, title, stems))
        for stem in EXTRA.get(ch_dir, []):
            for path, done, ext, tail in find(src_dir, stem):
                new = u"%s%s%s%s" % (stem, tail, u"_완성" if done else u"", ext)
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
