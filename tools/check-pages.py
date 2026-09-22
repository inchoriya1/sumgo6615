# -*- coding: utf-8 -*-
u"""강의안 HTML을 훑어 **차례와 본문이 어긋난 곳**을 찾습니다.

  python tools/check-pages.py excel-m3.html excel-m4.html
  python tools/check-pages.py                # excel-m*.html · excel-b*.html 전부

보는 것 넷
  · 차례의 `#앵커` 가 **실제 section id** 와 하나씩 맞는가 (순서까지)
  · 페이지 안의 `href="…html"` 이 **있는 파일** 을 가리키는가
  · `<code>` 안에 적어 둔 **강의 번호** 가 강의예제 폴더에 있는가
    (본문에서는 `1-03_기본기_1` 처럼 줄여 쓰므로 **번호만** 봅니다)
  · 태그가 짝이 맞는가 (여는 수 = 닫는 수)
"""
import io
import os
import re
import sys

BLOCK = ("section", "div", "table", "figure", "ol", "ul", "nav", "main", "header")


def check(path):
    s = io.open(path, encoding="utf-8").read()
    bad = []

    toc = re.findall(r'<li><a href="#([^"]+)"', s)
    ids = re.findall(r'<section id="([^"]+)"', s)
    if toc != ids:
        bad.append(u"차례 %d개 ↔ 섹션 %d개 — 차례 %s / 섹션 %s"
                   % (len(toc), len(ids), toc, ids))

    for href in set(re.findall(r'href="([^"#]+\.html)"', s)):
        if not os.path.exists(href):
            bad.append(u"없는 쪽으로 가는 링크: %s" % href)

    for tag in BLOCK:
        o = len(re.findall(r"<%s[ >]" % tag, s))
        c = len(re.findall(r"</%s>" % tag, s))
        if o != c:
            bad.append(u"<%s> %d개인데 </%s> 는 %d개" % (tag, o, tag, c))

    lessons = set()
    for root, _, files in os.walk(u"강의예제"):
        for f in files:
            m = re.match(r"^([0-9AB]-\d\d)_", f)
            if m:
                lessons.add(m.group(1))
    for code in set(re.findall(r"<code>([^<]+)</code>", s)):
        m = re.match(r"^([0-9AB]-\d\d)[_.]", code)
        if m and m.group(1) not in lessons:
            bad.append(u"그런 강의 번호가 없음: %s" % code)

    return bad


def main():
    args = sys.argv[1:]
    if not args:
        args = sorted(f for f in os.listdir(".")
                      if re.match(r"^excel-[mb]\d+\.html$", f))
    total = 0
    for path in args:
        bad = check(path)
        total += len(bad)
        print(u"%-16s %s" % (path, u"괜찮습니다" if not bad else u"★ %d곳" % len(bad)))
        for b in bad:
            print(u"    %s" % b)
    print(u"\n모두 %d곳" % total)
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
