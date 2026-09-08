# -*- coding: utf-8 -*-
"""
암호를 걸어 배포하는 페이지를 만듭니다.

  python tools/build-locked-page.py excel-4.src.html excel-4.html 6615

왜 이렇게 하나
  자바스크립트로 "암호가 맞으면 보여 주기"만 하면 아무 소용이 없습니다.
  글이 이미 페이지 안에 들어 있어서, 소스 보기 한 번이면 다 읽힙니다.
  그래서 여기서는 **본문을 실제로 암호화**합니다.
  배포되는 파일에는 알아볼 수 없는 덩어리만 들어가고,
  암호를 맞게 넣어야 브라우저가 그것을 풀어서 화면에 붙입니다.

방식
  PBKDF2-SHA256 (31만 번) 으로 암호에서 열쇠를 만들고,
  AES-256-GCM 으로 본문을 잠급니다. 브라우저의 Web Crypto 로 그대로 풉니다.

한계 — 반드시 알고 쓰세요
  · 암호가 짧으면(네 자리 같은) 작정하고 덤비는 사람은 뚫을 수 있습니다.
    검색엔진·크롤러·지나가는 사람을 막는 수준입니다.
  · 더 길게 하려면 이 명령의 마지막 값만 바꿔 다시 만드세요.
  · file:// 로 그냥 열면 브라우저가 Web Crypto 를 막아 풀리지 않습니다.
    내 컴퓨터에서 볼 때는 원본(*.src.html)을 여세요.
"""
import base64
import io
import os
import re
import sys

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

ITERATIONS = 310000


def build(src_path, out_path, password):
    src = io.open(src_path, encoding="utf-8").read()

    # <html> 의 속성, <title>, 그리고 본문(.shell 통째)을 뽑아 씁니다.
    html_attrs = re.search(r"<html([^>]*)>", src).group(1).strip()
    title = re.search(r"<title>(.*?)</title>", src, re.S).group(1).strip()
    body = re.search(r'(<div class="shell.*?)\n<script src="script\.js">', src, re.S)
    if not body:
        raise SystemExit("본문(.shell ~ script.js 앞)을 찾지 못했습니다.")
    secret = body.group(1).strip()

    salt = os.urandom(16)
    iv = os.urandom(12)
    key = PBKDF2HMAC(
        algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITERATIONS
    ).derive(password.encode("utf-8"))
    blob = AESGCM(key).encrypt(iv, secret.encode("utf-8"), None)

    b64 = lambda raw: base64.b64encode(raw).decode("ascii")
    page = SHELL.format(
        html_attrs=html_attrs,
        title=title,
        salt=b64(salt),
        iv=b64(iv),
        data=b64(blob),
        iters=ITERATIONS,
    )
    io.open(out_path, "w", encoding="utf-8", newline="\n").write(page)

    # 잠근 글이 결과 파일에 그대로 남지 않았는지 스스로 확인합니다.
    # 제목은 자물쇠 화면에 그대로 남습니다(탭에 보여야 하니까).
    # 그래서 제목까지 넣은 상태를 '남아도 되는 것'으로 봅니다.
    # → 제목에는 실명을 쓰지 마세요. 아래 출력으로 매번 확인하세요.
    written = io.open(out_path, encoding="utf-8").read()
    allowed = SHELL.format(
        html_attrs=html_attrs, title=title, salt="", iv="", data="", iters=0
    )
    leaked = find_leaks(secret, written, allowed)
    if leaked:
        raise SystemExit(u"본문이 결과 파일에 남아 있습니다: " + u", ".join(leaked[:8]))

    print(u"만들었습니다: %s  (%.0f KB, 원본 %.0f KB)"
          % (out_path, len(page) / 1024.0, len(src) / 1024.0))
    print(u"암호 없이 보이는 제목 → %s" % title)


def find_leaks(secret, written, allowed):
    """잠갔어야 할 글이 결과 파일에 남아 있으면 그 낱말들을 돌려줍니다.

    지킬 말(실명 등)을 이 파일에 적어 두지 않습니다 — 이 파일도 공개
    저장소에 올라가므로, 적어 두면 그것이 곧 유출입니다.
    대신 **원본에서 뽑아** 검사합니다. 두 글자 이상 이어진 한글을 모두 모아
    결과 파일에 남아 있는지 봅니다. 암호문은 base64(영문·숫자)라 한글이
    나올 수 없으므로, 정상이라면 걸리는 것은 자물쇠 화면(allowed)이
    쓰는 낱말뿐입니다. 그래서 그것들은 빼고 봅니다.
    """
    return sorted(
        p for p in set(re.findall(r"[가-힣]{2,}", secret))
        if p in written and p not in allowed
    )


SHELL = u"""<!doctype html>
<html{html_attrs}>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>{title}</title>
<link rel="stylesheet" href="style.css">
<script>
(function(){{try{{
  var d=document.documentElement;
  var t=localStorage.getItem('sumgo-theme');
  if(t==='dark'||t==='light') d.setAttribute('data-theme',t);
  var s=localStorage.getItem('sumgo-size');
  if(s==='large'||s==='xlarge') d.setAttribute('data-size',s);
}}catch(e){{}}}})();
</script>
</head>
<body>

<div class="shell" id="gate">
  <header class="masthead">
    <a class="back" href="excel.html">← 엑셀 과목 첫 화면으로</a>
    <p class="eyebrow">1과목 · 엑셀 · 실무 편 · 1일차</p>
    <h1>암호를 넣어<br><em>열어 주세요</em></h1>
    <p class="lede">
      이 강의안에는 <b>회사에서 실제로 쓰는 자료</b>가 들어 있어
      암호를 걸어 두었습니다. 강의 때 알려 드린 <b>네 자리</b>를 넣으세요.
    </p>

    <form id="unlock" class="prefs" style="margin-left:0;margin-top:26px">
      <p class="prefs-title"><label for="pw">암호</label></p>
      <input id="pw" type="password" inputmode="numeric" autocomplete="off"
             aria-describedby="msg"
             style="font:inherit;font-size:1.1rem;padding:11px 14px;width:190px;
                    border:1px solid var(--rule-2);border-radius:9px;
                    background:var(--paper);color:var(--ink)">
      <button type="submit" class="theme-btn" style="margin-left:8px">열기</button>
      <p id="msg" class="prefs-title" style="margin-top:14px;min-height:1.4em"></p>
    </form>
  </header>

  <div class="guide">
    <h2>암호를 모르시면</h2>
    <p>
      이 강의안은 <b>사내 교육용</b>입니다. 강사에게 물어보세요.
      암호 없이 볼 수 있는 <b>엑셀 1·2·3부</b>는
      <a href="excel.html">엑셀 과목 첫 화면</a>에 있습니다.
    </p>
  </div>
</div>

<script>
(function () {{
  'use strict';
  var LOCK = {{
    salt: '{salt}',
    iv:   '{iv}',
    data: '{data}',
    iters: {iters}
  }};
  var KEY = 'sumgo-unlock:' + location.pathname.split('/').pop();

  var form = document.getElementById('unlock');
  var pw   = document.getElementById('pw');
  var msg  = document.getElementById('msg');

  function bytes(b64) {{
    var s = atob(b64), a = new Uint8Array(s.length);
    for (var i = 0; i < s.length; i++) a[i] = s.charCodeAt(i);
    return a;
  }}

  function open_(html) {{
    document.getElementById('gate').outerHTML = html;
    var s = document.createElement('script');
    s.src = 'script.js';
    document.body.appendChild(s);
    if (location.hash) {{
      var t = document.getElementById(location.hash.slice(1));
      if (t) t.scrollIntoView();
    }}
  }}

  function unlock(password) {{
    var c = window.crypto && window.crypto.subtle;
    if (!c) {{
      return Promise.reject(new Error(
        '이 브라우저에서는 열 수 없습니다. 파일을 그냥 열지 말고 ' +
        '주소(https://…)로 들어와 주세요.'));
    }}
    return c.importKey('raw', new TextEncoder().encode(password),
                       'PBKDF2', false, ['deriveKey'])
      .then(function (base) {{
        return c.deriveKey(
          {{ name: 'PBKDF2', salt: bytes(LOCK.salt),
             iterations: LOCK.iters, hash: 'SHA-256' }},
          base, {{ name: 'AES-GCM', length: 256 }}, false, ['decrypt']);
      }})
      .then(function (key) {{
        return c.decrypt({{ name: 'AES-GCM', iv: bytes(LOCK.iv) }},
                         key, bytes(LOCK.data));
      }})
      .then(function (buf) {{ return new TextDecoder().decode(buf); }});
  }}

  /* 두 번 눌러도 한 번만 돌게 잠가 둡니다. */
  var busy = false;

  function attempt() {{
    var v = pw.value;
    if (busy || !v) return;
    busy = true;
    msg.textContent = '여는 중입니다…';
    unlock(v).then(function (html) {{
      try {{ sessionStorage.setItem(KEY, v); }} catch (err) {{}}
      open_(html);
    }}, function (err) {{
      busy = false;
      msg.textContent = err && err.message && err.message.indexOf('브라우저') > -1
        ? err.message : '암호가 맞지 않습니다.';
      pw.value = '';
      pw.focus();
    }});
  }}

  form.addEventListener('submit', function (e) {{ e.preventDefault(); attempt(); }});

  /* 엔터로도 열립니다. 폼이 알아서 보내 주기도 하지만,
     busy 잠금이 있어 두 번 도는 일은 없습니다. */
  pw.addEventListener('keydown', function (e) {{
    if (e.key === 'Enter' || e.keyCode === 13) {{ e.preventDefault(); attempt(); }}
  }});

  /* 한 번 열었으면 그 창에서는 다시 묻지 않습니다(장을 옮겨 다닐 때). */
  var saved = null;
  try {{ saved = sessionStorage.getItem(KEY); }} catch (e) {{}}
  if (saved) {{
    unlock(saved).then(open_, function () {{
      try {{ sessionStorage.removeItem(KEY); }} catch (e) {{}}
    }});
  }}
  pw.focus();
}})();
</script>
</body>
</html>
"""


if __name__ == "__main__":
    if len(sys.argv) != 4:
        raise SystemExit("사용법: python tools/build-locked-page.py <원본> <결과> <암호>")
    build(sys.argv[1], sys.argv[2], sys.argv[3])
