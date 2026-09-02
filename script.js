/* ══════════════════════════════════════════════════════════════
   sumgo — 공통 스크립트

   다섯 가지 일만 합니다.
     1) 글자 크기 3단계 (보통 / 크게 / 아주 크게)
     2) 밝은 화면 ↔ 어두운 화면
     3) 왼쪽 목차에서 지금 읽는 장 표시
     4) 확인 목록 체크 상태 기억
     5) "이 말 복사하기" 단추 (AI 과목)

   저장은 전부 localStorage 입니다. 브라우저를 닫아도 남습니다.
   사생활 보호 모드 등에서 localStorage 가 아예 막히는 경우가 있어서
   읽기·쓰기를 모두 try 로 감쌌습니다. 막혀도 페이지는 그대로 동작합니다.

   ※ 화면이 번쩍하지 않도록, 테마·글자크기를 실제로 "적용"하는 코드는
     각 페이지 <head> 안에 짧게 인라인으로 들어 있습니다.
     여기서는 버튼을 눌렀을 때의 동작만 담당합니다.
   ══════════════════════════════════════════════════════════════ */

(function () {
  'use strict';

  var root = document.documentElement;
  var KEY_THEME = 'sumgo-theme';
  var KEY_SIZE  = 'sumgo-size';
  var KEY_CHECK = 'sumgo-check:';

  function save(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }
  function load(k)    { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function drop(k)    { try { localStorage.removeItem(k); } catch (e) {} }

  /* ───────── 1. 글자 크기 ───────── */
  var sizeBtns = document.querySelectorAll('[data-size-set]');

  function paintSize() {
    var cur = root.getAttribute('data-size') || 'normal';
    for (var i = 0; i < sizeBtns.length; i++) {
      var v = sizeBtns[i].getAttribute('data-size-set');
      sizeBtns[i].setAttribute('aria-pressed', v === cur ? 'true' : 'false');
    }
  }

  for (var i = 0; i < sizeBtns.length; i++) {
    sizeBtns[i].addEventListener('click', function () {
      var v = this.getAttribute('data-size-set');
      if (v === 'normal') { root.removeAttribute('data-size'); drop(KEY_SIZE); }
      else { root.setAttribute('data-size', v); save(KEY_SIZE, v); }
      paintSize();
    });
  }
  paintSize();

  /* ───────── 2. 밝은 화면 / 어두운 화면 ───────── */
  var themeBtn = document.getElementById('themeBtn');

  function isDark() {
    var set = root.getAttribute('data-theme');
    if (set === 'dark')  return true;
    if (set === 'light') return false;
    return window.matchMedia &&
           window.matchMedia('(prefers-color-scheme: dark)').matches;
  }

  function paintTheme() {
    if (!themeBtn) return;
    themeBtn.textContent = isDark() ? '☀ 밝은 화면으로' : '☾ 어두운 화면으로';
  }

  if (themeBtn) {
    themeBtn.addEventListener('click', function () {
      var next = isDark() ? 'light' : 'dark';
      root.setAttribute('data-theme', next);
      save(KEY_THEME, next);
      paintTheme();
    });
    paintTheme();
  }

  /* ───────── 3. 목차에서 지금 읽는 장 표시 ───────── */
  var links = document.querySelectorAll('.toc a[href^="#"]');
  if (links.length && 'IntersectionObserver' in window) {
    var map = {};
    var watched = [];

    for (var j = 0; j < links.length; j++) {
      var id = links[j].getAttribute('href').slice(1);
      var sec = document.getElementById(id);
      if (sec) { map[id] = links[j]; watched.push(sec); }
    }

    var seen = {};
    var io = new IntersectionObserver(function (entries) {
      for (var k = 0; k < entries.length; k++) {
        seen[entries[k].target.id] = entries[k].isIntersecting;
      }
      /* 화면에 걸친 것 중 문서 순서상 가장 위의 장을 현재로 본다 */
      var cur = null;
      for (var m = 0; m < watched.length; m++) {
        if (seen[watched[m].id]) { cur = watched[m].id; break; }
      }
      for (var id2 in map) {
        if (id2 === cur) map[id2].setAttribute('aria-current', 'true');
        else map[id2].removeAttribute('aria-current');
      }
    }, { rootMargin: '-15% 0px -70% 0px', threshold: 0 });

    for (var n = 0; n < watched.length; n++) io.observe(watched[n]);
  }

  /* ───────── 4. 확인 목록 기억 ─────────
     페이지마다 따로 저장합니다. 주소가 바뀌면 다른 목록으로 봅니다. */
  var page = location.pathname.split('/').pop() || 'index.html';
  var boxes = document.querySelectorAll('.checklist input[type="checkbox"]');

  for (var p = 0; p < boxes.length; p++) {
    (function (box, idx) {
      var key = KEY_CHECK + page + ':' + idx;
      if (load(key) === '1') box.checked = true;
      box.addEventListener('change', function () {
        if (box.checked) save(key, '1'); else drop(key);
      });
    })(boxes[p], p);
  }

  /* ───────── 5. "이 말 복사하기" ─────────
     .say 칸의 문단만 이어 붙여 복사합니다. 라벨("이렇게 말하세요")과
     아래쪽 설명(.after)은 붙여넣을 말이 아니므로 뺍니다.

     복사는 두 가지 방법을 씁니다. 요즘 방법(navigator.clipboard)은
     https 나 localhost 에서만 동작해서, 파일을 그냥 두 번 눌러 연
     경우(file://)에는 막힙니다. 그때는 예전 방법으로 물러섭니다. */
  var sayCards = document.querySelectorAll('.say');

  function sayText(card) {
    var ps = card.querySelectorAll('p');
    var out = [];
    for (var q = 0; q < ps.length; q++) {
      var t = ps[q].innerText.replace(/\s+/g, ' ').replace(/^\s+|\s+$/g, '');
      if (t) out.push(t);
    }
    return out.join('\n');
  }

  function oldWayCopy(text) {
    var ta = document.createElement('textarea');
    ta.value = text;
    ta.style.cssText = 'position:fixed;top:-1000px;opacity:0';
    document.body.appendChild(ta);
    ta.select();
    var ok = false;
    try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    document.body.removeChild(ta);
    return ok;
  }

  for (var r = 0; r < sayCards.length; r++) {
    (function (card) {
      var btn = card.querySelector('.say-copy');
      if (!btn) return;

      function tell(ok) {
        btn.textContent = ok ? '복사했습니다 ✓' : '복사가 안 됩니다';
        if (ok) btn.classList.add('done');
        setTimeout(function () {
          btn.textContent = '이 말 복사하기';
          btn.classList.remove('done');
        }, 2000);
      }

      btn.addEventListener('click', function () {
        var text = sayText(card);
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(text).then(
            function () { tell(true); },
            function () { tell(oldWayCopy(text)); }
          );
        } else {
          tell(oldWayCopy(text));
        }
      });
    })(sayCards[r]);
  }
})();
