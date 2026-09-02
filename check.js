/* ══════════════════════════════════════════════════════════════
   check.js — 실력 진단

   문항은 HTML 에 있습니다. 여기서는 세 가지만 합니다.
     1. 답을 읽어서 과목별 점수를 냅니다
     2. 낮게 답한 항목이 가리키는 '먼저 볼 곳' 을 모읍니다
     3. 결과를 종이(인쇄·PDF)와 사진(PNG)으로 내보냅니다

   답은 localStorage 에 남습니다. 사생활 보호 모드에서 막히는 경우가
   있어서 저장·읽기를 전부 try 로 감쌌습니다. 막혀도 진단은 됩니다.
   ══════════════════════════════════════════════════════════════ */

(function () {
  'use strict';

  var KEY = 'sumgo-check';

  function save(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }
  function load(k)    { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function drop(k)    { try { localStorage.removeItem(k); } catch (e) {} }

  var form    = document.getElementById('quiz');
  var result  = document.getElementById('result');
  if (!form || !result) return;

  var btnSee   = document.getElementById('seeResult');
  var btnReset = document.getElementById('resetAll');
  var counter  = document.getElementById('counter');

  var items = [].slice.call(form.querySelectorAll('.q-item'));
  var TOTAL = items.length;

  /* 과목 이름 — data-subject 값과 짝 */
  var SUBJECTS = [
    { key: 'basic',  name: '컴퓨터 기본' },
    { key: 'excel',  name: '엑셀' },
    { key: 'ppt',    name: '파워포인트' },
    { key: 'hangul', name: '한글' },
    { key: 'word',   name: '워드' },
    { key: 'ai',     name: 'AI에게 시키기' }
  ];

  /* '먼저 볼 곳' 이름과 주소 */
  var PARTS = {
    'basic':    { label: '어느 과목이든 00장 — 손동작 세 가지', href: 'excel-1.html#start' },
    'excel-1':  { label: '엑셀 1부 — 처음 켜고, 표 하나 만들기', href: 'excel-1.html' },
    'excel-2':  { label: '엑셀 2부 — 계산을 시키기',             href: 'excel-2.html' },
    'excel-3':  { label: '엑셀 3부 — 보기 좋게, 찾기 쉽게',      href: 'excel-3.html' },
    'ppt-1':    { label: '파워포인트 1부 — 장을 만들고 글자 넣기', href: 'ppt-1.html' },
    'ppt-2':    { label: '파워포인트 2부 — 보여줄 것 넣기',        href: 'ppt-2.html' },
    'ppt-3':    { label: '파워포인트 3부 — 통일하고, 발표하기',    href: 'ppt-3.html' },
    'hangul-1': { label: '한글 1부 — 글을 쓰고 모양 잡기',        href: 'hangul-1.html' },
    'hangul-2': { label: '한글 2부 — 표와 그림 넣기',             href: 'hangul-2.html' },
    'hangul-3': { label: '한글 3부 — 문서답게 만들고 인쇄하기',   href: 'hangul-3.html' },
    'word-1':   { label: '워드 1부 — 문서를 쓰고 모양 잡기',      href: 'word-1.html' },
    'word-2':   { label: '워드 2부 — 스타일로 문서 짜기',         href: 'word-2.html' },
    'word-3':   { label: '워드 3부 — 긴 문서 다루고 인쇄하기',    href: 'word-3.html' },
    'ai-1':     { label: 'AI 1부 — 시작하기',                     href: 'ai-1.html' },
    'ai-2':     { label: 'AI 2부 — 파일과 폴더 맡기기',           href: 'ai-2.html' },
    'ai-3':     { label: 'AI 3부 — 엑셀 맡기기',                  href: 'ai-3.html' },
    'ai-4':     { label: 'AI 4부 — 한글·워드 맡기기',             href: 'ai-4.html' },
    'ai-5':     { label: 'AI 5부 — 파워포인트 맡기기',            href: 'ai-5.html' }
  };

  /* '먼저 볼 곳' 에 한 번에 보여줄 곳 수.
     못한다고 답한 곳을 전부 늘어놓으면 열 곳이 넘어갑니다.
     그러면 '먼저' 가 아니게 됩니다. 세 곳만 보여주고 나머지는 한 줄로 적습니다. */
  var TOP_N = 3;

  /* '엑셀 2부 — 계산을 시키기' → '엑셀 2부' */
  function shortName(label) {
    var i = label.indexOf(' — ');
    return i > 0 ? label.slice(0, i) : label;
  }

  /* 고를 때는 못한다고 답한 개수가 많은 곳부터,
     보여줄 때는 배우는 차례대로 (그래야 '위에서부터' 가 맞습니다) */
  function pickParts(r) {
    var order = Object.keys(PARTS);
    var hit   = order.filter(function (k) { return r.weak[k]; });

    var top = hit.slice().sort(function (a, b) {
      var d = r.weak[b].length - r.weak[a].length;
      return d !== 0 ? d : order.indexOf(a) - order.indexOf(b);
    }).slice(0, TOP_N);

    return {
      total: hit.length,
      shown: order.filter(function (k) { return top.indexOf(k) >= 0; }),
      rest:  hit.filter(function (k) { return top.indexOf(k) < 0; })
    };
  }

  function restLine(rest) {
    if (!rest.length) return '';
    var names = rest.slice(0, 3).map(function (k) { return shortName(PARTS[k].label); });
    return '나머지 ' + rest.length + '곳(' + names.join(', ') +
           (rest.length > 3 ? ' 등' : '') + ')은 이 세 곳을 마친 뒤에 보시면 됩니다.';
  }

  /* 0~4 를 말로 */
  function levelWord(pct) {
    if (pct >= 85) return '매우 잘함';
    if (pct >= 65) return '잘함';
    if (pct >= 40) return '보통';
    if (pct >= 15) return '못함';
    return '전혀 모름';
  }

  /* ───────── 답 기억하기 ───────── */
  function restore() {
    var raw = load(KEY);
    if (!raw) return;
    var saved;
    try { saved = JSON.parse(raw); } catch (e) { return; }
    if (!saved || typeof saved !== 'object') return;
    Object.keys(saved).forEach(function (name) {
      var el = form.querySelector('input[name="' + name + '"][value="' + saved[name] + '"]');
      if (el) el.checked = true;
    });
  }

  function collect() {
    var out = {};
    items.forEach(function (li) {
      var picked = li.querySelector('input:checked');
      if (picked) out[picked.name] = picked.value;
    });
    return out;
  }

  function answeredCount() {
    return items.filter(function (li) { return li.querySelector('input:checked'); }).length;
  }

  function paintCounter() {
    var n = answeredCount();
    counter.textContent = TOTAL + '개 중 ' + n + '개 답함';
    btnSee.disabled = (n === 0);
    if (n === TOTAL) counter.textContent = '다 답하셨습니다 (' + TOTAL + '개)';
  }

  form.addEventListener('change', function (e) {
    if (e.target && e.target.type === 'radio') {
      save(KEY, JSON.stringify(collect()));
      paintCounter();
    }
  });

  /* ───────── 점수 내기 ───────── */
  function score() {
    var bySubject = {};
    SUBJECTS.forEach(function (s) { bySubject[s.key] = { got: 0, max: 0, answered: 0, n: 0 }; });

    var weak = {};   /* part -> [문항 글]  */

    items.forEach(function (li) {
      var subj = li.getAttribute('data-subject');
      var part = li.getAttribute('data-part');
      var box  = bySubject[subj];
      if (!box) return;
      box.n += 1;

      var picked = li.querySelector('input:checked');
      if (!picked) return;

      var v = parseInt(picked.value, 10);
      box.got += v;
      box.max += 4;
      box.answered += 1;

      /* '못함'(1) 이하면 그 부를 권합니다 */
      if (v <= 1 && part) {
        var t = li.querySelector('.q-label') ? li.querySelector('.q-label').textContent.trim() : '';
        if (!weak[part]) weak[part] = [];
        weak[part].push(t);
      }
    });

    var rows = SUBJECTS.map(function (s) {
      var b = bySubject[s.key];
      var pct = b.max > 0 ? Math.round(b.got / b.max * 100) : null;
      return { key: s.key, name: s.name, pct: pct, answered: b.answered, n: b.n };
    });

    return { rows: rows, weak: weak };
  }

  /* ───────── 결과 그리기 ───────── */
  function stamp() {
    var d = new Date();
    var p = function (x) { return (x < 10 ? '0' : '') + x; };
    return d.getFullYear() + '년 ' + (d.getMonth() + 1) + '월 ' + d.getDate() + '일 ' +
           p(d.getHours()) + ':' + p(d.getMinutes());
  }

  var lastResult = null;   /* 사진으로 저장할 때 다시 씁니다 */

  function render() {
    var r = score();
    lastResult = r;

    document.getElementById('resWhen').textContent = stamp() + ' 기준';

    /* 막대 */
    var bars = document.getElementById('bars');
    bars.innerHTML = '';
    r.rows.forEach(function (row) {
      var div = document.createElement('div');
      div.className = 'bar-row';

      var nm = document.createElement('b');
      nm.textContent = row.name;

      var track = document.createElement('div');
      track.className = 'bar-track';
      var fill = document.createElement('div');
      fill.className = 'bar-fill';
      fill.style.width = (row.pct === null ? 0 : row.pct) + '%';
      track.appendChild(fill);

      var lv = document.createElement('span');
      lv.className = 'bar-lv';
      lv.textContent = row.answered === 0
        ? '안 함'
        : levelWord(row.pct) + ' · ' + row.pct + '%';

      div.appendChild(nm); div.appendChild(track); div.appendChild(lv);
      bars.appendChild(div);
    });

    /* 먼저 볼 곳 */
    var list = document.getElementById('resList');
    var note = document.getElementById('resNote');
    list.innerHTML = '';
    list.className = 'res-list';

    var p = pickParts(r);

    if (p.total === 0) {
      list.className = 'res-list r-ok';
      var li = document.createElement('li');
      var a = document.createElement('span');
      a.className = 'r-where';
      a.textContent = '따로 먼저 볼 곳이 없습니다';
      var b = document.createElement('span');
      b.className = 'r-why';
      b.textContent = '못한다고 답하신 항목이 없습니다. 순서대로 훑어보시면서 손에 익히시면 됩니다.';
      li.appendChild(a); li.appendChild(b);
      list.appendChild(li);
      note.textContent = '';
    } else {
      p.shown.forEach(function (k) {
        var li = document.createElement('li');
        var w = document.createElement('span');
        w.className = 'r-where';
        w.textContent = PARTS[k].label;
        var y = document.createElement('span');
        y.className = 'r-why';
        y.textContent = '못하겠다고 답하신 것 ' + r.weak[k].length + '가지가 여기 있습니다';
        li.appendChild(w); li.appendChild(y);

        /* 어떤 문항이었는지는 접어 둡니다. 궁금한 분만 펴 보시면 됩니다. */
        var more = document.createElement('details');
        more.className = 'r-more';
        var sum = document.createElement('summary');
        sum.textContent = '어떤 것인지 보기';
        more.appendChild(sum);
        var ul = document.createElement('ul');
        r.weak[k].forEach(function (t) {
          var q = document.createElement('li');
          q.textContent = t;
          ul.appendChild(q);
        });
        more.appendChild(ul);
        li.appendChild(more);

        list.appendChild(li);
      });
      note.textContent = ('위에서부터 차례대로 보시면 됩니다. ' +
        '한 부가 한 번 앉은 자리에서 끝낼 분량입니다. ' + restLine(p.rest)).trim();
    }

    result.hidden = false;
    result.scrollIntoView({ block: 'start' });
  }

  btnSee.addEventListener('click', render);

  btnReset.addEventListener('click', function () {
    if (!window.confirm('답한 것을 모두 지웁니다. 계속할까요?')) return;
    form.querySelectorAll('input:checked').forEach(function (el) { el.checked = false; });
    drop(KEY);
    result.hidden = true;
    paintCounter();
    window.scrollTo(0, 0);
  });

  /* ───────── 종이로 (인쇄 → PDF) ───────── */
  var btnPrint = document.getElementById('savePdf');
  if (btnPrint) {
    btnPrint.addEventListener('click', function () {
      document.body.classList.add('printing-result');
      window.print();
    });
    window.addEventListener('afterprint', function () {
      document.body.classList.remove('printing-result');
    });
  }

  /* ───────── 사진으로 (PNG) ─────────
     바깥 라이브러리 없이 canvas 에 직접 그립니다.
     파일을 그냥 열어도(인터넷 없이도) 동작해야 하기 때문입니다. */
  function drawPng() {
    if (!lastResult) return;
    var r = lastResult;

    var W = 1000, pad = 56, y = 0;
    var lines = [];   /* 먼저 높이를 계산하려고 그릴 것을 모읍니다 */

    /* 화면 색을 그대로 가져옵니다 (어두운 화면이면 어둡게 나옵니다) */
    var cs = getComputedStyle(document.documentElement);
    var C = {
      paper:  cs.getPropertyValue('--paper').trim()  || '#FBFAF7',
      panel:  cs.getPropertyValue('--panel').trim()  || '#FFFFFF',
      ink:    cs.getPropertyValue('--ink').trim()    || '#1A1A1A',
      ink2:   cs.getPropertyValue('--ink-2').trim()  || '#444',
      ink3:   cs.getPropertyValue('--ink-3').trim()  || '#777',
      rule:   cs.getPropertyValue('--rule').trim()   || '#DDD',
      accent: cs.getPropertyValue('--accent').trim() || '#1E6B45',
      dim:    cs.getPropertyValue('--accent-dim').trim() || '#EEF5F1',
      code:   cs.getPropertyValue('--code-bg').trim() || '#F1F1EF'
    };
    var FONT = '"Malgun Gothic","맑은 고딕",AppleSDGothicNeo-Regular,"Apple SD Gothic Neo",sans-serif';

    /* 글줄 나누기 — 한글은 어절 단위로 */
    function wrap(ctx, text, maxW) {
      var words = String(text).split(' ');
      var out = [], cur = '';
      words.forEach(function (w) {
        var t = cur ? cur + ' ' + w : w;
        if (ctx.measureText(t).width > maxW && cur) { out.push(cur); cur = w; }
        else cur = t;
      });
      if (cur) out.push(cur);
      return out;
    }

    /* 1차: 높이 재기 */
    var probe = document.createElement('canvas').getContext('2d');
    probe.font = '400 17px ' + FONT;
    var innerW = W - pad * 2;

    var H = pad;
    H += 46 + 12;                       /* 제목 */
    H += 26 + 30;                       /* 날짜 */
    H += 30;                            /* '과목별' 소제목 */
    H += r.rows.length * 46 + 24;       /* 막대 */
    H += 34;                            /* '먼저 볼 곳' 소제목 */

    var p = pickParts(r);
    var blocks = [];
    probe.font = '400 15px ' + FONT;
    if (p.total === 0) {
      blocks.push({ where: '따로 먼저 볼 곳이 없습니다',
                    why: ['못한다고 답하신 항목이 없습니다. 순서대로 훑어보시면 됩니다.'] });
    } else {
      p.shown.forEach(function (k) {
        blocks.push({
          where: PARTS[k].label,
          why: ['못하겠다고 답하신 것 ' + r.weak[k].length + '가지가 여기 있습니다']
        });
      });
    }
    blocks.forEach(function (b) { H += 20 + 26 + b.why.length * 23 + 12; });

    var rest = p.rest.length ? wrap(probe, restLine(p.rest), innerW) : [];
    if (rest.length) H += 8 + rest.length * 23;
    H += 40 + pad;                      /* 꼬리말 */

    /* 2차: 실제로 그리기 (선명하게 2배) */
    var S = 2;
    var cv = document.createElement('canvas');
    cv.width = W * S; cv.height = H * S;
    var g = cv.getContext('2d');
    g.scale(S, S);

    /* 오래된 브라우저에는 roundRect 가 없습니다. 없으면 그냥 네모로 그립니다. */
    if (typeof g.roundRect !== 'function') {
      g.roundRect = function (x, y, w, h) { this.rect(x, y, w, h); return this; };
    }

    g.fillStyle = C.paper; g.fillRect(0, 0, W, H);

    y = pad;
    g.fillStyle = C.accent; g.font = '700 15px ' + FONT;
    g.fillText('sumgo · 컴퓨터 기초 배움터', pad, y); y += 26;

    g.fillStyle = C.ink; g.font = '700 34px ' + FONT;
    g.fillText('컴퓨터 실력 진단 결과', pad, y + 22); y += 46 + 12;

    g.fillStyle = C.ink3; g.font = '400 15px ' + FONT;
    g.fillText(stamp() + ' 기준', pad, y); y += 26 + 30;

    g.fillStyle = C.ink; g.font = '700 19px ' + FONT;
    g.fillText('과목별로 지금 어디쯤인가', pad, y); y += 30;

    r.rows.forEach(function (row) {
      g.fillStyle = C.ink; g.font = '700 16px ' + FONT;
      g.fillText(row.name, pad, y + 16);

      var bx = pad + 170, bw = innerW - 170 - 150;
      g.fillStyle = C.code;
      g.beginPath(); g.roundRect(bx, y + 5, bw, 14, 7); g.fill();
      if (row.pct) {
        g.fillStyle = C.accent;
        g.beginPath(); g.roundRect(bx, y + 5, Math.max(6, bw * row.pct / 100), 14, 7); g.fill();
      }

      g.fillStyle = C.ink2; g.font = '400 15px ' + FONT;
      var txt = row.answered === 0 ? '안 함' : levelWord(row.pct) + ' · ' + row.pct + '%';
      g.textAlign = 'right';
      g.fillText(txt, W - pad, y + 16);
      g.textAlign = 'left';

      g.strokeStyle = C.rule; g.lineWidth = 1;
      g.beginPath(); g.moveTo(pad, y + 34.5); g.lineTo(W - pad, y + 34.5); g.stroke();
      y += 46;
    });
    y += 24;

    g.fillStyle = C.ink; g.font = '700 19px ' + FONT;
    g.fillText('먼저 볼 곳', pad, y); y += 34;

    blocks.forEach(function (b) {
      var h = 20 + 26 + b.why.length * 23 - 6;
      g.fillStyle = C.dim;
      g.beginPath(); g.roundRect(pad, y, innerW, h, 9); g.fill();
      g.fillStyle = C.accent;
      g.fillRect(pad, y, 3, h);

      g.fillStyle = C.ink; g.font = '700 16px ' + FONT;
      g.fillText(b.where, pad + 18, y + 26);
      g.fillStyle = C.ink2; g.font = '400 15px ' + FONT;
      b.why.forEach(function (ln, i) { g.fillText(ln, pad + 18, y + 26 + 24 + i * 23); });

      y += h + 12;
    });

    if (rest.length) {
      y += 8;
      g.fillStyle = C.ink2; g.font = '400 15px ' + FONT;
      rest.forEach(function (ln, i) { g.fillText(ln, pad, y + 15 + i * 23); });
      y += rest.length * 23;
    }

    y += 26;
    g.fillStyle = C.ink3; g.font = '400 14px ' + FONT;
    g.fillText('이 종이는 점수표가 아닙니다. 어디부터 보면 좋을지 알려주는 안내입니다.', pad, y);

    return cv;
  }

  var btnPng = document.getElementById('savePng');
  if (btnPng) {
    btnPng.addEventListener('click', function () {
      var cv;
      try { cv = drawPng(); } catch (e) { cv = null; }
      if (!cv) {
        window.alert('사진으로 저장이 안 됩니다. 아래 "종이로 저장하기"를 눌러 PDF로 받으세요.');
        return;
      }
      var d = new Date();
      var name = 'sumgo_진단결과_' + d.getFullYear() +
                 ('0' + (d.getMonth() + 1)).slice(-2) +
                 ('0' + d.getDate()).slice(-2) + '.png';
      try {
        cv.toBlob(function (blob) {
          if (!blob) { window.alert('사진으로 저장이 안 됩니다. "종이로 저장하기"를 써 주세요.'); return; }
          var url = URL.createObjectURL(blob);
          var a = document.createElement('a');
          a.href = url; a.download = name;
          document.body.appendChild(a); a.click(); a.remove();
          setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
        }, 'image/png');
      } catch (e) {
        window.alert('사진으로 저장이 안 됩니다. "종이로 저장하기"를 써 주세요.');
      }
    });
  }

  /* 시작 */
  restore();
  paintCounter();
})();
