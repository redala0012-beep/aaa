/* 해병대 전사(戰史) 디지털 연표 - 변환기
 * 엑셀(.xlsx) + 이미지 폴더 → 사이트(zip)
 * 외부 라이브러리 없이 동작 (인터넷 연결 불필요). Chrome / Edge 최신판 기준.
 */
(function () {
  'use strict';

  var WATERMARK = '해병대 군사연구소';
  var FULL_MAX = 1600, THUMB_MAX = 480, JPEG_Q = 0.8;

  /* ---------- ZIP 읽기 ---------- */
  function u16(v, o) { return v.getUint16(o, true); }
  function u32(v, o) { return v.getUint32(o, true); }

  async function inflateRaw(bytes) {
    var ds = new DecompressionStream('deflate-raw');
    var out = new Blob([bytes]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(out).arrayBuffer());
  }

  async function readZip(buf) {
    var v = new DataView(buf), bytes = new Uint8Array(buf), eocd = -1;
    for (var i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 66000); i--) {
      if (u32(v, i) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('엑셀 파일 형식이 아닙니다 (.xlsx 파일인지 확인하십시오).');
    var n = u16(v, eocd + 10), p = u32(v, eocd + 16), files = {}, dec = new TextDecoder();
    for (var k = 0; k < n; k++) {
      var method = u16(v, p + 10), csize = u32(v, p + 20), nlen = u16(v, p + 28),
          elen = u16(v, p + 30), clen = u16(v, p + 32), loff = u32(v, p + 42);
      var name = dec.decode(bytes.subarray(p + 46, p + 46 + nlen));
      var dataStart = loff + 30 + u16(v, loff + 26) + u16(v, loff + 28);
      files[name] = { method: method, data: bytes.subarray(dataStart, dataStart + csize) };
      p += 46 + nlen + elen + clen;
    }
    return {
      has: function (name) { return !!files[name]; },
      text: async function (name) {
        var f = files[name]; if (!f) return null;
        var raw = f.method === 0 ? f.data : await inflateRaw(f.data);
        return new TextDecoder('utf-8').decode(raw);
      }
    };
  }

  /* ---------- XLSX 읽기 ---------- */
  function parseXml(s) { return new DOMParser().parseFromString(s, 'application/xml'); }
  function byTag(node, tag) { return Array.prototype.slice.call(node.getElementsByTagNameNS('*', tag)); }
  function allText(node) { return byTag(node, 't').map(function (t) { return t.textContent; }).join(''); }
  function colIndex(ref) {
    var m = /^([A-Z]+)/.exec(ref), n = 0;
    for (var i = 0; i < m[1].length; i++) n = n * 26 + (m[1].charCodeAt(i) - 64);
    return n - 1;
  }

  async function readXlsx(file) {
    var zip = await readZip(await file.arrayBuffer());
    var wb = parseXml(await zip.text('xl/workbook.xml'));
    var rels = parseXml(await zip.text('xl/_rels/workbook.xml.rels'));
    var relMap = {};
    byTag(rels, 'Relationship').forEach(function (r) { relMap[r.getAttribute('Id')] = r.getAttribute('Target'); });
    var shared = [];
    if (zip.has('xl/sharedStrings.xml')) {
      shared = byTag(parseXml(await zip.text('xl/sharedStrings.xml')), 'si').map(allText);
    }
    var sheets = {};
    for (var s of byTag(wb, 'sheet')) {
      var rid = s.getAttribute('r:id') || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
      var target = relMap[rid].replace(/^\/?(xl\/)?/, '');
      var doc = parseXml(await zip.text('xl/' + target));
      var rows = [];
      byTag(doc, 'row').forEach(function (r) {
        var row = [];
        byTag(r, 'c').forEach(function (c) {
          var t = c.getAttribute('t'), vEl = byTag(c, 'v')[0], val;
          if (t === 's') val = shared[+vEl.textContent];
          else if (t === 'inlineStr') val = allText(c);
          else if (t === 'b') val = vEl && vEl.textContent === '1' ? 'TRUE' : 'FALSE';
          else if (vEl) val = t === 'str' || t === 'e' ? vEl.textContent : Number(vEl.textContent);
          else val = '';
          row[colIndex(c.getAttribute('r'))] = val;
        });
        rows.push(row);
      });
      sheets[s.getAttribute('name')] = rows;
    }
    return sheets;
  }

  function toObjects(rows) {
    if (!rows || !rows.length) return [];
    var head = rows[0].map(function (h) { return String(h == null ? '' : h).trim(); });
    var out = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i] || [], o = { _row: i + 1 }, any = false;
      head.forEach(function (h, k) {
        if (!h) return;
        var v = r[k]; if (v == null) v = '';
        if (typeof v === 'string') v = v.trim();
        if (v !== '') any = true;
        o[h] = v;
      });
      if (any) out.push(o);
    }
    return out;
  }

  /* ---------- 값 정리 ---------- */
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function normDate(v) {
    if (v === '' || v == null) return '';
    if (typeof v === 'number') {
      var d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
      return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
    }
    var m = String(v).match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
    if (m) return m[1] + '-' + pad(+m[2]) + '-' + pad(+m[3]);
    m = String(v).match(/^(\d{4})(\d{2})(\d{2})$/);
    if (m) return m[1] + '-' + m[2] + '-' + m[3];
    return null;
  }
  function yes(v) { return /^(o|ㅇ|○|y|yes|예|true|1)$/i.test(String(v).trim()); }
  function str(v) { return v == null ? '' : String(v).trim(); }
  function safeId(s) { return String(s).replace(/[^A-Za-z0-9_-]/g, '_'); }

  /* ---------- 데이터 검사·변환 ---------- */
  var BATTLE_REQ = ['전투ID', '전투명', '시작일', '전쟁구분', '장소', '참전부대', '한줄요약', '경과1', '의의교훈', '검수상태'];
  var PHOTO_REQ = ['자료ID', '파일명', '제목', '전투ID', '자료종류', '출처', '대표사진', '공개여부'];

  function buildData(sheets, imageMap) {
    var errors = [], warnings = [];
    var bRows = toObjects(sheets['전투']), pRows = toObjects(sheets['자료']);
    if (!sheets['전투']) errors.push('엑셀에 [전투] 시트가 없습니다.');
    if (!sheets['자료']) errors.push('엑셀에 [자료] 시트가 없습니다.');

    var battles = [], byId = {};
    bRows.forEach(function (r) {
      var where = '[전투] ' + r._row + '행';
      BATTLE_REQ.forEach(function (k) { if (str(r[k]) === '') errors.push(where + ': "' + k + '" 값이 비어 있습니다.'); });
      var id = str(r['전투ID']);
      if (id && byId[id]) errors.push(where + ': 전투ID "' + id + '"가 중복됩니다.');
      var start = normDate(r['시작일']), end = normDate(r['종료일']);
      if (start === null) errors.push(where + ': 시작일 "' + r['시작일'] + '" 형식이 올바르지 않습니다 (예: 1951-06-04).');
      if (end === null) errors.push(where + ': 종료일 "' + r['종료일'] + '" 형식이 올바르지 않습니다.');
      var status = str(r['검수상태']);
      if (status && status !== '초안' && status !== '검수완료') warnings.push(where + ': 검수상태는 "초안" 또는 "검수완료"로 적어 주십시오. "초안"으로 처리합니다.');
      var steps = [];
      for (var n = 1; n <= 5; n++) {
        var t = str(r['경과' + n]); if (!t) continue;
        var bar = t.indexOf('|');
        steps.push(bar >= 0 ? { title: t.slice(0, bar).trim(), body: t.slice(bar + 1).trim() } : { title: '경과 ' + n, body: t });
      }
      var b = {
        id: id, name: str(r['전투명']), start: start || '', end: end || '', war: str(r['전쟁구분']),
        place: str(r['장소']), units: str(r['참전부대']), oneLine: str(r['한줄요약']), summary: str(r['요약']),
        steps: steps, lesson: str(r['의의교훈']), status: status === '검수완료' ? '검수완료' : '초안',
        intro: { soldier: str(r['도입_장병']), officer: str(r['도입_간부']), public: str(r['도입_일반']) },
        lessons: { soldier: str(r['교훈_장병']), officer: str(r['교훈_간부']), public: str(r['교훈_일반']) },
        memorial: str(r['추모문구']), photos: [], quizzes: [], sources: [], decisions: []
      };
      if (id) byId[id] = b;
      battles.push(b);
    });

    var jobs = [], seenPhoto = {}, hidden = 0;
    pRows.forEach(function (r) {
      var where = '[자료] ' + r._row + '행';
      PHOTO_REQ.forEach(function (k) { if (str(r[k]) === '') errors.push(where + ': "' + k + '" 값이 비어 있습니다.'); });
      var pid = str(r['자료ID']);
      if (pid && seenPhoto[pid]) errors.push(where + ': 자료ID "' + pid + '"가 중복됩니다.');
      seenPhoto[pid] = 1;
      if (!yes(r['공개여부'])) { hidden++; return; }
      var bid = str(r['전투ID']), b = byId[bid];
      if (bid && !b) { errors.push(where + ': 전투ID "' + bid + '"가 [전투] 시트에 없습니다.'); return; }
      var fname = str(r['파일명']), file = imageMap[fname.toLowerCase()];
      if (fname && !file) { errors.push(where + ': 이미지 파일 "' + fname + '"을 이미지 폴더에서 찾을 수 없습니다.'); return; }
      if (!b || !file) return;
      var key = safeId(pid);
      var p = {
        id: pid, file: 'images/full/' + key + '.jpg', thumb: 'images/thumb/' + key + '.jpg',
        title: str(r['제목']), desc: str(r['설명']), type: str(r['자료종류']), year: str(r['연도']),
        source: str(r['출처']), holder: str(r['소장처']), cover: yes(r['대표사진']),
        edu: str(r['교육용']) === '' ? '' : (yes(r['교육용']) ? 'O' : 'X'),
        order: r['표시순서'] === '' ? 9999 : Number(r['표시순서'])
      };
      b.photos.push(p);
      jobs.push({ photo: p, file: file, key: key });
    });

    battles.forEach(function (b) {
      b.photos.sort(function (x, y) { return x.order - y.order; });
      var covers = b.photos.filter(function (p) { return p.cover; }).length;
      if (covers === 0) errors.push('전투 "' + b.name + '": 공개된 대표 사진이 없습니다. [자료] 시트에서 1장을 대표사진 O로 표시하십시오.');
      if (covers > 1) errors.push('전투 "' + b.name + '": 대표 사진이 ' + covers + '장입니다. 1장만 O로 표시하십시오.');
      b.photos.forEach(function (p) { delete p.order; });
    });
    if (!battles.length) errors.push('[전투] 시트에 전투가 한 건도 없습니다.');

    /* [퀴즈] 시트 (선택) */
    var AUD = { '전체': 'all', '장병': 'soldier', '간부': 'officer', '일반': 'public' };
    var seenQuiz = {};
    if (!sheets['퀴즈']) warnings.push('[퀴즈] 시트가 없어 퀴즈 없이 만듭니다.');
    toObjects(sheets['퀴즈']).forEach(function (r) {
      var where = '[퀴즈] ' + r._row + '행';
      var before = errors.length;
      ['퀴즈ID', '전투ID', '대상', '위치', '유형', '문제', '정답', '해설'].forEach(function (k) {
        if (str(r[k]) === '') errors.push(where + ': "' + k + '" 값이 비어 있습니다.');
      });
      var qid = str(r['퀴즈ID']);
      if (qid && seenQuiz[qid]) errors.push(where + ': 퀴즈ID "' + qid + '"가 중복됩니다.');
      seenQuiz[qid] = 1;
      var b = byId[str(r['전투ID'])];
      if (str(r['전투ID']) && !b) errors.push(where + ': 전투ID "' + str(r['전투ID']) + '"가 [전투] 시트에 없습니다.');
      var target = AUD[str(r['대상'])];
      if (str(r['대상']) && !target) errors.push(where + ': 대상은 전체 / 장병 / 간부 / 일반 중 하나로 적어 주십시오.');
      var pos = str(r['위치']), after;
      if (pos === '끝') after = 99;
      else if (/^\d+$/.test(pos)) {
        after = +pos;
        if (b && (after < 1 || after > b.steps.length)) errors.push(where + ': 위치 ' + pos + '에 해당하는 경과가 없습니다 (이 전투의 경과는 ' + b.steps.length + '개).');
      } else if (pos) errors.push(where + ': 위치는 경과 번호(숫자) 또는 "끝"으로 적어 주십시오.');
      var type = str(r['유형']).toUpperCase(), options, answer, ans = str(r['정답']).toUpperCase();
      if (type === 'OX') {
        options = ['O', 'X']; answer = ans === 'O' || ans === '○' ? 0 : ans === 'X' || ans === '×' ? 1 : -1;
        if (ans && answer < 0) errors.push(where + ': OX 문제의 정답은 O 또는 X로 적어 주십시오.');
      } else if (type === '객관식') {
        options = [1, 2, 3, 4].map(function (n) { return str(r['보기' + n]); });
        while (options.length && !options[options.length - 1]) options.pop();
        if (options.length < 2 || options.indexOf('') >= 0) errors.push(where + ': 객관식은 보기1부터 빈칸 없이 2개 이상 적어 주십시오.');
        answer = /^\d$/.test(ans) ? +ans - 1 : -1;
        if (ans && (answer < 0 || answer >= options.length)) errors.push(where + ': 정답은 보기 번호(1~' + options.length + ')로 적어 주십시오.');
      } else if (type) errors.push(where + ': 유형은 OX 또는 객관식으로 적어 주십시오.');
      if (errors.length > before || !b) return;
      b.quizzes.push({ id: qid, target: target, after: after, type: type === 'OX' ? 'ox' : 'mc',
        q: str(r['문제']), options: options, answer: answer, explain: str(r['해설']) });
    });

    /* [용어] 시트 (선택) */
    var terms = [];
    toObjects(sheets['용어']).forEach(function (r) {
      if (str(r['용어']) && str(r['풀이'])) terms.push({ term: str(r['용어']), desc: str(r['풀이']) });
      else warnings.push('[용어] ' + r._row + '행: 용어와 풀이를 모두 적어 주십시오. 이 행은 건너뜁니다.');
    });

    /* [근거] 시트 (선택): 개요·경과의 신뢰 수준과 근거 */
    var LEVEL = { '확인': 'confirmed', '추정': 'estimated', '이견': 'disputed', '기록마다 다름': 'disputed' };
    toObjects(sheets['근거']).forEach(function (r) {
      var where = '[근거] ' + r._row + '행', before = errors.length;
      ['전투ID', '위치', '신뢰수준', '근거'].forEach(function (k) { if (str(r[k]) === '') errors.push(where + ': "' + k + '" 값이 비어 있습니다.'); });
      var b = byId[str(r['전투ID'])];
      if (str(r['전투ID']) && !b) errors.push(where + ': 전투ID "' + str(r['전투ID']) + '"가 [전투] 시트에 없습니다.');
      var pos = str(r['위치']), at = null;
      if (pos === '개요') at = 0;
      else if (/^\d+$/.test(pos)) { at = +pos; if (b && (at < 1 || at > b.steps.length)) errors.push(where + ': 위치 ' + pos + '에 해당하는 경과가 없습니다 (이 전투의 경과는 ' + b.steps.length + '개).'); }
      else if (pos) errors.push(where + ': 위치는 "개요" 또는 경과 번호(숫자)로 적어 주십시오.');
      var level = LEVEL[str(r['신뢰수준'])];
      if (str(r['신뢰수준']) && !level) errors.push(where + ': 신뢰수준은 확인 / 추정 / 이견 중 하나로 적어 주십시오.');
      if (errors.length > before || !b) return;
      (b.sources = b.sources || []).push({ at: at, level: level, text: str(r['근거']) });
    });

    /* [결심] 시트 (선택): "지휘관이라면?" 시나리오 */
    var seenDec = {};
    toObjects(sheets['결심']).forEach(function (r) {
      var where = '[결심] ' + r._row + '행', before = errors.length;
      ['결심ID', '전투ID', '대상', '위치', '상황', '선택1', '선택2', '실제선택', '실제결과'].forEach(function (k) {
        if (str(r[k]) === '') errors.push(where + ': "' + k + '" 값이 비어 있습니다.');
      });
      var did = str(r['결심ID']);
      if (did && seenDec[did]) errors.push(where + ': 결심ID "' + did + '"가 중복됩니다.');
      seenDec[did] = 1;
      var b = byId[str(r['전투ID'])];
      if (str(r['전투ID']) && !b) errors.push(where + ': 전투ID "' + str(r['전투ID']) + '"가 [전투] 시트에 없습니다.');
      var target = AUD[str(r['대상'])];
      if (str(r['대상']) && !target) errors.push(where + ': 대상은 전체 / 장병 / 간부 / 일반 중 하나로 적어 주십시오.');
      var pos = str(r['위치']), after = null;
      if (/^\d+$/.test(pos)) { after = +pos; if (b && after > b.steps.length) errors.push(where + ': 위치 ' + pos + '에 해당하는 경과가 없습니다 (0 = 경과 시작 전, 이 전투의 경과는 ' + b.steps.length + '개).'); }
      else if (pos) errors.push(where + ': 위치는 숫자로 적어 주십시오 (0 = 경과 시작 전, n = 경과 n 뒤).');
      var options = [1, 2, 3].map(function (n) { return { text: str(r['선택' + n]), note: str(r['설명' + n]) }; });
      while (options.length && !options[options.length - 1].text) options.pop();
      if (options.some(function (o) { return !o.text; })) errors.push(where + ': 선택지는 선택1부터 빈칸 없이 적어 주십시오.');
      var act = str(r['실제선택']), actual = /^\d$/.test(act) ? +act - 1 : -1;
      if (act && (actual < 0 || actual >= options.length)) errors.push(where + ': 실제선택은 선택지 번호(1~' + options.length + ')로 적어 주십시오.');
      if (errors.length > before || !b) return;
      (b.decisions = b.decisions || []).push({ id: did, target: target, after: after, situation: str(r['상황']),
        options: options, actual: actual, result: str(r['실제결과']) });
    });

    /* [개정] 시트 (선택): 개정 이력 */
    var revisions = [];
    toObjects(sheets['개정']).forEach(function (r) {
      var where = '[개정] ' + r._row + '행', before = errors.length;
      if (str(r['날짜']) === '') errors.push(where + ': "날짜" 값이 비어 있습니다.');
      if (str(r['내용']) === '') errors.push(where + ': "내용" 값이 비어 있습니다.');
      var date = normDate(r['날짜']);
      if (str(r['날짜']) && !date) errors.push(where + ': 날짜 "' + r['날짜'] + '" 형식이 올바르지 않습니다 (예: 2026-11-17).');
      var bid = str(r['전투ID']);
      if (bid && !byId[bid]) errors.push(where + ': 전투ID "' + bid + '"가 [전투] 시트에 없습니다. 전체 개정이면 비워 두십시오.');
      if (errors.length > before) return;
      revisions.push({ date: date, battle: bid, text: str(r['내용']), basis: str(r['근거']) });
    });
    revisions.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });

    return { battles: battles, terms: terms, revisions: revisions, jobs: jobs, errors: errors, warnings: warnings, hidden: hidden };
  }

  /* ---------- 이미지 처리 ---------- */
  async function processImage(file, maxSide) {
    var bmp = await createImageBitmap(file);
    var s = Math.min(1, maxSide / Math.max(bmp.width, bmp.height));
    var w = Math.round(bmp.width * s), h = Math.round(bmp.height * s);
    var cv = document.createElement('canvas'); cv.width = w; cv.height = h;
    var ctx = cv.getContext('2d');
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, w, h);
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(bmp, 0, 0, w, h);
    bmp.close && bmp.close();
    var fs = Math.max(12, Math.round(Math.min(w, h * 1.5) * 0.032)), m = Math.round(fs * 0.7);
    ctx.font = 'bold ' + fs + 'px "Malgun Gothic","맑은 고딕",sans-serif';
    ctx.textAlign = 'right'; ctx.textBaseline = 'bottom';
    ctx.shadowColor = 'rgba(0,0,0,.75)'; ctx.shadowBlur = Math.max(2, fs / 5);
    ctx.shadowOffsetX = 1; ctx.shadowOffsetY = 1;
    ctx.fillStyle = 'rgba(255,255,255,.78)';
    ctx.fillText(WATERMARK, w - m, h - m);
    var blob = await new Promise(function (res) { cv.toBlob(res, 'image/jpeg', JPEG_Q); });
    return { bytes: new Uint8Array(await blob.arrayBuffer()), w: w, h: h };
  }

  /* ---------- ZIP 쓰기 (무압축) ---------- */
  var CRC = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  function crc32(b) { var c = 0xFFFFFFFF; for (var i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }

  function makeZip(entries) {
    var enc = new TextEncoder(), parts = [], central = [], offset = 0;
    var now = new Date();
    var dtime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    var ddate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    entries.forEach(function (e) {
      var name = enc.encode(e.name), data = e.data, crc = crc32(data);
      var lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
      lh.setUint16(8, 0, true); lh.setUint16(10, dtime, true); lh.setUint16(12, ddate, true);
      lh.setUint32(14, crc, true); lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true);
      lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
      parts.push(lh.buffer, name, data);
      var ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true);
      ch.setUint16(8, 0x0800, true); ch.setUint16(10, 0, true); ch.setUint16(12, dtime, true); ch.setUint16(14, ddate, true);
      ch.setUint32(16, crc, true); ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true);
      ch.setUint16(28, name.length, true); ch.setUint32(42, offset, true);
      central.push(ch.buffer, name);
      offset += 30 + name.length + data.length;
    });
    var csize = central.reduce(function (s, p) { return s + p.byteLength; }, 0);
    var end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
    end.setUint32(12, csize, true); end.setUint32(16, offset, true);
    return new Blob(parts.concat(central, [end.buffer]), { type: 'application/zip' });
  }

  /* ---------- 전체 실행 ---------- */
  function today() { var d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }

  async function run(xlsxFile, imageFiles, onProgress) {
    onProgress = onProgress || function () {};
    var imageMap = {};
    imageFiles.forEach(function (f) { if (/\.(jpe?g|png|gif|bmp|webp|tiff?)$/i.test(f.name)) imageMap[f.name.toLowerCase()] = f; });
    onProgress('엑셀을 읽는 중…', 0);
    var sheets = await readXlsx(xlsxFile);
    var d = buildData(sheets, imageMap);
    if (d.errors.length) return { ok: false, errors: d.errors, warnings: d.warnings };

    var entries = [], enc = new TextEncoder();
    for (var i = 0; i < d.jobs.length; i++) {
      var j = d.jobs[i];
      onProgress('사진 처리 중 (' + (i + 1) + ' / ' + d.jobs.length + '): ' + j.file.name, (i / d.jobs.length));
      try {
        var full = await processImage(j.file, FULL_MAX);
        var th = await processImage(j.file, THUMB_MAX);
      } catch (e) {
        return { ok: false, errors: ['이미지 "' + j.file.name + '"을 열 수 없습니다. JPG 또는 PNG 형식인지 확인하십시오.'], warnings: d.warnings };
      }
      entries.push({ name: j.photo.file, data: full.bytes }, { name: j.photo.thumb, data: th.bytes });
    }
    var data = { generatedAt: today(), battles: d.battles, terms: d.terms, revisions: d.revisions };
    entries.unshift(
      { name: 'index.html', data: enc.encode(window.SITE_TEMPLATE) },
      { name: 'data.js', data: enc.encode('window.SITE_DATA = ' + JSON.stringify(data, null, 1) + ';\n') }
    );
    onProgress('압축 파일을 만드는 중…', 1);
    return {
      ok: true, zip: makeZip(entries), warnings: d.warnings,
      stats: { battles: d.battles.length, photos: d.jobs.length, hidden: d.hidden,
        quizzes: d.battles.reduce(function (s, b) { return s + b.quizzes.length; }, 0), terms: d.terms.length,
        sources: d.battles.reduce(function (s, b) { return s + (b.sources || []).length; }, 0),
        decisions: d.battles.reduce(function (s, b) { return s + (b.decisions || []).length; }, 0),
        revisions: d.revisions.length }
    };
  }

  window.Converter = { run: run, readXlsx: readXlsx };
})();
