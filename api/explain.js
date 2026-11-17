// 전투 설명을 대상자(장병·간부·방문객) 눈높이에 맞춰 3줄로 풀어 주는 서버 함수 (Vercel Serverless, Node 18+)
// - Gemini API 키는 서버 환경변수 GEMINI_API_KEY 에서만 읽는다. 브라우저로는 절대 보내지 않는다.
// - 브라우저는 전투 ID와 대상자만 보낸다. 본문은 서버가 site/data.js 에서 직접 읽으므로
//   임의의 프롬프트를 넣어 AI를 다른 용도로 쓰는 것을 막는다.
const fs = require('fs');
const path = require('path');

// 기본 모델: 응답이 1~2초로 빠른 경량 모델 (2026-11 시험: gemini-flash-latest·3.8-flash는 30초 넘게 응답 없음)
const MODEL = process.env.GEMINI_MODEL || 'gemini-3.5-flash-lite';
const AUD = {
  soldier: '해병대 장병(20대 초반 병사). 선배 해병의 정신과 자부심이 느껴지게, 군대식 존댓말(~습니다)로',
  officer: '해병대 간부·교육생. 작전 배경, 지휘 결심, 전술적 교훈이 드러나게, 간결한 보고체로',
  public: '처음 방문한 일반 시민(청소년 포함). 군사 용어는 쉬운 말로 바꾸고, 친근한 존댓말(~요)로'
};

let DATA = null;
function battles() {
  if (!DATA) {
    const raw = fs.readFileSync(path.join(process.cwd(), 'site', 'data.js'), 'utf8');
    DATA = JSON.parse(raw.slice(raw.indexOf('{'), raw.lastIndexOf('}') + 1));
  }
  return DATA.battles || [];
}

const cache = new Map(); // 같은 전투·대상자 요청은 다시 호출하지 않음 (인스턴스가 살아 있는 동안)

function buildPrompt(b, aud) {
  const steps = (b.steps || []).map((s, i) => `${i + 1}. ${s.title}: ${s.body}`).join('\n');
  return [
    `너는 해병대 군사연구소의 전사(戰史) 해설가다. 아래 [자료]만 근거로, ${AUD[aud]} 이 전투를 설명하라.`,
    '규칙:',
    '- 정확히 3줄. 각 줄은 한 문장, 80자 이내. 번호·기호·머리말 없이 줄바꿈으로만 구분.',
    '- [자료]에 없는 사실, 숫자, 인명, 날짜를 지어내지 말 것. 확실하지 않으면 쓰지 말 것.',
    b.memorial ? '- 희생자가 있는 전투다. 추모의 마음을 담아 차분하고 정중하게. 감탄사·느낌표 금지.' : '- 과장된 표현과 느낌표는 쓰지 말 것.',
    '',
    '[자료]',
    `전투: ${b.name} (${b.war}, ${b.start}${b.end && b.end !== b.start ? ' ~ ' + b.end : ''})`,
    `장소: ${b.place} / 참전 부대: ${b.units}`,
    `한 줄 요약: ${b.oneLine}`,
    `개요: ${b.summary || ''}`,
    `경과:\n${steps}`,
    `의의: ${(b.lessons && b.lessons[aud]) || b.lesson || ''}`
  ].join('\n');
}

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  if (req.method !== 'POST') { res.statusCode = 405; return res.end(JSON.stringify({ error: 'POST만 허용됩니다.' })); }
  const key = process.env.GEMINI_API_KEY;
  if (!key) { res.statusCode = 503; return res.end(JSON.stringify({ error: 'AI 설정이 아직 되지 않았습니다.' })); }

  let body = req.body;
  if (!body || typeof body === 'string') {
    try {
      body = typeof body === 'string' ? JSON.parse(body) : await new Promise((ok, no) => {
        let s = ''; req.on('data', c => { s += c; if (s.length > 2000) no(new Error('too large')); });
        req.on('end', () => { try { ok(JSON.parse(s || '{}')); } catch (e) { no(e); } }); req.on('error', no);
      });
    } catch (e) { res.statusCode = 400; return res.end(JSON.stringify({ error: '요청 형식이 잘못되었습니다.' })); }
  }
  const id = String(body.id || ''), aud = String(body.aud || '');
  const b = battles().find(x => x.id === id);
  if (!b || !AUD[aud]) { res.statusCode = 400; return res.end(JSON.stringify({ error: '전투 또는 대상자를 찾을 수 없습니다.' })); }

  const ck = id + '|' + aud;
  if (cache.has(ck)) return res.end(JSON.stringify({ lines: cache.get(ck), cached: true }));

  try {
    const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(MODEL)}:generateContent`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-goog-api-key': key },
      body: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: buildPrompt(b, aud) }] }] }),
      signal: AbortSignal.timeout(25000)
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) {
      console.error('Gemini error', r.status, j && j.error && j.error.message); // 키는 기록하지 않음
      res.statusCode = r.status === 429 ? 429 : 502;
      return res.end(JSON.stringify({ error: r.status === 429 ? '요청이 많습니다. 잠시 후 다시 눌러 주세요.' : 'AI 응답을 받지 못했습니다.' }));
    }
    const text = ((j.candidates || [])[0]?.content?.parts || []).map(p => p.text || '').join('');
    const lines = text.split(/\n+/).map(s => s.replace(/^\s*(?:[-*•·]|\d+[.)])\s*/, '').trim()).filter(Boolean).slice(0, 3);
    if (!lines.length) { res.statusCode = 502; return res.end(JSON.stringify({ error: 'AI가 빈 응답을 보냈습니다.' })); }
    cache.set(ck, lines);
    return res.end(JSON.stringify({ lines }));
  } catch (e) {
    console.error('Gemini call failed', e && e.name);
    res.statusCode = 504;
    return res.end(JSON.stringify({ error: 'AI 응답이 늦어 중단했습니다. 다시 눌러 주세요.' }));
  }
};
