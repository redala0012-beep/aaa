// 로컬 시험용 서버: 정적 파일 + /api/explain (Vercel 함수와 같은 코드)
// 사용: node tools/dev-server.js  →  http://localhost:3000/site/index.html
// GEMINI_API_KEY 는 프로젝트 루트의 .env.local 에서 읽는다 (git에 올라가지 않음). 키 값은 화면에 출력하지 않는다.
const http = require('http');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
process.chdir(root);
const envFile = path.join(root, '.env.local');
if (fs.existsSync(envFile)) {
  for (const line of fs.readFileSync(envFile, 'utf8').replace(/^﻿/, '').split(/\r?\n/)) {
    const m = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*)\s*$/);
    if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^["']|["']$/g, '');
  }
}
console.log('GEMINI_API_KEY:', process.env.GEMINI_API_KEY ? '설정됨' : '없음 (.env.local 확인)');

const explain = require(path.join(root, 'api', 'explain.js'));
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css', '.json': 'application/json',
  '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png', '.svg': 'image/svg+xml', '.xlsx': 'application/octet-stream' };
const port = +process.env.PORT || 3000;

http.createServer((req, res) => {
  const url = decodeURIComponent(req.url.split('?')[0]);
  if (url === '/api/explain') return explain(req, res);
  if (url === '/') { res.writeHead(302, { Location: '/site/index.html' }); return res.end(); }
  const file = path.join(root, url);
  if (!file.startsWith(root) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) { res.statusCode = 404; return res.end('Not found'); }
  res.setHeader('Content-Type', TYPES[path.extname(file).toLowerCase()] || 'application/octet-stream');
  fs.createReadStream(file).pipe(res);
}).listen(port, () => console.log(`http://localhost:${port}/site/index.html`));
