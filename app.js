// app.js — تطبيق تقرير طلاب QUREO (يعمل أونلاين على السيرفر للموبايل ومحليًا على الكمبيوتر — بدون كلمة مرور)
const http = require('http');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');
const XLSX = require('xlsx');

process.on('uncaughtException', (err) => {
  console.error('Unexpected error:', err);
});

const PORT = process.env.PORT || 3456;
const PORTAL = 'https://me-portal.qureo.education';
const APP = 'https://me-tp.qureo.education';
const CHAPTERS_PATH = '/api/study/students/sections/5/chapters';
const TOTAL_CHAPTERS = 42;
const CONCURRENCY = 6;

const state = {
  running: false,
  stopRequested: false,
  total: 0,
  done: 0,
  okCount: 0,
  failCount: 0,
  rows: [],
  xlsxBuffer: null,
  error: '',
  elapsedSec: 0,
};

function normalizeStudentRow(rawObj) {
  let name = '';
  let className = '';
  let studentId = '';
  let password = '';
  let inferredClassFromHeader = '';

  for (const [rawKey, rawVal] of Object.entries(rawObj)) {
    const key = String(rawKey).replace(/[\r\n]+/g, ' ').trim();
    const compactKey = key.replace(/["\s;،:]/g, '').toLowerCase();
    const val = String(rawVal ?? '').trim();

    if (!name && (compactKey === 'name' || compactKey === 'الاسم' || key.includes('اسم الطالب'))) {
      name = val;
      const match = key.match(/(?:المدرسة\s*:\s*)?(.+?)\s*اسم\s*الطالب/);
      if (match && match[1]) inferredClassFromHeader = match[1].trim();
    } else if (!className && (compactKey === 'class' || compactKey === 'الفصل' || compactKey === 'فصل')) {
      className = val;
    } else if (
      !studentId &&
      (compactKey === 'student_id' || compactKey === 'studentid' || key.includes('المستخدم') || key.includes('كود الطالب'))
    ) {
      studentId = val;
    } else if (!password && (compactKey === 'password' || key.includes('كلمة المرور') || key.includes('كلمة السر'))) {
      password = val;
    }
  }

  return {
    name,
    class: className || inferredClassFromHeader,
    student_id: studentId,
    password,
  };
}

function parseStudentsFile(base64Data, fileName = '') {
  const buf = Buffer.from(base64Data, 'base64');
  const isExcel = /\.(xlsx|xls)$/i.test(fileName);

  if (isExcel) {
    const wb = XLSX.read(buf, { type: 'buffer' });
    const firstSheet = wb.Sheets[wb.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });

    return rawRows
      .map(normalizeStudentRow)
      .filter((r) => r.student_id || r.password || r.name);
  }

  const text = buf.toString('utf8').replace(/^\uFEFF/, '');
  const lines = text.split(/\r?\n/).filter((l) => l.trim() !== '');
  if (!lines.length) return [];

  const splitSmart = (line) => {
    const semiCount = (line.match(/;/g) || []).length;
    const commaCount = (line.match(/,/g) || []).length;
    const sep = semiCount >= commaCount ? ';' : ',';

    const out = [];
    let cell = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      const next = line[i + 1];
      if (inQuotes) {
        if (ch === '"' && next === '"') {
          cell += '"';
          i++;
        } else if (ch === '"') {
          inQuotes = false;
        } else {
          cell += ch;
        }
      } else {
        if (ch === '"') {
          inQuotes = true;
        } else if (ch === sep) {
          out.push(cell);
          cell = '';
        } else {
          cell += ch;
        }
      }
    }
    out.push(cell);
    return out.map((c) => c.trim()).filter((c, idx, arr) => !(c === '' && idx === arr.length - 1 && arr.length > 4));
  };

  const rawHeaders = lines[0]
    .split(/[;,]/)
    .map((h) => h.replace(/["\s;،]/g, ''))
    .filter(Boolean);

  return lines
    .slice(1)
    .map((line) => {
      const trimmedLine = line.replace(/,+$/g, '');
      const cells = splitSmart(trimmedLine);
      const obj = {};
      rawHeaders.forEach((h, idx) => {
        obj[h] = (cells[idx] ?? '').trim();
      });
      return normalizeStudentRow(obj);
    })
    .filter((r) => r.student_id || r.password || r.name);
}

function formatEgyptTime(isoStr) {
  if (!isoStr) return '';
  const d = new Date(isoStr);
  if (Number.isNaN(d.getTime())) return isoStr;
  return new Intl.DateTimeFormat('sv-SE', {
    timeZone: 'Africa/Cairo',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  }).format(d);
}

class CookieJar {
  constructor() {
    this.cookies = new Map();
  }
  saveFrom(res) {
    const list =
      typeof res.headers.getSetCookie === 'function'
        ? res.headers.getSetCookie()
        : res.headers.get('set-cookie')
          ? [res.headers.get('set-cookie')]
          : [];
    for (const raw of list) {
      const pair = raw.split(';')[0];
      const eq = pair.indexOf('=');
      if (eq > 0) {
        this.cookies.set(pair.slice(0, eq).trim(), pair.slice(eq + 1).trim());
      }
    }
  }
  header() {
    return Array.from(this.cookies.entries())
      .map(([k, v]) => `${k}=${v}`)
      .join('; ');
  }
}

async function fetchWithCookies(jar, url, opts = {}, maxRedirects = 8) {
  let currentUrl = url;
  let method = opts.method || 'GET';
  let body = opts.body;

  for (let i = 0; i <= maxRedirects; i++) {
    const headers = {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/130.0.0.0 Safari/537.36',
      ...(opts.headers || {}),
    };
    const cookieStr = jar.header();
    if (cookieStr) headers['Cookie'] = cookieStr;

    const res = await fetch(currentUrl, {
      method,
      headers,
      body,
      redirect: 'manual',
    });
    jar.saveFrom(res);

    if ([301, 302, 303, 307, 308].includes(res.status)) {
      const loc = res.headers.get('location');
      if (!loc) return res;
      currentUrl = new URL(loc, currentUrl).toString();
      if (res.status === 303 || ((res.status === 301 || res.status === 302) && method === 'POST')) {
        method = 'GET';
        body = undefined;
      }
      continue;
    }
    return res;
  }
  throw new Error('too many redirects');
}

async function checkStudent(s) {
  const jar = new CookieJar();
  try {
    const studentId = String(s.student_id || '').trim();
    const password = String(s.password || '').trim();

    const loginRes = await fetchWithCookies(jar, `${PORTAL}/oauth2/password-login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Accept: 'application/json, text/plain, */*',
        Referer: `${PORTAL}/login?referer=${encodeURIComponent(PORTAL + '/')}`,
        Origin: PORTAL,
      },
      body: JSON.stringify({ student_id: studentId, password }),
    });

    if (!loginRes.ok) throw new Error(`login HTTP ${loginRes.status}`);

    await fetchWithCookies(jar, `${PORTAL}/`).catch(() => {});
    await fetchWithCookies(jar, `${APP}/`).catch(() => {});

    const chRes = await fetchWithCookies(jar, `${APP}${CHAPTERS_PATH}`, {
      headers: {
        Accept: 'application/json, text/plain, */*',
        Referer: `${APP}/`,
      },
    });

    if (!chRes.ok) throw new Error(`chapters HTTP ${chRes.status}`);
    const text = await chRes.text();
    let chapters;
    try {
      chapters = JSON.parse(text);
    } catch {
      throw new Error('الاستجابة ليست JSON');
    }
    if (!Array.isArray(chapters)) throw new Error('unexpected response');

    const dates = chapters.flatMap((c) => [c.started_at, c.completed_at]).filter(Boolean).sort();
    return {
      status: 'ok',
      chaptersCount: chapters.length,
      diamonds: chapters.reduce((sum, c) => sum + (c.diamond || 0), 0),
      lastActivity: formatEgyptTime(dates.at(-1) || ''),
    };
  } catch (e) {
    return { status: `failed: ${e.message}`, chaptersCount: '', diamonds: '', lastActivity: '' };
  }
}

function buildExcelWorkbook(rows) {
  const wb = XLSX.utils.book_new();
  wb.Workbook = { Views: [{ RTL: true }] };

  const studentHeader = ['الاسم', 'الفصل', 'رقم الطالب', 'الحالة', 'الفصول', `من ${TOTAL_CHAPTERS}`, 'النسبة %', 'الألماس', 'آخر نشاط'];
  const studentData = [studentHeader];

  let c0 = 0;
  let c1_10 = 0;
  let c11_17 = 0;
  let c18_41 = 0;
  let c42 = 0;
  let failedTotal = 0;

  for (const r of rows) {
    if (!r) continue;
    studentData.push([
      r.name,
      r.className,
      r.studentId,
      r.status,
      r.chaptersCount === '' ? '' : Number(r.chaptersCount),
      r.totalChapters,
      r.pct === '' ? '' : Number(r.pct),
      r.diamonds === '' ? '' : Number(r.diamonds),
      r.lastActivity,
    ]);

    if (r.status !== 'ok' || r.chaptersCount === '') {
      failedTotal++;
    } else {
      const n = Number(r.chaptersCount);
      if (n === 0) c0++;
      else if (n >= 1 && n <= 10) c1_10++;
      else if (n >= 11 && n <= 17) c11_17++;
      else if (n >= 18 && n <= 41) c18_41++;
      else if (n >= 42) c42++;
    }
  }

  const wsStudents = XLSX.utils.aoa_to_sheet(studentData);
  wsStudents['!views'] = [{ rightToLeft: true }];
  wsStudents['!cols'] = [
    { wch: 32 },
    { wch: 24 },
    { wch: 15 },
    { wch: 18 },
    { wch: 10 },
    { wch: 10 },
    { wch: 11 },
    { wch: 10 },
    { wch: 18 },
  ];
  XLSX.utils.book_append_sheet(wb, wsStudents, 'بيانات الطلاب');

  const checkedTotal = c0 + c1_10 + c11_17 + c18_41 + c42;
  const pctOf = (count) => (checkedTotal > 0 ? Math.round((count / checkedTotal) * 1000) / 10 : 0);

  const summaryData = [
    ['الفئة / مستوى الإنجاز', 'عدد الطلاب', 'النسبة من المفحوصين %'],
    ['لم ينجزوا أي فصل (0 فصل)', c0, pctOf(c0)],
    ['أنجزوا من 1 إلى 10 فصول', c1_10, pctOf(c1_10)],
    ['أنجزوا من 11 إلى 17 فصلًا', c11_17, pctOf(c11_17)],
    ['أنجزوا من 18 إلى 41 فصلًا', c18_41, pctOf(c18_41)],
    ['أتموا المنهج كاملًا (42 فصلًا)', c42, pctOf(c42)],
    ['إجمالي الطلاب المفحوصين بنجاح', checkedTotal, checkedTotal > 0 ? 100 : 0],
  ];

  if (failedTotal > 0) {
    summaryData.push(['تعذر فحصهم (خطأ في الدخول)', failedTotal, '']);
    summaryData.push(['الإجمالي الكلي بالملف', checkedTotal + failedTotal, '']);
  }

  const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
  wsSummary['!views'] = [{ rightToLeft: true }];
  wsSummary['!cols'] = [
    { wch: 36 },
    { wch: 16 },
    { wch: 24 },
  ];
  XLSX.utils.book_append_sheet(wb, wsSummary, 'ملخص الإحصائيات');

  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

async function runJob(fileBase64, fileName) {
  const startTime = Date.now();
  const students = parseStudentsFile(fileBase64, fileName);

  Object.assign(state, {
    running: true,
    stopRequested: false,
    total: students.length,
    done: 0,
    okCount: 0,
    failCount: 0,
    rows: new Array(students.length),
    xlsxBuffer: null,
    error: '',
    elapsedSec: 0,
  });

  try {
    let nextIndex = 0;

    async function worker() {
      while (nextIndex < students.length && !state.stopRequested) {
        const i = nextIndex++;
        const s = students[i];
        const r = await checkStudent(s);
        const pct = r.chaptersCount === '' ? '' : Math.round((r.chaptersCount / TOTAL_CHAPTERS) * 100);

        state.rows[i] = {
          name: String(s.name || '').trim(),
          className: String(s.class || '').trim(),
          studentId: String(s.student_id || '').trim(),
          status: r.status,
          chaptersCount: r.chaptersCount,
          totalChapters: TOTAL_CHAPTERS,
          pct,
          diamonds: r.diamonds,
          lastActivity: r.lastActivity,
        };

        state.done++;
        if (r.status === 'ok') state.okCount++;
        else state.failCount++;
        state.elapsedSec = ((Date.now() - startTime) / 1000).toFixed(1);
      }
    }

    const workers = Array.from({ length: Math.min(CONCURRENCY, students.length || 1) }, () => worker());
    await Promise.all(workers);

    state.xlsxBuffer = buildExcelWorkbook(state.rows);
    if (!process.env.PORT) {
      const outDir = process.pkg ? path.dirname(process.execPath) : process.cwd();
      fs.writeFileSync(path.join(outDir, 'report.xlsx'), state.xlsxBuffer);
    }
  } catch (e) {
    state.error = e.message;
  } finally {
    state.running = false;
    state.elapsedSec = ((Date.now() - startTime) / 1000).toFixed(1);
  }
}

const HTML = `<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>تقرير تقدم طلاب QUREO</title>
<style>
  * { box-sizing: border-box; font-family: 'Segoe UI', Tahoma, sans-serif; }
  body { margin: 0; background: #f1f5f9; color: #0f172a; padding: 14px; }
  .container { max-width: 1100px; margin: 0 auto; background: #fff; border-radius: 14px; padding: 18px; box-shadow: 0 4px 20px rgba(0,0,0,0.06); }
  h1 { margin: 0 0 16px; font-size: 20px; color: #1e3a8a; }
  .controls { display: flex; flex-wrap: wrap; gap: 12px; align-items: end; background: #f8fafc; padding: 14px; border-radius: 10px; border: 1px solid #e2e8f0; }
  .field { display: flex; flex-direction: column; gap: 6px; flex: 1; min-width: 220px; }
  label { font-size: 13px; font-weight: 600; color: #334155; }
  input[type="file"] { padding: 9px 10px; border: 1px solid #cbd5e1; border-radius: 8px; background: #fff; font-size: 14px; width: 100%; }
  button { padding: 10px 18px; border: none; border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer; }
  .btn-start { background: #2563eb; color: #fff; }
  .btn-start:disabled { background: #94a3b8; cursor: not-allowed; }
  .btn-stop { background: #dc2626; color: #fff; display: none; }
  .btn-dl { background: #16a34a; color: #fff; display: none; }
  .progress-wrap { margin: 16px 0; display: none; }
  .bar-bg { width: 100%; height: 14px; background: #e2e8f0; border-radius: 99px; overflow: hidden; }
  .bar-fill { height: 100%; width: 0%; background: #2563eb; transition: width 0.25s; }
  .stats { display: flex; gap: 14px; margin-top: 10px; font-size: 13px; font-weight: 600; flex-wrap: wrap; }
  .ok { color: #16a34a; } .fail { color: #dc2626; }
  .table-wrap { margin-top: 16px; max-height: 440px; overflow: auto; border: 1px solid #e2e8f0; border-radius: 8px; -webkit-overflow-scrolling: touch; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th, td { padding: 9px 10px; border-bottom: 1px solid #e2e8f0; text-align: right; white-space: nowrap; }
  th { background: #f8fafc; position: sticky; top: 0; color: #1e293b; }
  .badge-ok { background: #dcfce7; color: #166534; padding: 2px 8px; border-radius: 99px; font-weight: 600; }
  .badge-fail { background: #fee2e2; color: #991b1b; padding: 2px 8px; border-radius: 99px; font-weight: 600; }
  .err { color: #dc2626; margin-top: 10px; font-weight: 600; }
</style>
</head>
<body>
<div class="container">
  <h1>📊 أداة استخراج تقرير تقدم طلاب QUREO</h1>
  <div class="controls">
    <div class="field">
      <label>ملف الطلاب (Excel أو CSV)</label>
      <input type="file" id="csvFile" accept=".xlsx,.xls,.csv" />
    </div>
    <button class="btn-start" id="startBtn" onclick="startCheck()">ابدأ الفحص</button>
    <button class="btn-stop" id="stopBtn" onclick="stopCheck()">إيقاف</button>
    <button class="btn-dl" id="dlBtn" onclick="window.location='/api/download'">⬇️ تحميل التقرير (Excel)</button>
  </div>
  <div id="errBox" class="err"></div>
  <div class="progress-wrap" id="progWrap">
    <div class="bar-bg"><div class="bar-fill" id="barFill"></div></div>
    <div class="stats">
      <span id="stCount">0 / 0</span>
      <span class="ok" id="stOk">نجح: 0</span>
      <span class="fail" id="stFail">فشل: 0</span>
      <span id="stTime">الوقت: 0 ثانية</span>
    </div>
  </div>
  <div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>#</th><th>الاسم</th><th>الفصل</th><th>رقم الطالب</th><th>الحالة</th>
          <th>الفصول</th><th>النسبة %</th><th>الألماس</th><th>آخر نشاط</th>
        </tr>
      </thead>
      <tbody id="tbody"></tbody>
    </table>
  </div>
</div>
<script>
let timer = null;

async function startCheck() {
  const file = document.getElementById('csvFile').files[0];
  const errBox = document.getElementById('errBox');
  errBox.textContent = '';

  if (!file) { errBox.textContent = 'يرجى اختيار ملف الطلاب (Excel أو CSV) أولاً.'; return; }

  const arrayBuf = await file.arrayBuffer();
  const bytes = new Uint8Array(arrayBuf);
  let binary = '';
  for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i]);
  const fileBase64 = btoa(binary);

  const res = await fetch('/api/start', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ fileBase64, fileName: file.name }),
  });
  if (!res.ok) { errBox.textContent = await res.text(); return; }

  document.getElementById('startBtn').disabled = true;
  document.getElementById('stopBtn').style.display = 'inline-block';
  document.getElementById('dlBtn').style.display = 'none';
  document.getElementById('progWrap').style.display = 'block';
  poll();
  timer = setInterval(poll, 500);
}

async function stopCheck() {
  await fetch('/api/stop', { method: 'POST' });
}

async function poll() {
  const res = await fetch('/api/status');
  const s = await res.json();
  const pct = s.total ? Math.round((s.done / s.total) * 100) : 0;
  document.getElementById('barFill').style.width = pct + '%';
  document.getElementById('stCount').textContent = s.done + ' / ' + s.total + ' (' + pct + '%)';
  document.getElementById('stOk').textContent = 'نجح: ' + s.okCount;
  document.getElementById('stFail').textContent = 'فشل: ' + s.failCount;
  document.getElementById('stTime').textContent = 'الوقت: ' + s.elapsedSec + ' ثانية';
  if (s.error) document.getElementById('errBox').textContent = 'خطأ: ' + s.error;

  const html = s.rows.map((r, idx) => {
    if (!r) return '';
    const badge = r.status === 'ok' ? '<span class="badge-ok">ok</span>' : '<span class="badge-fail">' + r.status + '</span>';
    return '<tr><td>' + (idx + 1) + '</td><td>' + r.name + '</td><td>' + r.className + '</td><td>' + r.studentId +
      '</td><td>' + badge + '</td><td>' + r.chaptersCount + '</td><td>' + (r.pct !== '' ? r.pct + '%' : '') +
      '</td><td>' + r.diamonds + '</td><td>' + r.lastActivity + '</td></tr>';
  }).join('');
  document.getElementById('tbody').innerHTML = html;

  if (!s.running) {
    clearInterval(timer);
    document.getElementById('startBtn').disabled = false;
    document.getElementById('stopBtn').style.display = 'none';
    if (s.hasReport) document.getElementById('dlBtn').style.display = 'inline-block';
  }
}
</script>
</body>
</html>`;

const server = http.createServer((req, res) => {
  const reqUrl = new URL(req.url, `http://${req.headers.host || 'localhost'}`);

  if (req.method === 'GET' && reqUrl.pathname === '/') {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    return res.end(HTML);
  }
  if (req.method === 'POST' && reqUrl.pathname === '/api/start') {
    if (state.running) {
      res.writeHead(409, { 'Content-Type': 'text/plain; charset=utf-8' });
      return res.end('الفحص يعمل بالفعل');
    }
    let body = '';
    req.on('data', (chunk) => (body += chunk));
    req.on('end', () => {
      try {
        const { fileBase64, fileName } = JSON.parse(body);
        runJob(fileBase64, fileName);
        res.writeHead(200);
        res.end('ok');
      } catch (e) {
        res.writeHead(400);
        res.end(e.message);
      }
    });
    return;
  }
  if (req.method === 'POST' && reqUrl.pathname === '/api/stop') {
    state.stopRequested = true;
    res.writeHead(200);
    return res.end('stopping');
  }
  if (req.method === 'GET' && reqUrl.pathname === '/api/status') {
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    return res.end(
      JSON.stringify({
        running: state.running,
        total: state.total,
        done: state.done,
        okCount: state.okCount,
        failCount: state.failCount,
        elapsedSec: state.elapsedSec,
        error: state.error,
        hasReport: Boolean(state.xlsxBuffer),
        rows: state.rows,
      })
    );
  }
  if (req.method === 'GET' && reqUrl.pathname === '/api/download') {
    res.writeHead(200, {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="report.xlsx"',
    });
    return res.end(state.xlsxBuffer);
  }
  res.writeHead(404);
  res.end('Not Found');
});

server.listen(PORT, '0.0.0.0', () => {
  const url = `http://127.0.0.1:${PORT}`;
  console.log('====================================================');
  console.log(` QUREO Report Server running on port ${PORT}`);
  console.log('====================================================');
  if (!process.env.PORT && process.platform === 'win32') exec(`start "" "${url}"`);
});