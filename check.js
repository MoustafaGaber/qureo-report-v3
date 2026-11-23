// check.js (نسخة 4 — سريعة بالتوازي + حجب الموارد الثقيلة + إغلاق فوري)
// التشغيل:  node check.js          (كل الطلاب)
//           node check.js 10       (أول 10 طلاب للتجربة)
const { chromium } = require('playwright');
const { parse } = require('csv-parse/sync');
const fs = require('fs');

const PORTAL = 'https://me-portal.qureo.education';
const APP = 'https://me-tp.qureo.education';
const CHAPTERS_PATH = '/api/study/students/sections/5/chapters';
const TOTAL_CHAPTERS = 54;
const CONCURRENCY = 4;     // عدد الطلاب في نفس الوقت (يمكن رفعه إلى 5 أو 6 لو الإنترنت سريع)
const REPORT_FILE = 'report.csv';

const limit = parseInt(process.argv[2], 10) || Infinity;
const students = parse(fs.readFileSync('students.csv', 'utf8'), {
  columns: true,
  skip_empty_lines: true,
  bom: true,
}).slice(0, limit);

const csvCell = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`;

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

function debug(s, info) {
  fs.appendFileSync('debug.log', `--- ${s.student_id}\n${JSON.stringify(info, null, 2)}\n`);
}

async function checkStudent(browser, s) {
  const ctx = await browser.newContext();
  const page = await ctx.newPage();
  const info = {};

  try {
    // حجب الصور والخطوط والتنسيقات لتسريع التحميل
    await page.route('**/*', (route) => {
      const type = route.request().resourceType();
      if (['image', 'stylesheet', 'font', 'media'].includes(type)) {
        return route.abort();
      }
      return route.continue();
    });

    // 1) تسجيل الدخول
    const login = await ctx.request.post(`${PORTAL}/oauth2/password-login`, {
      data: { student_id: s.student_id, password: s.password },
      headers: { Referer: `${PORTAL}/login?referer=${encodeURIComponent(PORTAL + '/')}` },
    });
    info.loginStatus = login.status();
    if (!login.ok()) throw new Error(`login HTTP ${login.status()}`);

    // 2) التقاط الاستجابة فور وصولها دون انتظار اكتمال الصفحة
    const responsePromise = new Promise((resolve) => {
      const handler = async (res) => {
        if (res.url().includes(CHAPTERS_PATH) && res.request().method() === 'GET' && res.status() === 200) {
          try {
            const text = await res.text();
            if (text && text.trim()) {
              page.off('response', handler);
              resolve({
                status: res.status(),
                ctype: res.headers()['content-type'] || '',
                body: text,
                source: 'captured from app',
              });
            }
          } catch (err) {
            info.captureError = err.message;
          }
        }
      };
      page.on('response', handler);
      setTimeout(() => {
        page.off('response', handler);
        resolve(null);
      }, 25000);
    });

    // نفتح البوابة ثم التطبيق بدون انتظار networkidle
    await page.goto(`${PORTAL}/`, { waitUntil: 'domcontentloaded' }).catch(() => {});
    const navPromise = page.goto(`${APP}/`, { waitUntil: 'domcontentloaded' }).catch(() => {});

    // بمجرد وصول البيانات نتوقف فورًا
    const capturedData = await responsePromise;
    await navPromise.catch(() => {});
    info.pageUrl = page.url();

    let status, ctype, body;
    if (capturedData) {
      ({ status, ctype, body } = capturedData);
      info.source = capturedData.source;
    } else {
      const r = await ctx.request.get(`${APP}${CHAPTERS_PATH}`);
      status = r.status();
      ctype = r.headers()['content-type'] || '';
      body = await r.text();
      info.source = 'direct request';
    }

    Object.assign(info, { status, ctype, snippet: body.slice(0, 300) });

    if (status !== 200) throw new Error(`chapters HTTP ${status}`);
    let chapters;
    try {
      chapters = JSON.parse(body);
    } catch {
      throw new Error('response is not JSON (see debug.log)');
    }
    if (!Array.isArray(chapters)) throw new Error('unexpected response');

    const dates = chapters.flatMap((c) => [c.started_at, c.completed_at]).filter(Boolean).sort();
    return {
      status: 'ok',
      started: chapters.length,
      completed: chapters.filter((c) => c.completed_at).length,
      diamonds: chapters.reduce((sum, c) => sum + (c.diamond || 0), 0),
      lastActivity: formatEgyptTime(dates.at(-1) || ''),
    };
  } catch (e) {
    debug(s, info);
    return { status: `failed: ${e.message}`, started: '', completed: '', diamonds: '', lastActivity: '' };
  } finally {
    await ctx.close();
  }
}

(async () => {
  const startTime = Date.now();
  const browser = await chromium.launch({ headless: true });
  const header = ['الاسم', 'رقم الطالب', 'الحالة', 'فصول بدأها', 'فصول مكتملة', `من ${TOTAL_CHAPTERS}`, 'النسبة %', 'الألماس', 'آخر نشاط'];

  fs.writeFileSync(REPORT_FILE, '\uFEFF' + header.map(csvCell).join(',') + '\n', 'utf8');

  let nextIndex = 0;
  let doneCount = 0;

  async function worker() {
    while (nextIndex < students.length) {
      const i = nextIndex++;
      const s = students[i];
      const r = await checkStudent(browser, s);
      const pct = r.completed === '' ? '' : Math.round((r.completed / TOTAL_CHAPTERS) * 100);
      const row = [
        s.name || '',
        s.student_id,
        r.status,
        r.started,
        r.completed,
        TOTAL_CHAPTERS,
        pct,
        r.diamonds,
        r.lastActivity,
      ]
        .map(csvCell)
        .join(',');

      fs.appendFileSync(REPORT_FILE, row + '\n', 'utf8');
      doneCount++;
      console.log(`[${doneCount}/${students.length}] ${s.student_id} -> ${r.status} | مكتمل: ${r.completed}`);
    }
  }

  try {
    const workers = Array.from({ length: Math.min(CONCURRENCY, students.length) }, () => worker());
    await Promise.all(workers);
  } finally {
    await browser.close();
  }

  const seconds = ((Date.now() - startTime) / 1000).toFixed(1);
  console.log(`تم حفظ ${REPORT_FILE} بالكامل في ${seconds} ثانية`);
})();