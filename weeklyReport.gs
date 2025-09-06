/***** 設定（プロパティから読み込み） *****/
const PROPS = PropertiesService.getScriptProperties();

const OUTPUT_MODE = (PROPS.getProperty('OUTPUT_MODE') || 'both').toLowerCase(); // 'slack' | 'sheet' | 'both'
const SLACK_WEBHOOK_URL = PROPS.getProperty('SLACK_WEBHOOK_URL') || '';
const SPREADSHEET_ID = PROPS.getProperty('SPREADSHEET_ID') || '';

const CALENDAR_NAME = PROPS.getProperty('CALENDAR_NAME') || '業務カレンダー';
const CATEGORIES = JSON.parse(PROPS.getProperty('CATEGORIES') ||
  '["MTG","移動","資料作成","問合せ対応"]');

// 終日イベントを含める（既定：true）／方式：'24h' or 'fixed'
const INCLUDE_ALL_DAY = (PROPS.getProperty('INCLUDE_ALL_DAY') || 'true') === 'true';
const ALL_DAY_MODE = PROPS.getProperty('ALL_DAY_MODE') || '24h';
const ALL_DAY_FIXED_HOURS = Number(PROPS.getProperty('ALL_DAY_FIXED_HOURS') || 8);

// プロパティチェック
function validateConfig_() {
  const mode = (OUTPUT_MODE || '').toLowerCase();
  const MODE_SET = new Set(['slack','sheet','both']);
  if (!MODE_SET.has(mode)) throw new Error(`OUTPUT_MODE が不正です: ${OUTPUT_MODE}`);

  if ((mode === 'slack' || mode === 'both') && !SLACK_WEBHOOK_URL) {
    throw new Error('OUTPUT_MODE に slack/both が含まれるのに SLACK_WEBHOOK_URL が未設定です。');
  }
  if ((mode === 'sheet' || mode === 'both') && !SPREADSHEET_ID) {
    throw new Error('OUTPUT_MODE に sheet/both が含まれるのに SPREADSHEET_ID が未設定です。');
  }  
  // CATEGORIES
  if (!Array.isArray(CATEGORIES) || CATEGORIES.length === 0) {
    throw new Error('CATEGORIES は1個以上のカテゴリを含むJSON配列で設定してください。');
  }
  const bad = CATEGORIES.find(c => typeof c !== 'string' || !c.trim());
  if (bad !== undefined) throw new Error('CATEGORIES 内に空文字または非文字列が含まれています。');

  // 終日設定
  if (!['24h','fixed'].includes(ALL_DAY_MODE)) {
    throw new Error(`ALL_DAY_MODE は '24h' か 'fixed' を指定してください。（現在: ${ALL_DAY_MODE}）`);
  }
  if (ALL_DAY_MODE === 'fixed') {
    if (!Number.isFinite(ALL_DAY_FIXED_HOURS)) {
      throw new Error('ALL_DAY_FIXED_HOURS は数値で指定してください。');
    }
    if (ALL_DAY_FIXED_HOURS < 0 || ALL_DAY_FIXED_HOURS > 24) {
      throw new Error('ALL_DAY_FIXED_HOURS は 0〜24 の範囲で指定してください。');
    }
  }
}

/***** トリガー設定 *****/
function runWeeklyReport() {
  validateConfig_();

  const { weekStart, nextWeekStart } = getWeekRange_(-1);
  const report = computeWeeklyReport_({ weekStart, nextWeekStart });

  if (OUTPUT_MODE === 'slack' || OUTPUT_MODE === 'both') sendToSlack_(report);
  if (OUTPUT_MODE === 'sheet' || OUTPUT_MODE === 'both') writeWeeklyLongToSheet_(report);
}

/***** 集計本体 *****/
function computeWeeklyReport_({ weekStart, nextWeekStart }) {
  const calendar = getCalendarByName_(CALENDAR_NAME);
  if (!calendar) throw new Error(`カレンダー「${CALENDAR_NAME}」が見つかりません`);

  const timeByCategory = Object.fromEntries(CATEGORIES.map(c => [c, 0]));
  const dailyMap = {};

  for (let i = 0; i < 7; i++) {
    const dayStart = new Date(weekStart.getTime() + i * 86400000);
    dayStart.setHours(0, 0, 0, 0);
    const dayEnd = new Date(dayStart.getTime() + 86400000);

    // 開始時刻に依存しない1日イベント取得
    const events = calendar.getEventsForDay(dayStart);
    const dateKey = format_(dayStart, 'yyyy/MM/dd');
    if (!dailyMap[dateKey]) dailyMap[dateKey] = {};

    for (const ev of events) {
      const cat = extractCategory_(ev.getTitle(), CATEGORIES);
      if (!cat) continue;

      if (ev.isAllDayEvent()) {
        if (!INCLUDE_ALL_DAY) continue;
        const hours = (ALL_DAY_MODE === 'fixed') ? ALL_DAY_FIXED_HOURS : 24;
        timeByCategory[cat] += hours;
        dailyMap[dateKey][cat] = (dailyMap[dateKey][cat] || 0) + hours;
        continue;
      }

      // 日内オーバーラップのみ加算（同一イベントが複数日に出ても日割りされる）
      const overlapStart = new Date(Math.max(ev.getStartTime().getTime(), dayStart.getTime()));
      const overlapEnd   = new Date(Math.min(ev.getEndTime().getTime(),   dayEnd.getTime()));
      if (overlapEnd > overlapStart) {
        const hours = (overlapEnd - overlapStart) / 3600000;
        if (Number.isFinite(hours) && hours > 0) {
          timeByCategory[cat] += hours;
          dailyMap[dateKey][cat] = (dailyMap[dateKey][cat] || 0) + hours;
        }
      }
    }
  }

  const totalHours = Object.values(timeByCategory).reduce((s, h) => s + h, 0);
  return {
    weekStart,
    weekEnd: new Date(nextWeekStart.getTime() - 1),
    categories: CATEGORIES.slice(),
    timeByCategory,
    totalHours,
    dailyMap
  };
}

/***** Slack 出力 *****/
function sendToSlack_(report) {
  if (!SLACK_WEBHOOK_URL) throw new Error('SLACK_WEBHOOK_URL が未設定です。');

  const payload = buildSlackBlocks_(report);
  const res = UrlFetchApp.fetch(SLACK_WEBHOOK_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`Slack通知失敗 status=${code} body=${res.getContentText()}`);
  }
}

function buildSlackBlocks_(report) {
  const { weekStart, weekEnd, timeByCategory, totalHours, categories } = report;
  const header = `📊 *週次時間レポート*（${format_(weekStart, 'yyyy/MM/dd')}〜${format_(weekEnd, 'yyyy/MM/dd')}）\n🕒 合計: *${totalHours.toFixed(1)}h*`;

  const fields = categories.map(cat => {
    const h = timeByCategory[cat] || 0;
    const pct = totalHours > 0 ? ((h / totalHours) * 100).toFixed(1) : '0.0';
    const blocks = totalHours > 0 ? Math.round((h / totalHours) * 10) : 0;
    const bar = '■'.repeat(blocks) + '□'.repeat(10 - blocks);
    return { type: 'mrkdwn', text: `*${cat}*\n${h.toFixed(1)}h (${pct}%)\n${bar}` };
  });

  return {
    text: '週次時間レポート',
    blocks: [
      { type: 'section', text: { type: 'mrkdwn', text: header } },
      { type: 'divider' },
      { type: 'section', fields }
    ]
  };
}

/***** シート出力 *****/
function writeWeeklyLongToSheet_(report) {
  if (!SPREADSHEET_ID) throw new Error('SPREADSHEET_ID が未設定です。');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // 確実に取得 or 例外
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ensureSheet_(ss, 'weekly_summary_long');
    const headers = ['week_start', 'week_end', 'date', 'category', 'hours', 'created_at'];
    ensureHeader_(sheet, headers);

    const weekStartStr = format_(report.weekStart, 'yyyy/MM/dd');
    const weekEndStr = format_(report.weekEnd, 'yyyy/MM/dd');
    const createdAt = format_(new Date(), 'yyyy/MM/dd HH:mm:ss');

    const rows = [];
    Object.keys(report.dailyMap).forEach(dateStr => {
      const cats = report.dailyMap[dateStr];
      Object.keys(cats).forEach(cat => {
        const h = Number(cats[cat] || 0);
        if (Number.isFinite(h) && h > 0) {
          rows.push([weekStartStr, weekEndStr, dateStr, cat, h, createdAt]);
        }
      });
    });

    if (rows.length) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
    }
  } finally {
    lock.releaseLock();
  }
}

/***** 共通ユーティリティ *****/
function getCalendarByName_(name) {
  const cals = CalendarApp.getCalendarsByName(name);
  return (cals && cals.length) ? cals[0] : null;
}

// 月曜始まり：offsetWeeks=-1 → 前週
function getWeekRange_(offsetWeeks = 0) {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const today = new Date(Utilities.formatDate(now, tz, 'yyyy/MM/dd HH:mm:ss'));

  const dow = today.getDay(); // 0=日
  const diffToMonday = (dow === 0 ? -6 : 1) - dow;

  const base = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const weekStart = new Date(base);
  weekStart.setDate(base.getDate() + diffToMonday + offsetWeeks * 7);
  weekStart.setHours(0, 0, 0, 0);

  const nextWeekStart = new Date(weekStart);
  nextWeekStart.setDate(weekStart.getDate() + 7);
  return { weekStart, nextWeekStart };
}

function ensureSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureHeader_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function extractCategory_(title, categories) {
  if (!title) return null;
  // 先頭の [カテゴリ] / ［カテゴリ］ / 【カテゴリ】 に対応
  const m = title.trim().match(/^(\[([^\]]+)\]|［([^］]+)］|【([^】]+)】)/);
  if (!m) return null;
  const cat = (m[2] || m[3] || m[4] || '').trim();
  return categories.includes(cat) ? cat : null;
}


function format_(date, pattern) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), pattern);
}
