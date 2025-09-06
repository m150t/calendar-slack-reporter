/***** トリガー設定 *****/
function runMonthlyReport() {
  validateConfig_(); // 週次と共通のプロパティ検証

  // 前月: [monthStart, nextMonthStart)
  const { monthStart, nextMonthStart } = getMonthRange_(-1);
  const report = computeMonthlyReport_({ monthStart, nextMonthStart });

  if (OUTPUT_MODE === 'slack' || OUTPUT_MODE === 'both') {
    sendToSlackMonthly_(report);
  }
  if (OUTPUT_MODE === 'sheet' || OUTPUT_MODE === 'both') {
     writeMonthlyTotalsToSheet_(report);
  }
}

/***** 集計本体　*****/
function computeMonthlyReport_({ monthStart, nextMonthStart }) {
  const calendar = getCalendarByName_(CALENDAR_NAME);
  if (!calendar) throw new Error(`カレンダー「${CALENDAR_NAME}」が見つかりません`);

  const timeByCategory = Object.fromEntries(CATEGORIES.map(c => [c, 0]));
  const dailyMap = {}; 

  // 1日ずつループ
  for (let d = new Date(monthStart); d < nextMonthStart; d = new Date(d.getTime() + 86400000)) {
    const dayStart = new Date(d); dayStart.setHours(0,0,0,0);
    const dayEnd   = new Date(dayStart.getTime() + 86400000);

    // 開始日に依存しない取得（複数日/日跨ぎ/終日OK）
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

      // 当日分のオーバーラップのみ加算
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

  const totalHours = Object.values(timeByCategory).reduce((s,h)=>s+h,0);
  return {
    monthStart,
    monthEnd: new Date(nextMonthStart.getTime() - 1), // 表示用
    categories: CATEGORIES.slice(),
    timeByCategory,
    totalHours,
    dailyMap
  };
}

/***** 月範囲（offsetMonths: -1=前月, 0=今月, +1=来月） *****/
function getMonthRange_(offsetMonths = -1) {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const today = new Date(Utilities.formatDate(now, tz, 'yyyy/MM/dd HH:mm:ss'));

  const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  monthStart.setMonth(monthStart.getMonth() + offsetMonths);
  monthStart.setHours(0,0,0,0);

  const nextMonthStart = new Date(monthStart);
  nextMonthStart.setMonth(nextMonthStart.getMonth() + 1);
  nextMonthStart.setHours(0,0,0,0);

  return { monthStart, nextMonthStart };
}

/***** Slack（月次） *****/
function sendToSlackMonthly_(report) {
  if (!SLACK_WEBHOOK_URL) throw new Error('SLACK_WEBHOOK_URL が未設定です。');

  const payload = buildSlackBlocksMonthly_(report);
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

function buildSlackBlocksMonthly_(report) {
  const { monthStart, monthEnd, timeByCategory, totalHours, categories } = report;
  const header = `📅 *月次時間レポート*（${format_(monthStart,'yyyy/MM/dd')}〜${format_(monthEnd,'yyyy/MM/dd')}）\n🕒 合計: *${totalHours.toFixed(1)}h*`;

  const fields = categories.map(cat => {
    const h = timeByCategory[cat] || 0;
    const pct = totalHours > 0 ? ((h/totalHours)*100).toFixed(1) : '0.0';
    const blocks = totalHours > 0 ? Math.round((h/totalHours)*10) : 0;
    const bar = '■'.repeat(blocks) + '□'.repeat(10 - blocks);
    return { type: 'mrkdwn', text: `*${cat}*\n${h.toFixed(1)}h (${pct}%)\n${bar}` };
  });

  return {
    text: '月次時間レポート',
    blocks: [
      { type: 'section', text: { type: 'mrkdwn', text: header } },
      { type: 'divider' },
      { type: 'section', fields }
    ]
  };
}

/***** スプレッドシート出力 *****/
function writeMonthlyTotalsToSheet_(report) {
  if (!SPREADSHEET_ID) throw new Error('SPREADSHEET_ID が未設定です。');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ensureSheet_(ss, 'monthly_totals_long');
    const headers = ['month_start','month_end','category','hours','created_at'];
    ensureHeader_(sheet, headers);

    const monthStartStr = format_(report.monthStart, 'yyyy/MM/dd');
    const monthEndStr   = format_(report.monthEnd,   'yyyy/MM/dd');
    const createdAt     = format_(new Date(), 'yyyy/MM/dd HH:mm:ss');

    const rows = [];
    for (const cat of report.categories) {
      const h = Number(report.timeByCategory[cat] || 0);
      if (Number.isFinite(h) && h > 0) {
        rows.push([monthStartStr, monthEndStr, cat, h, createdAt]);
      }
    }

    if (rows.length) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
    }
  } finally {
    lock.releaseLock();
  }
}
