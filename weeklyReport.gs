function sendWeeklyReportToSlack() {
  const SLACK_WEBHOOK_URL = 'https://hooks.slack.com/services/XXXX/XXXX/XXXX';// Webhook URLを設定
  const CALENDAR_NAME = "副業カレンダー"; // 参照したいカレンダー名を設定
  const CATEGORIES = ["本業", "副業", "勉強", "プライベート"];

  const calendar = getCalendarByName(CALENDAR_NAME);
  if (!calendar) return;

  const today = new Date();
  const weekStart = getStartOfWeek(today);
  let weekEnd = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 6);

  const events = calendar.getEvents(weekStart, weekEnd);
  const timeByCategory = initializeTimeMap(CATEGORIES);

  for (const event of events) {
    const category = extractCategory(event.getTitle());
    if (category && CATEGORIES.includes(category)) {
      const durationHours = (event.getEndTime() - event.getStartTime()) / (1000 * 60 * 60);
      timeByCategory[category] += durationHours;
    }
  }
  const totalHours = Object.values(timeByCategory).reduce((sum, h) => sum + h, 0);
  const message = buildSlackMessage(weekStart, weekEnd, timeByCategory, totalHours, CATEGORIES);

  try {
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: message }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log("Slack通知エラー: " + e.message);
  }
}

function getCalendarByName(name) {
  const calendars = CalendarApp.getCalendarsByName(name);
  if (calendars.length === 0) {
    Logger.log(`カレンダー「${name}」が見つかりません`);
    return null;
  }
  return calendars[0];
}

function getStartOfWeek(date) {
  const day = date.getDay();
  const diff = date.getDate() - day + (day === 0 ? -6 : 1); // 月曜始まり
  return new Date(date.getFullYear(), date.getMonth(), diff);
}

function initializeTimeMap(categories) {
  const map = {};
  for (const cat of categories) {
    map[cat] = 0;
  }
  return map;
}

function extractCategory(title) {
  const match = title.match(/\[(.+?)\]/);
  return match ? match[1] : null;
}

function buildSlackMessage(start, end, timeMap, total, categories) {
  let message = `📊 今週の時間レポート（${formatDate(start)}〜${formatDate(end)}）\n`;
  for (const cat of categories) {
    const hours = timeMap[cat].toFixed(1);
    const percentage = total > 0 ? `（${((timeMap[cat] / total) * 100).toFixed(1)}%）` : '';
    message += `• *${cat}*: ${hours}h ${percentage}\n`;
  }
  message += `🕒 合計: ${total.toFixed(1)}h`;
  return message;
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}
