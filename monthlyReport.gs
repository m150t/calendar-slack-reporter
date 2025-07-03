function sendMonthlyReportToSlack() {
  const SLACK_WEBHOOK_URL = 'https://hooks.slack.com/services/XXXX/XXXX/XXXX'; // Webhook URLを設定
  const CALENDAR_NAME = "業務カレンダー"; // 参照したいカレンダー名を設定
  const CATEGORIES = ["本業", "副業", "勉強", "プライベート"];

  const calendar = getCalendarByName(CALENDAR_NAME);
  if (!calendar) return;

  const now = new Date();
  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  let endOfMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0); 

  const todayStr = formatDate(now);
  const endStr = formatDate(endOfMonth);

  if (todayStr !== endStr) return; // 月末だけ通知

  const events = calendar.getEvents(startOfMonth, endOfMonth);
  const timeByCategory = initializeTimeMap(CATEGORIES);

  for (const event of events) {
    const category = extractCategory(event.getTitle());
    if (category && CATEGORIES.includes(category)) {
      const hours = (event.getEndTime() - event.getStartTime()) / (1000 * 60 * 60);
      timeByCategory[category] += hours;
    }
  }

  const totalHours = Object.values(timeByCategory).reduce((sum, h) => sum + h, 0);

  let message = `📅 今月のカテゴリ別使用時間レポート（${formatDate(startOfMonth)}〜${formatDate(endOfMonth)}）\n`;
  for (const cat of CATEGORIES) {
    const h = timeByCategory[cat].toFixed(1);
    const percent = totalHours > 0 ? `（${((timeByCategory[cat] / totalHours) * 100).toFixed(1)}%）` : '';
    message += `• *${cat}*: ${h}h ${percent}\n`;
  }
  message += `🕒 合計: ${totalHours.toFixed(1)}h`;

  try {
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: message }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log("Slack通知エラー（月次）: " + e.message);
  }
}
