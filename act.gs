const GITHUB_TOKEN = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN'); 
const GITHUB_REPO = 'mycsd/mycsd';
const GITHUB_FILEPATH = 'activities.json';
const BRANCH = 'main';

/**
 * 格式化日期 -> yyyy-MM-dd
 */
function formatDate(date) {
  if (!date) return "";
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

/**
 * 格式化时间，兼容 Google Sheets 的三种情况
 */
function formatTime(time) {
  if (!time) return "";
  
  // Date 对象（来自 Google Sheets）
  if (time instanceof Date) {
    const correctedTime = new Date(time.getTime() + 42 * 60 * 1000); // 补偿42分钟
    return Utilities.formatDate(correctedTime, "Asia/Singapore", "HH:mm");
  }
  
  // 字符串（如 "10:00AM"）
  if (typeof time === "string") {
    const normalized = time.replace(/\s+/g, '').toUpperCase();
    const period = normalized.includes("PM") ? "PM" : "AM";
    const timePart = normalized.replace(/(AM|PM)/, "");
    const [h, m = "0"] = timePart.split(":");
    
    let hours = parseInt(h, 10);
    const minutes = parseInt(m, 10);
    
    if (period === "PM" && hours < 12) hours += 12;
    if (period === "AM" && hours === 12) hours = 0;
    
    return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
  }
  
  // 数字格式（Google Sheets 内部表示）
  if (typeof time === "number") {
    const totalMinutes = Math.round((time * 24 * 60) + 42); // 补偿42分钟
    const hours = Math.floor(totalMinutes / 60) % 24;
    const minutes = totalMinutes % 60;
    return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
  }
  
  return "Invalid Time";
}

/**
 * 主函数：提取数据并上传到 GitHub
 */
function uploadToGitHub_mycsd() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Future");
  const data = sheet.getRange("B4:J" + sheet.getLastRow()).getValues();
  
  let todayEvents = [];
  let upcomingEvents = [];

  data.forEach(row => {
    const status = row[0];     // B列: TODAY! 或 INCOMING
    const date = formatDate(row[2]);   // C列
    const eventName = row[3];          // D列
    const startTime = formatTime(row[4]); // E列
    const endTime = formatTime(row[5]);   // F列
    const venue = row[6];              // G列
    const fee = row[7];                // H列
    const regLink = row[8];            // J列

    if (status === "TODAY!") {
      todayEvents.push({
        "date": date,
        "event name": eventName,
        "time": startTime,
        "end time": endTime,
        "venue": venue,
        "fee": fee,
        "registration link": regLink,
        "status": "today"
      });
    } else if (status === "INCOMING") {
      upcomingEvents.push({
        "date": date,
        "event name": eventName,
        "time": startTime,
        "end time": endTime,
        "venue": venue,
        "fee": fee,
        "registration link": regLink,
        "status": "upcoming"
      });
    }
  });

  const jsonContent = JSON.stringify({
    today_events: todayEvents,
    upcoming_events: upcomingEvents
  }, null, 2);

  Logger.log("最终JSON内容:\n" + jsonContent);

  // 上传 GitHub
  const apiUrl = `https://api.github.com/repos/${GITHUB_REPO}/contents/${GITHUB_FILEPATH}`;
  const options = {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${GITHUB_TOKEN}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      message: "自动更新事件数据",
      content: Utilities.base64Encode(Utilities.newBlob(jsonContent).getBytes()),
      sha: getCurrentSha()
    })
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log("更新成功: " + response.getContentText());
  } catch (e) {
    Logger.log("更新失败: " + e.message);
  }
}

/**
 * 获取 GitHub 上现有文件的 SHA（更新文件必须提供）
 */
function getCurrentSha() {
  const apiUrl = `https://api.github.com/repos/${GITHUB_REPO}/contents/${GITHUB_FILEPATH}`;
  const response = UrlFetchApp.fetch(apiUrl, {
    headers: { 'Authorization': `Bearer ${GITHUB_TOKEN}` }
  });
  return JSON.parse(response.getContentText()).sha;
}
