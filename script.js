// 處理 GitHub 前端發來的 POST 請求（提交資料）
function doPost(e) {
  var response;
  try {
    var rawData = e.postData.contents;
    var formData = JSON.parse(rawData);
    var action = formData.action;

    if (action === "submitOrder") {
      response = submitOrder(formData.data);
    } else if (action === "submitFeedback") {
      response = submitFeedback(formData.data);
    } else {
      response = { success: false, error: "未知的 action 動作" };
    }
  } catch (err) {
    response = { success: false, error: err.toString() };
  }

  // 設定 CORS 回傳，允許 GitHub 跨網域讀取
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// 處理 GitHub 前端發來的 GET 請求（讀取列表資料）
function doGet(e) {
  var action = e.parameter.action;
  var response = [];

  if (action === "getOrderList") {
    response = getOrderList();
  } else if (action === "getFeedbackList") {
    response = getFeedbackList();
  }

  // 設定 CORS 回傳
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheetByName(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

// 提交「輸入頁面」資料
function submitOrder(formData) {
  var sheet = getSheetByName("排單系統");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["時間", "暱稱/FB姓名", "委託類型", "委託字數", "付款方式", "指定日期", "委託進度"]);
  }
  
  var timestamp = new Date();
  var timeStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  
  var name = formData.name;
  var type = formData.type;
  var words = formData.words;
  var payment = formData.payment;
  var requiredDate = formData.requiredDate ? formData.requiredDate : "無";
  var status = "排單中";
  
  sheet.appendRow([timeStr, name, type, words, payment, requiredDate, status]);
  
  var estDate = new Date(timestamp.getTime() + 14 * 24 * 60 * 60 * 1000);
  var estDateStr = estDate.getFullYear() + "/" + (estDate.getMonth() + 1) + "/" + estDate.getDate();
  
  return {
    success: true,
    data: {
      name: name,
      type: type,
      words: words,
      payment: payment,
      requiredDate: requiredDate,
      estDate: estDateStr,
      status: status
    }
  };
}

// 取得「排單頁面」列表
function getOrderList() {
  var sheet = getSheetByName("排單系統");
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  var values = sheet.getRange(2, 1, lastRow - 1, 7).getDisplayValues();
  var orders = [];
  
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var timestamp = new Date(row[0]);
    var estDateStr = "無";
    if(!isNaN(timestamp.getTime())) {
      var estDate = new Date(timestamp.getTime() + 14 * 24 * 60 * 60 * 1000);
      estDateStr = estDate.getFullYear() + "/" + (estDate.getMonth() + 1) + "/" + estDate.getDate();
    }
    
    orders.push({
      name: row[1],
      type: row[2],
      words: row[3],
      payment: row[4],
      requiredDate: row[5] ? row[5] : "無",
      estDate: estDateStr,
      status: row[6] ? row[6] : "排單中"
    });
  }
  return orders;
}

// 提交「委託回饋」資料
function submitFeedback(formData) {
  var sheet = getSheetByName("委託回饋");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Message Date", "User", "Level", "Message"]);
  }
  
  var timestamp = new Date();
  var dateStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy/MM/dd");
  
  sheet.appendRow([dateStr, formData.user, formData.level, formData.message]);
  return { success: true };
}

// 取得所有「委託回饋」
function getFeedbackList() {
  var sheet = getSheetByName("委託回饋");
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  var values = sheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
  var feedbacks = [];
  
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    feedbacks.push({
      date: row[0],
      user: row[1],
      level: row[2],
      message: row[3]
    });
  }
  return feedbacks;
}