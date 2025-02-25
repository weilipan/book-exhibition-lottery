function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('📚 書展抽獎程式')
        .addItem('🔹 產生表單與試算表', 'createFormAndSheet')
        .addItem('🎲 進行抽獎', 'runLottery')
        .addItem('📜 顯示得獎名單', 'showWinners')
        .addItem('🗑️ 清除中獎名單', 'clearWinners')
        .addItem('✉️ 寄送中獎名單', 'emailWinners')
        .addItem('📝 完成簽領單', 'generateSignOffSheetAndEmail')
        .addToUi();
}

// 📌 產生 Google Form 並連結 Google Sheets
function createFormAndSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 📌 **1. 先建立 Google Form**
    var form = FormApp.create('📚 書展抽獎報名表');
    form.setDescription('請填寫以下資料以參加書展抽獎');

    // 📌 **2. 設定 Google Form 的預設行為**
    form.setCollectEmail(true); // 📩 收集電子郵件地址
    form.setAllowResponseEdits(false); // ❌ 不允許編輯回覆
    form.setLimitOneResponsePerUser(true); // 🔐 限制每人僅可填寫 1 次
    form.setAcceptingResponses(true); // ✅ 允許提交回應
    form.setPublishingSummary(false); // ❌ 不公開統計摘要
    form.setConfirmationMessage("✅ 您的表單已成功提交！感謝您的參與。\n📚 書展抽獎結果將另行通知。"); // 📩 自動回覆訊息

    // 📌 **3. 提醒使用者手動設定機構限制**
    Browser.msgBox(
        "⚠️ 請手動設定「僅限臺北市立建國高級中學及其信任機構中的使用者」\n\n"
        + "步驟：\n1️⃣ 開啟 Google Forms\n2️⃣ 點選「設定」\n3️⃣ 選擇「限制使用者範圍」\n"
        + "4️⃣ 勾選「僅限臺北市立建國高級中學及其信任機構中的使用者」"
    );

    // 📌 **4. 新增表單問題（全部設為必填）**
    var questions = [
        ["抽獎序號 (學號)", FormApp.TextItem],
        ["所屬單位或班級", FormApp.TextItem],
        ["姓名", FormApp.TextItem],
        ["聯絡方式(手機)", FormApp.TextItem],
        ["好書推薦(ISBN)", FormApp.TextItem],
        ["好書推薦(定價)", FormApp.TextItem],
        ["好書推薦(作者)", FormApp.TextItem],
        ["好書推薦(書名)", FormApp.TextItem],
        ["請簡單說明推薦此書之原因", FormApp.TextItem]
    ];

    questions.forEach(q => {
        form.addTextItem().setTitle(q[0]).setRequired(true);
    });

    // 📌 **5. 提醒使用者手動新增「檔案上傳」欄位**
    Browser.msgBox("✅ 表單已建立！請手動新增「好書推薦(封面)」的檔案上傳欄位，並允許 PDF/圖片格式。\n\n表單連結：" + form.getPublishedUrl());

    // 📌 **6. 讓 Google Form 自動產生對應的 Google Sheets**
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

    // 📌 **7. 等待 Google Form 自動產生「表單回應」工作表**
    Utilities.sleep(5000); // 等待 5 秒，確保表單連結完成

    // 📌 **8. 取得 Google Form 自動生成的「表單回應」Sheet**
    var formSheet = ss.getSheets().find(sheet => sheet.getName().includes("表單回應"));
    
    if (!formSheet) {
        Browser.msgBox("⚠️ 找不到「表單回應」工作表，請手動檢查試算表！");
        return;
    }

    // 📌 **9. 刪除所有其他工作表，僅保留「表單回應」**
    var sheets = ss.getSheets();
    sheets.forEach(sheet => {
        if (sheet.getName() !== formSheet.getName()) {
            ss.deleteSheet(sheet);
        }
    });

    // 📌 **10. 設定試算表標題行（包含「中獎名單」）**
    var headers = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];
    headers.push("中獎名單"); // 在最後一欄加上「中獎名單」

    // 📌 **11. 更新標題行**
    formSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // 📌 **12. 凍結標題列**
    formSheet.setFrozenRows(1);

    // 📌 **13. 設定「中獎名單」欄位背景顏色**
    var newColumnIndex = headers.length;
    formSheet.getRange(1, newColumnIndex).setBackground("#FFFF99");

    Browser.msgBox("⚠️ 請手動設定「僅限臺北市立建國高級中學及其信任機構中的使用者」\n\n"
        + "步驟：\n1️⃣ 開啟 Google Forms\n2️⃣ 點選「設定」\n3️⃣ 選擇「限制使用者範圍」\n"
        + "4️⃣ 勾選「僅限臺北市立建國高級中學及其信任機構中的使用者」\n"
        + "✅ 試算表已建立，包含『中獎名單』欄位（不會出現在表單中）。\n所有其他工作表已刪除，僅保留「表單回應」。\n"
        + "✅ 表單已建立！請手動新增「好書推薦(封面)」的檔案上傳欄位，並允許 PDF/圖片格式。\n\n表單連結：" + form.getPublishedUrl());
}

// 抽獎功能
function runLottery() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var ui = SpreadsheetApp.getUi();

  // 檢查目前最大抽獎次數
  var data = sheet.getRange(2, 13, lastRow - 1).getValues();
  var roundNumbers = data.flat().map(Number).filter(Boolean);
  var maxRound = Math.max(0, ...roundNumbers);

  // 提示下一次應該進行的抽獎次數，且不可重複
  var promptText = `目前已進行了 ${maxRound} 次抽獎，請輸入本次抽獎的次數 (必須大於 ${maxRound})`;
  var roundResponse = ui.prompt(promptText);
  var roundNumber = parseInt(roundResponse.getResponseText());

  if (isNaN(roundNumber) || roundNumber <= maxRound) {
    ui.alert(`請輸入一個大於 ${maxRound} 的有效數字！`);
    return;
  }

  var numberResponse = ui.prompt('請輸入本次抽獎的人數');
  var numWinners = parseInt(numberResponse.getResponseText());
  if (isNaN(numWinners)) {
    ui.alert('請輸入有效的人數！');
    return;
  }
  
  // 取得所有資料
  var dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  var fullData = dataRange.getValues();

  // 過濾出還未得獎的學生 (中獎名單欄位應為空，表示未得獎)
  var eligible = [];
  for (var i = 0; i < fullData.length; i++) {
    if (fullData[i][12] === "" || fullData[i][12] === null || fullData[i][12] === undefined) { // 第13欄(中獎名單)
      eligible.push(i);
    }
  }

  // 如果可抽獎人數不足
  if (eligible.length < numWinners) {
    ui.alert('可抽獎的人數不足！可抽獎人數：' + eligible.length);
    return;
  }
  
  // 隨機選取中獎者
  var winners = [];
  while (winners.length < numWinners) {
    var randomIndex = Math.floor(Math.random() * eligible.length);
    winners.push(eligible[randomIndex]);
    eligible.splice(randomIndex, 1); // 避免重複選中同一人
  }
  
  // 用來儲存顯示給使用者的結果
  var displayWinners = [];

  // 寫入中獎結果並整理顯示資料
  for (var j = 0; j < winners.length; j++) {
    sheet.getRange(winners[j] + 2, 13).setValue(roundNumber); // 第13欄是"中獎名單"
    var studentNumber = fullData[winners[j]][2]; // 抽獎序號 (第3欄)
    var studentName = anonymizeName(fullData[winners[j]][4]); // 姓名 (第5欄)
    var bookTitle = fullData[winners[j]][9]; // 書名 (第10欄)
    var isbn = fullData[winners[j]][6]; // ISBN (第7欄)
    displayWinners.push(`${studentNumber} ${studentName} ${bookTitle} ${isbn}`);
  }
  
  // 顯示中獎名單
  var resultMessage = displayWinners.join('\n') + `\n\n目前已進行了 ${roundNumber} 次抽獎。恭喜以上得獎的人員！！獲得價值600元以下書展現場展示個人指定書籍乙冊。`;
  ui.alert('本次抽獎結果：\n' + resultMessage);
}

// 顯示得獎名單功能
function showWinners() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var ui = SpreadsheetApp.getUi();

  // 檢查目前已進行的抽獎次數
  var data = sheet.getRange(2, 13, lastRow - 1).getValues();
  var roundNumbers = data.flat().map(Number).filter(Boolean);
  var maxRound = Math.max(0, ...roundNumbers);

  // 讓使用者選擇查看哪一次的得獎名單
  var roundResponse = ui.prompt(`目前已進行了 ${maxRound} 次抽獎，請輸入要查看的抽獎次數`);
  var roundNumber = parseInt(roundResponse.getResponseText());
  if (isNaN(roundNumber) || roundNumber > maxRound) {
    ui.alert(`請輸入一個有效的次數 (1 到 ${maxRound})！`);
    return;
  }

  // 取得所有資料
  var dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  var data = dataRange.getValues();
  
  var winnersList = [];
  
  // 篩選出中獎者
  for (var i = 0; i < data.length; i++) {
    if (data[i][12] == roundNumber) { // 第13欄是"中獎名單"
      var studentNumber = data[i][2]; // 抽獎序號 (第3欄)
      var studentName = anonymizeName(data[i][4]); // 姓名 (第5欄)
      var bookTitle = data[i][9]; // 書名 (第10欄)
      var isbn = data[i][6]; // ISBN (第7欄)
      winnersList.push(`${studentNumber} ${studentName} ${bookTitle} ${isbn}`);
    }
  }

  if (winnersList.length > 0) {
    ui.alert(`第 ${roundNumber} 次抽獎的得獎者：\n` + winnersList.join('\n') + `\n\n目前已進行了 ${maxRound} 次抽獎。恭喜以上得獎的人員！！獲得價值600元以下書展現場展示個人指定書籍乙冊。`);
  } else {
    ui.alert(`沒有找到第 ${roundNumber} 次抽獎的得獎者。`);
  }
}

// 清除中獎名單功能
function clearWinners() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var ui = SpreadsheetApp.getUi();
  
  // 讓使用者選擇清除哪一次的中獎名單
  var roundResponse = ui.prompt('請輸入要清除的抽獎次數（輸入0來清除所有抽獎結果）');
  var roundNumber = parseInt(roundResponse.getResponseText());
  if (isNaN(roundNumber)) {
    ui.alert('請輸入有效的次數！');
    return;
  }

  // 取得中獎名單資料
  var dataRange = sheet.getRange(2, 13, lastRow - 1);  // 這裡應選取中獎名單所在的範圍
  var data = dataRange.getValues();
  
  var clearCount = 0;
  
  // 清除對應次數或全部次數的中獎名單
  if (roundNumber === 0) {
    // 清除所有中獎名單
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] !== '') {
        sheet.getRange(i + 2, 13).setValue('');  // 清除第13欄（中獎名單）
        clearCount++;
      }
    }
    ui.alert(`已清除所有中獎記錄，共 ${clearCount} 名中獎者。`);
  } else {
    // 清除指定次數的中獎名單
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] == roundNumber) {  // 依據中獎次數進行匹配
        sheet.getRange(i + 2, 13).setValue('');  // 清除第13欄（中獎名單）
        clearCount++;
      }
    }
    ui.alert(`已清除第 ${roundNumber} 次抽獎的 ${clearCount} 名中獎者。`);
  }
}
// 寄送中獎名單功能
function emailWinners() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var ui = SpreadsheetApp.getUi();

  // 檢查目前已進行的抽獎次數
  var data = sheet.getRange(2, 13, lastRow - 1).getValues();
  var roundNumbers = data.flat().map(Number).filter(Boolean);
  var maxRound = Math.max(0, ...roundNumbers);
  
  // 讓使用者選擇寄送哪一次的得獎名單
  var roundResponse = ui.prompt(`目前已進行了 ${maxRound} 次抽獎，請輸入要寄送的抽獎次數`);
  var roundNumber = parseInt(roundResponse.getResponseText());
  if (isNaN(roundNumber) || roundNumber > maxRound) {
    ui.alert(`請輸入一個有效的次數 (1 到 ${maxRound})！`);
    return;
  }
  
  // 輸入收件人郵箱
  var emailResponse = ui.prompt('請輸入收件人的電子郵件地址（多個請以逗號隔開）');
  var emailAddresses = emailResponse.getResponseText();
  
  // 取得所有資料
  var dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  var data = dataRange.getValues();
  
  var winnersList = [];
  
  // 篩選出中獎者
  for (var i = 0; i < data.length; i++) {
    if (data[i][12] == roundNumber) { // 第13欄是"中獎名單"
      var studentNumber = data[i][2]; // 抽獎序號 (第3欄)
      var studentName = data[i][4]; // 姓名完整顯示 (第5欄)
      var bookTitle = data[i][9]; // 書名 (第10欄)
      var isbn = data[i][6]; // ISBN (第7欄)
      winnersList.push(`${studentNumber} ${studentName} ${bookTitle} ${isbn}`);
    }
  }

  if (winnersList.length > 0) {
    var currentDate = new Date();
    var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy年MM月dd日');
    var subject = `${formattedDate} 第 ${roundNumber} 次書展抽獎得獎名單`;
    var body = '得獎者名單如下：\n' + winnersList.join('\n') + `\n\n目前已進行了 ${maxRound} 次抽獎。\n恭喜以上得獎的人員！`;
    MailApp.sendEmail(emailAddresses, subject, body);
    ui.alert(`第 ${roundNumber} 次得獎名單已寄送至：${emailAddresses}`);
  } else {
    ui.alert(`沒有找到第 ${roundNumber} 次抽獎的得獎者。`);
  }
}

// 產生簽領單並以Google Docs連結寄送
function generateSignOffSheetAndEmail() {
  var ui = SpreadsheetApp.getUi();
  
  // 輸入學年度、學期、主題書展名、感謝單位、領取期間
  var academicYear = ui.prompt('請輸入學年度').getResponseText();
  var semester = ui.prompt('請輸入學期').getResponseText();
  var bookFairTitle = ui.prompt('請輸入主題書展名').getResponseText();
  var thanksUnits = ui.prompt('請輸入感謝單位（多個單位以逗號分隔）').getResponseText();
  var collectionPeriod = ui.prompt('請輸入領取期間 (如 10/2(一)~10/6(五))').getResponseText();
  
  // 取得要寄送的收件人電子郵件地址
  var emailResponse = ui.prompt('請輸入收件人的電子郵件地址（多個請以逗號隔開）').getResponseText();

  // 取得抽獎次數並按次數由小到大排序
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var dataRange = sheet.getRange(2, 13, lastRow - 1);
  var roundNumbers = [...new Set(dataRange.getValues().flat().map(Number).filter(Boolean))].sort((a, b) => a - b);  // 取得不重複的抽獎次數，並排序
  
  // 使用 Google 文書處理來產生文件
  var doc = DocumentApp.create('簽領單');
  var body = doc.getBody();
  
  // 插入標題和輸入資訊
  body.appendParagraph(`【公告】${academicYear}學年度第${semester}學期${bookFairTitle}得獎名單與領獎通知`);
  body.appendParagraph(`感謝單位：${thanksUnits}`);
  body.appendParagraph(`領取期間：${collectionPeriod}`);
  
  // 插入表格標題
  var table = body.appendTable();
  var headerRow = table.appendTableRow();
  headerRow.appendTableCell('抽獎日期');
  headerRow.appendTableCell('得獎名單與推薦書籍 【序號 姓名 書名 ISBN】');
  headerRow.appendTableCell('簽領/日期時間');

  // 調整第一欄寬度
  table.setColumnWidth(0, 80); // 第一欄寬度設為 80

  // 填入每次抽獎的中獎名單，按抽獎次數排序
  roundNumbers.forEach(function(round, index) {
    var row = table.appendTableRow();
    var winnersList = sheet.getRange(2, 1, lastRow - 1, 13).getValues().filter(function(row) {
      return row[12] == round;
    }).map(function(row) {
      return `${row[2]} ${row[4]}\n${row[9]} ${row[6]}`;  // 姓名和書名換行
    }).join('\n\n'); // 每筆資料之間插入兩行空白
    
    row.appendTableCell(`第 ${round} 次抽獎`);
    row.appendTableCell(winnersList);
    row.appendTableCell('');
  });
  
  doc.saveAndClose();
  
  // 獲取 Google Docs 文檔並設定權限
  var docFile = DriveApp.getFileById(doc.getId());
  docFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // 設定為「知道連結的使用者」只能讀取
  
  // 獲取 Google Docs 文檔連結
  var docUrl = doc.getUrl();
  
  // 寄送文件的連結給指定的收件人
  var subject = `${academicYear}學年度第${semester}學期${bookFairTitle}簽領單`;
  var bodyMessage = `簽領單已生成，請點擊以下連結查看：\n${docUrl}\n\n感謝您的參與！`;
  
  // 發送郵件
  MailApp.sendEmail(emailResponse, subject, bodyMessage);
  
  ui.alert('簽領單已生成並發送至指定的電子郵件地址！');
}

// 將姓名的第二個字替換為O
function anonymizeName(name) {
  if (name.length < 2) return name;
  return name[0] + 'O' + name.slice(2);
}

