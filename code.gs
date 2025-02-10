// 當試算表開啟時，在選單中新增「產生學期行事曆」
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('學期行事曆')
    .addItem('產生學期行事曆', 'generateCalendar')
    .addToUi();
}

function generateCalendar() {
  var ui = SpreadsheetApp.getUi();
  // 取得當前日期，作為對話框預設值
  var now = new Date();
  var defaultDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd");
  
  // 依序取得使用者輸入的起始與結束日期，預設值皆填入當下日期的年度、月份、日期
  // 讓使用者輸入起始與結束日期，格式為 YYYYMMDD
  var startInput = ui.prompt('請輸入起始日期 (YYYYMMDD)', '例如：' + defaultDate, ui.ButtonSet.OK).getResponseText();
  var endInput = ui.prompt('請輸入結束日期 (YYYYMMDD)', '例如：' + defaultDate, ui.ButtonSet.OK).getResponseText();
  
  // 確保輸入為8位數字
  if (!/^\d{8}$/.test(startInput) || !/^\d{8}$/.test(endInput)) {
    ui.alert("請輸入正確的日期格式 (YYYYMMDD)");
    return;
  }
  
  // 解析使用者輸入的日期
  var startYear = parseInt(startInput.substring(0, 4), 10);
  var startMonth = parseInt(startInput.substring(4, 6), 10) - 1;
  var startDay = parseInt(startInput.substring(6, 8), 10);
  var endYear = parseInt(endInput.substring(0, 4), 10);
  var endMonth = parseInt(endInput.substring(4, 6), 10) - 1;
  var endDay = parseInt(endInput.substring(6, 8), 10);
  
  var startDate = new Date(startYear, startMonth, startDay);
  var endDate = new Date(endYear, endMonth, endDay);
  
  // 取得目前作用的工作表並清除所有內容
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  
  // 先產生行事曆內容（不包含教師說明列）
  // 先把標題列放在第一列，資料從第二列開始
  
  // 標題列
  var headers = ["週次", "月份", "日", "一", "二", "三", "四", "五", "六", "課程目標", "課程簡介", "重要事項"];
  sheet.appendRow(headers);
  
  // 美化標題列：藍底白字、粗體、置中
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4a86e8')
             .setFontColor('white')
             .setFontWeight('bold')
             .setHorizontalAlignment('center')
             .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 30);
  
  // 取得臺灣假日日曆，建立日期對應假日名稱的物件
  var holidayCalendar = CalendarApp.getCalendarById("zh-tw.taiwan#holiday@group.v.calendar.google.com");
  var holidays = {};
  if (holidayCalendar) {
    var holidayEvents = holidayCalendar.getEvents(startDate, endDate);
    for (var i = 0; i < holidayEvents.length; i++) {
      var event = holidayEvents[i];
      var holidayDate = Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (holidays[holidayDate]) {
        holidays[holidayDate] += "; " + event.getTitle();
      } else {
        holidays[holidayDate] = event.getTitle();
      }
    }
  } else {
    Logger.log("找不到臺灣假日日曆，請確認行事曆ID與訂閱狀態");
  }
  
  // 產生行事曆資料（以週為單位），資料從第二列開始
  // 先將起始日期回溯至該週的星期日
  var current = new Date(startDate);
  current.setDate(current.getDate() - current.getDay());
  var weekNumber = 1;
  var rows = [];
  
  while (current <= endDate) {
    var row = [];
    // 週次
    row[0] = weekNumber;
    
    // 取得該週的月份：取本週內第一個落在範圍內的日期之月份
    var weekMonth = "";
    for (var d = 0; d < 7; d++) {
      var temp = new Date(current);
      temp.setDate(current.getDate() + d);
      if (temp >= startDate && temp <= endDate) {
        weekMonth = temp.getMonth() + 1;
        break;
      }
    }
    row[1] = weekMonth;
    
    // 產生星期日到星期六的日期 (若超出範圍則留空)
    for (var d = 0; d < 7; d++) {
      var dayDate = new Date(current);
      dayDate.setDate(current.getDate() + d);
      if (dayDate >= startDate && dayDate <= endDate) {
        row[2 + d] = dayDate.getDate();
      } else {
        row[2 + d] = "";
      }
    }
    
    // 預留「課程目標」與「課程簡介」欄 (分別為第10與第11欄)
    row[9] = "";
    row[10] = "";
    
    // 將該週內的假日資訊填入「重要事項」欄 (第12欄)
    var impText = "";
    for (var d = 0; d < 7; d++) {
      var dayDate = new Date(current);
      dayDate.setDate(current.getDate() + d);
      if (dayDate >= startDate && dayDate <= endDate) {
        var dateStr = Utilities.formatDate(dayDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        if (holidays[dateStr]) {
          impText += Utilities.formatDate(dayDate, Session.getScriptTimeZone(), "M/d") + " " + holidays[dateStr] + "; ";
        }
      }
    }
    row[11] = impText;
    
    rows.push(row);
    weekNumber++;
    current.setDate(current.getDate() + 7);
  }
  
  // 將所有週資料一次性寫入工作表，從第二列開始（因為第一列已是標題）
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  
  // 統一設定整個行事曆區域（包含標題與資料）的字型、字級與置中
  var totalRows = sheet.getLastRow();
  var totalCols = headers.length;
  var calendarRange = sheet.getRange(1, 1, totalRows, totalCols);
  calendarRange.setFontFamily("微軟正黑體")
               .setFontSize(12)
               .setHorizontalAlignment('center')
               .setVerticalAlignment('middle');
  
  // 設定指定欄位寬度：
  sheet.setColumnWidth(1, 60);  // 週次
  sheet.setColumnWidth(2, 60);  // 月份
  // 課程目標、課程簡介、重要事項調整為 200 像素
  sheet.setColumnWidth(10, 200);
  sheet.setColumnWidth(11, 200);
  sheet.setColumnWidth(12, 200);
  
  // 將「課程目標」、「課程簡介」、「重要事項」 (第10～12欄，自第一列起) 設為靠左並自動換行
  var leftAlignRange = sheet.getRange(1, 10, totalRows, 3);
  leftAlignRange.setHorizontalAlignment('left');
  leftAlignRange.setWrap(true);
  
  // 設定內外框線：
  // 先以細線設定整個行事曆區域
  calendarRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  // 再以粗線設定該區域的外框（僅更改上、下、左、右邊界）
  calendarRange.setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  // 完成行事曆內容後，先於最上方插入一列作為教師說明列
  sheet.insertRowBefore(1);
  // 取得第一列範圍、合併為一個儲存格，並設定預設文字為「教師說明」
  var teacherRange = sheet.getRange(1, 1, 1, totalCols);
  teacherRange.merge();
  teacherRange.setValue("教師說明");
  teacherRange.setBackground("white");
  teacherRange.setFontColor("black");
  // 設定教師說明列為靠左且靠上對齊，並啟用自動換行
  teacherRange.setHorizontalAlignment("left");
  teacherRange.setVerticalAlignment("top");
  teacherRange.setWrap(true);
  sheet.setRowHeight(1, 200);
  
  // 凍結前兩列（教師說明列與原標題列）
  sheet.setFrozenRows(2);
  
  // 調整每個週中日期儲存格格式 (日期在第3～9欄)
  // 現在資料列從第三列開始，故使用 r+3 作為行索引
  for (var r = 0; r < rows.length; r++) {
    // 計算該週的起始日期：使用 startDate 回溯到星期日後，加上 7*r 天
    var weekStart = new Date(startDate);
    weekStart.setDate(weekStart.getDate() - weekStart.getDay() + 7 * r);
    for (var d = 0; d < 7; d++) {
      var dayDate = new Date(weekStart);
      dayDate.setDate(weekStart.getDate() + d);
      if (dayDate < startDate || dayDate > endDate) continue;
      var dayOfWeek = dayDate.getDay();
      var dateStr = Utilities.formatDate(dayDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      // 資料從第三列開始，日期欄從第3欄起，故使用 r+3
      var cell = sheet.getRange(r + 3, 3 + d);
      if (holidays[dateStr]) {
        // 假日：僅將文字設為紅色，不改變背景
        cell.setFontColor("red");
      } else {
        if (dayOfWeek === 6 || dayOfWeek === 0) {
          cell.setFontColor("red");
        } else {
          cell.setFontColor("black");
        }
      }
    }
  }
  
  // 加入交替條件格式 (從資料列第三列開始)
  totalRows = sheet.getLastRow();
  var dataRange = sheet.getRange(3, 1, totalRows - 2, totalCols);
  var rule = SpreadsheetApp.newConditionalFormatRule()
              .whenFormulaSatisfied('=ISODD(ROW())')
              .setBackground('#f2f2f2')
              .setRanges([dataRange])
              .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
  
  SpreadsheetApp.flush();
}
