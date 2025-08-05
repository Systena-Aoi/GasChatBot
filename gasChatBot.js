function doGet() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('勤怠管理')
    .setTitle('チャットから取得のbot化');
  htmlOutput.append("<script>showLoading();</script>");
  return htmlOutput;
}

var globalExecDate = "";
var globalName = "";
var globalNamePos = "";
var spreadSheetUrl = "https://docs.google.com/spreadsheets/d/1oo_PaBrRj5BxeRPnDwl2cNfviPPOj35dDFB59c2P-Fw/edit?usp=sharing"
var dateInstance;

function dateInstanceToDateStr(dateInstance) {
  var dateStr = Utilities.formatDate(dateInstance, 'JST', 'yyyy/MM/dd');
  return dateStr;
}

function dayOfWeekToKanji(targetDate) {
  const dayOfWeek = targetDate.getDay();
  let dayOfWeekStr = "";

  switch (dayOfWeek) {
    case 0:
      dayOfWeekStr = "日";
      break;
    case 1:
      dayOfWeekStr = "月";
      break;
    case 2:
      dayOfWeekStr = "火";
      break;
    case 3:
      dayOfWeekStr = "水";
      break;
    case 4:
      dayOfWeekStr = "木";
      break;
    case 5:
      dayOfWeekStr = "金";
      break;
    case 6:
      dayOfWeekStr = "土";
      break;
  }
  return dayOfWeekStr;
}

function isWeekendOrHoliday(targetDate) {
  // 土日判定
  const day = targetDate.getDay();
  if (day === 0 || day === 6) {
    return day;
  }

  // 祝日判定
  // 参考：https://zenn.dev/gas/articles/f052a8ecb242e4
  const calendar = CalendarApp.getCalendarById('ja.japanese.official#holiday@group.v.calendar.google.com');
  const events = calendar.getEventsForDay(targetDate);
  if (events.length > 0) {
    return 10;
  }
  return day;
}

function processData(data, execDate, operator, spreadsheetUrlInput) {
  execDate = execDate.replace(/-/g, '/');
  const processedData = data.split("\n");
  spreadSheetUrl = spreadsheetUrlInput;

  var userData = new Map();
  var nextFlag = false;
  var counLine = 0;
  var ampm = 0;
  dateInstance = new Date();

  const result = execDate.match(/(\d+)\/(\d+)\/(\d+)/);
  // 開始日時を作成
  dateInstance = new Date(Number(result[1]), Number(result[2]) - 1, Number(result[3]));
  globalExecDate = dateInstanceToDateStr(dateInstance);

  processedData.forEach(function (key) {
    key = key.replace(', 編集済み', '');

    if ((key.includes('吉村')
      || key.includes('服部')
      || key.includes('相田')
      || key.includes('齊藤')
      || key.includes('佐藤')
      || key.includes('周')
      || key.includes('北村')
      || key.includes('葵')
      || key.includes('自分')) && nextFlag == false) {

      var c1 = key.split(",");
      var c1Len = c1.length;
      globalName = c1[0];
      userData.set('name', globalName);
      if (globalName.includes('吉村')) {
        globalNamePos = 0;
      } else if (globalName.includes('服部')) {
        globalNamePos = 1;
      } else if (globalName.includes('相田')) {
        globalNamePos = 2;
      } else if (globalName.includes('齊藤')) {
        globalNamePos = 3;
      } else if (globalName.includes('佐藤')) {
        globalNamePos = 4;
      } else if (globalName.includes('周')) {
        globalNamePos = 5;
      } else if (globalName.includes('北村')) {
        globalNamePos = 6;
      } else if (globalName.includes('葵')) {
        globalNamePos = 7;
      } else if (globalName.includes('自分')) {
        if (operator == "0") {
          globalNamePos = 0;
        } else if (operator == "1") {
          globalNamePos = 1;
        } else if (operator == "2") {
          globalNamePos = 2;
        } else if (operator == "3") {
          globalNamePos = 3;
        } else if (operator == "4") {
          globalNamePos = 4;
        } else if (operator == "5") {
          globalNamePos = 5;
        } else if (operator == "6") {
          globalNamePos = 6;
        } else if (operator == "7") {
          globalNamePos = 7;
        }
      }

      if (c1Len > 1) {
        var realTime = "";
        if (c1Len > 2) {
          realTime = c1[2].slice(-5).trim();
        } else {
          realTime = c1[1].slice(-5).trim();
        }
        userData.set('実際時間', realTime);
        nextFlag = true;
        counLine = 1;
      }

    } else if (nextFlag == true && counLine >= 1 && key != "") {
      if (key.includes('開始時刻') || key.includes('終了時刻')) {
        var dakokuTime = key.slice(-5).trim();
        if (key.includes('開始時刻')) {
          userData.set('開始時刻', dakokuTime);
          if (ampm == 1) {
            dateInstance.setDate(dateInstance.getDate() + 1);
            // 次の日が何曜日か、または祝日か確認
            var dayOfWeek = isWeekendOrHoliday(dateInstance);
            if (dayOfWeek == 6) {
              dateInstance.setDate(dateInstance.getDate() + 2);
            }
            dayOfWeek = isWeekendOrHoliday(dateInstance);
            if (dayOfWeek == 10) {
              dateInstance.setDate(dateInstance.getDate() + 1);
            }
            globalExecDate = dateInstanceToDateStr(dateInstance);
          }
          ampm = 0;
        } else {
          userData.set('終了時刻', dakokuTime);
          ampm = 1;
        }
        counLine = 2;
      } else if (key.includes('開始場所') || key.includes('終了場所')) {
        var place = key.replace('【開始場所】', '').replace('【終了場所】', '')
        if (key.includes('開始場所')) {
          userData.set('開始場所', place);
          if (ampm == 1) {
            dateInstance.setDate(dateInstance.getDate() + 1);
            globalExecDate = dateInstanceToDateStr(dateInstance);
          }
          ampm = 0;
        } else {
          userData.set('終了場所', place);
          ampm = 1;
        }
        counLine = 3;
        nextFlag = false;
      } else if (key.includes('@')) {
        userData.clear();
        nextFlag = false
        counLine = 0;
      }
    } else if (key == "" && counLine >= 1 && userData.size <= 3) {
      userData.clear();
      nextFlag = false
      counLine = 0;
    } else {
      if (userData.size >= 3 && globalNamePos != -1) {
        execSpreadSheet(userData);
        globalNamePos = -1;
      }
      else {
        userData.clear();
        nextFlag = false
        counLine = 0;
      }
    }
  });

  return 200;
}

function execSpreadSheet(userData) {
  var spreadSheet = SpreadsheetApp.openByUrl(spreadSheetUrl);
  let sheet = spreadSheet.getSheetByName("シート1")

  //対象となるシートの最終行を取得
  var lastRow = sheet.getDataRange().getLastRow();
  var executed = false;

  for (var i = 1; i <= lastRow; i++) {
    var row1Value = sheet.getRange(i, 1).getDisplayValue();
    if (row1Value == globalExecDate) {
      if (userData.has("開始時刻")) {
        sheet.getRange(i, 3 * globalNamePos + 4).setValue(userData.get("開始場所"));
        sheet.getRange(i, 3 * globalNamePos + 5).setValue(userData.get("開始時刻"));
        sheet.getRange(i, 3 * globalNamePos + 6).setValue(userData.get("実際時間"));
      } else {
        sheet.getRange(i + 1, 3).setValue("終了");
        sheet.getRange(i + 1, 3 * globalNamePos + 4).setValue(userData.get("終了場所"));
        sheet.getRange(i + 1, 3 * globalNamePos + 5).setValue(userData.get("終了時刻"));
        sheet.getRange(i + 1, 3 * globalNamePos + 6).setValue(userData.get("実際時間"));
      }
      userData.clear();
      executed = true;
      break;
    }
  }

  if (executed == false) {
    var i = lastRow + 1;
    sheet.getRange(i, 1).setValue(globalExecDate);
    sheet.getRange(i, 2).setValue(dayOfWeekToKanji(dateInstance));

    if (userData.has("開始時刻")) {
      sheet.getRange(i, 3).setValue("開始");
      sheet.getRange(i, 3 * globalNamePos + 4).setValue(userData.get("開始場所"));
      sheet.getRange(i, 3 * globalNamePos + 5).setValue(userData.get("開始時刻"));
      sheet.getRange(i, 3 * globalNamePos + 6).setValue(userData.get("実際時間"));
    } else {
      sheet.getRange(i + 1, 3).setValue("終了");
      sheet.getRange(i + 1, 3 * globalNamePos + 4).setValue(userData.get("終了場所"));
      sheet.getRange(i + 1, 3 * globalNamePos + 5).setValue(userData.get("終了時刻"));
      sheet.getRange(i + 1, 3 * globalNamePos + 6).setValue(userData.get("実際時間"));
    }
  }
}

// doPostの呼び出し
function processDataTest() {

  let data = `
`

  let execDate = "2025-07-01";
  let operator = "7";
  spreadsheetUrl = "https://docs.google.com/spreadsheets/d/1oo_PaBrRj5BxeRPnDwl2cNfviPPOj35dDFB59c2P-Fw/edit?usp=sharing";
  processData(data, execDate, operator, spreadsheetUrl);
}