function doGet() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('勤怠管理')
    .setTitle('GAS ボタンと Textarea');
  htmlOutput.append("<script>showLoading();</script>");
  return htmlOutput;
}

var globalExecDate = "";
var globalName = "";
var globalNamePos = "";

function processData(data, execDate) {
  globalExecDate = execDate.replace(/-/g, '/');
  const processedData = data.split("\n");

  var userData = new Map();
  // var name = "";
  var nextFlag = false;
  var counLine = 0;

  processedData.forEach(function (key) {

    if (key.includes('吉村')
      || key.includes('服部')
      || key.includes('相田')
      || key.includes('齊藤')
      || key.includes('佐藤')
      || key.includes('周')
      || key.includes('北村')
      || key.includes('葵')
      || key.includes('自分')) {

      var c1 = key.split(",");
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
        globalNamePos = 8;
      }

      var realTime = c1[1].slice(-5).trim();

      userData.set('実際時間', realTime);
      nextFlag = true;
      counLine = 5;
    } else if (nextFlag == true) {
      if (key.includes('開始時刻')
        || key.includes('終了時刻')
      ) {
        var dakokuTime = key.slice(-5).trim();
        if (key.includes('開始時刻')) {
          userData.set('開始時刻', dakokuTime);
        } else {
          userData.set('終了時刻', dakokuTime);
        }
        counLine--;
      } else if (key.includes('開始場所')
        || key.includes('終了場所')
      ) {
        var place = key.replace('【開始場所】', '').replace('【終了場所】', '')
        if (key.includes('開始場所')) {
          userData.set('開始場所', place);
        } else {
          userData.set('終了場所', place);
        }
        counLine--;
        nextFlag = false;
      }
    } else {
      if (userData.size >= 3 && globalNamePos != -1) {
        execSpreadSheet(userData);
        globalNamePos = -1;
      }
    }
  });

  // 処理完了後にローディング表示を非表示
  // const finalHtmlOutput = HtmlOutput.append("<script>document.getElementById('loading').style.display = 'none';</script>");
  return 200;
}

function execSpreadSheet(userData) {
  var spreadSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1oo_PaBrRj5BxeRPnDwl2cNfviPPOj35dDFB59c2P-Fw/edit?usp=sharing");
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

  let data = `周梓誠, 金 18:09
勤務終了いたします。
【終了時刻】18:00
【終了場所】川崎
【Canbus. 】投入済み
【  セキュリティカード  】持ち帰ります

佐藤光, 金 18:18
勤務終了いたします。
【終了時刻】18:00
【終了場所】川崎
【Canbus. 】投入済み
【  セキュリティカード  】持ち帰ります

服部亜希, 金 18:47
勤務終了いたします。
【終了時刻】18:30
【終了場所】川崎
【Canbus. 】投入済み
【  セキュリティカード  】持ち帰ります

齊藤秦平, 金 19:04
勤務終了いたします。
【終了時刻】18:00
【終了場所】川崎
【Canbus. 】投入済み
【  セキュリティカード  】持ち帰ります

相田羽遥, 金 19:04
勤務終了いたします。
【終了時刻】18:30
【終了場所】川崎
【Canbus. 】投入済み
【  セキュリティカード  】持ち帰ります

北村美沙, 金 19:32
勤務終了いたします。
【終了時刻】19:30
【終了場所】在宅
【Canbus. 】投入済

吉村美鈴, 金 19:39
勤務終了いたします。
【終了時刻】19:30
【終了場所】川崎
【Canbus. 】投入済み
【  セキュリティカード  】持ち帰ります
`

  let execDate = "2025-07-23"
  processData(data, execDate);
}