// 行位置定義
// globalでconstが使えない
var rowHeader = 1;
var rowDateOffset = 2;
// 行数
var rowDaysQuantity = 31;
var rowCnt = rowDateOffset + rowDaysQuantity + 0;
var rowAvr = rowDateOffset + rowDaysQuantity + 1;

// 列位置定義
// globalでconstが使えない
var columnDate = 1;
var columnStart = 2;
var columnFinish = 3;
var columnRest = 4;
var columnWorkTime = 5;
var columnNote = 6;
// 列相対位（workTimeからの）
var startR1C1 = columnStart - columnWorkTime;
var finishR1C1 = columnFinish - columnWorkTime;
var restR1C1 = columnRest - columnWorkTime;
var noteR1C1 = columnNote - columnWorkTime;

// 動詞定義
var verbStart = "start";
var verbFinish = "finish";

// 休憩時間デオフォルト値
var restDafault = "1:00:00";

// endpoint
function doGet(e) {
  Logger.log("doGet called");

  var resp = {
    status: "ok"
  };
  return ContentService.createTextOutput(JSON.stringify(resp)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// endpoint
function doPost(e) {
  Logger.log("doPost called");

  var VERIFICATION_TOKEN = PropertiesService.getScriptProperties().getProperty(
    "VERIFICATION_TOKEN"
  );

  var verificationToken = e.parameter.token;
  if (verificationToken != VERIFICATION_TOKEN) {
    throw new Error("Invalid token");
  }

  var texts = e.parameter.text.split(" ");
  var verb = texts[0];
  var work = texts[1];
  var time = texts[2];
  var note = texts[3];
  if (!verb || !work) {
    throw new Error("required verb and work.");
  }
  Logger.log("verb: " + verb);
  Logger.log("work: " + work);
  if (time) {
    Logger.log("time: " + time);
  }
  if (note) {
    Logger.log("note: " + note);
  }

  var date = getDate(time);
  var sheet = getMonthlySheet(work, date);
  write(sheet, date, work, verb, note);

  var url = Utilities.formatString(
    "%s#gid=%s",
    sheet.getParent().getUrl(),
    sheet.getSheetId()
  );
  var resp = Utilities.formatString(
    "saved you %s %s at %s.\n%s",
    verb,
    work,
    date,
    url
  );

  return ContentService.createTextOutput(resp).setMimeType(
    ContentService.MimeType.TEXT
  );
}

// 作業日付を返す
// 現在時間をデフォルトとして、入力した部分を上書きする
function getDate(input) {
  var date = new Date();
  if (!input || input == "now") {
    return date;
  }

  // 年月日と時刻パターン
  // e.g. 2019-12-16_15:00
  const dataMatch = input.match(/[0-9]{4}-[0-9]{2}-[0-9]{2}_[0-9]{2}:[0-9]{2}/);
  // 時刻パターン
  // e.g. 15:00
  const timeMatch = input.match(/[0-9]{2}:[0-9]{2}/);

  if (dataMatch) {
    const year = parseInt(dataMatch[0].substring(0, 4), 10);
    const month = parseInt(dataMatch[0].substring(5, 7), 10) - 1;
    const day = parseInt(dataMatch[0].substring(8, 10), 10);
    const hour = parseInt(dataMatch[0].substring(11, 13), 10);
    const minute = parseInt(dataMatch[0].substring(14, 16), 10);
    Logger.log("year: %s", year);
    Logger.log("month: %s", month);
    Logger.log("day: %s", day);
    Logger.log("hour: %s", hour);
    Logger.log("minute: %s", minute);

    date.setFullYear(year);
    date.setMonth(month);
    date.setDate(day);
    date.setHours(hour, minute, 0);
  } else if (timeMatch) {
    const h = parseInt(timeMatch[0].substring(0, 2), 10);
    const m = parseInt(timeMatch[0].substring(3, 5), 10);
    date.setHours(h, m, 0);
  } else {
    Logger.log("no match");
  }

  return date;
}
// デバッグ
function debugGetDate() {
  var date = getDate("2019-12-31_16:00");
  //var date = getDate('16:00');
  var now = new Date();
  Logger.log(
    "date: %s",
    Utilities.formatDate(date, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss")
  );
  Logger.log(
    "now : %s",
    Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss")
  );
}

// 作業する月次シートを返す
function getMonthlySheet(work, date) {
  const sheetName = Utilities.formatString(
    "%s_%s",
    work,
    Utilities.formatDate(date, "Asia/Tokyo", "yyyy/MM")
  );

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  // 無い場合は作成して初期化
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);

    // ヘッダーと補足を書く
    sheet.getRange(rowHeader, columnDate).setValue("date");
    sheet.getRange(rowHeader, columnStart).setValue(verbStart);
    sheet.getRange(rowHeader, columnFinish).setValue(verbFinish);
    sheet.getRange(rowHeader, columnRest).setValue("rest");
    sheet.getRange(rowHeader, columnWorkTime).setValue("workTime");
    sheet.getRange(rowHeader, columnNote).setValue("note");
    sheet.getRange(rowCnt, columnNote).setValue("count");
    sheet.getRange(rowAvr, columnNote).setValue("average");

    // 表示形式
    sheet
      .getRange(rowDateOffset, columnDate, rowDaysQuantity, 1)
      .setNumberFormat("yyyy/MM/dd");
    sheet
      .getRange(rowDateOffset, columnStart, rowDaysQuantity, 1)
      .setNumberFormat("H:mm:ss");
    sheet
      .getRange(rowDateOffset, columnFinish, rowDaysQuantity, 1)
      .setNumberFormat("H:mm:ss");
    sheet
      .getRange(rowDateOffset, columnRest, rowDaysQuantity, 1)
      .setNumberFormat("[h]:mm:ss");
    sheet
      .getRange(rowDateOffset, columnWorkTime, rowDaysQuantity, 1)
      .setNumberFormat("[h]:mm:ss");

    // 計算式
    for (i = rowDateOffset; i < rowDateOffset + rowDaysQuantity; i++) {
      var formulaWorkTime = Utilities.formatString(
        '=IF(AND(NOT(ISBLANK(R[0]C[%s])), NOT(ISBLANK(R[0]C[%s]))), (R[0]C[%s]-R[0]C[%s]-R[0]C[%s]), "")',
        finishR1C1,
        startR1C1,
        finishR1C1,
        startR1C1,
        restR1C1
      );
      sheet.getRange(i, columnWorkTime).setFormulaR1C1(formulaWorkTime);
    }
    var formulaCnt = Utilities.formatString(
      '=(ROWS(R[%s]C[0]:R[%s]C[0])-COUNTIF(R[%s]C[0]:R[%s]C[0],""))',
      -rowDaysQuantity,
      -1,
      -rowDaysQuantity,
      -1
    );
    sheet.getRange(rowCnt, columnWorkTime).setFormulaR1C1(formulaCnt);
    var formulaAvr = Utilities.formatString(
      '=SUM(R[%s]C[0]:R[%s]C[0])/(ROWS(R[%s]C[0]:R[%s]C[0])-COUNTIF(R[%s]C[0]:R[%s]C[0],""))',
      -rowDaysQuantity - 1,
      -2,
      -rowDaysQuantity - 1,
      -2,
      -rowDaysQuantity - 1,
      -2
    );
    sheet.getRange(rowAvr, columnWorkTime).setFormulaR1C1(formulaAvr);
  }

  return sheet;
}
// デバッグ
function debugMonthlySheet() {
  var now = new Date();
  var sheet = getMonthlySheet("nohana", now);
}

// 所定の日付の動詞に時刻を入力する
function write(sheet, date, work, verb, note) {
  const row = getRowDateIndex(date);

  var column = 0;
  switch (verb) {
    case verbStart:
      column = columnStart;
      // 日付部分も
      sheet.getRange(row, columnDate).setValue(date);
      break;
    case verbFinish:
      column = columnFinish;
      // 休憩部分も
      sheet.getRange(row, columnRest).setValue(restDafault);
      break;
    default:
      throw new Error(Utilities.formatString("undefined verb: %s", verb));
      break;
  }

  sheet.getRange(row, column).setValue(date);
  if (note) {
    sheet.getRange(row, columnNote).setValue(note);
  }
}
// デバッグ
function debugWrite() {
  const date = new Date(2020, 0, 4, 8, 15, 0);
  const work = "nohana";
  const verb = "start";
  const sheet = getMonthlySheet(work, date);

  write(sheet, date, work, verb);
}

// 日付から行位置を返す
function getRowDateIndex(date) {
  const d = date.getDate();

  return d - 1 + rowDateOffset;
}
// デバッグ
function debugGetRotDateIndex() {
  var now = new Date();
  const idx = getRowDateIndex(now);
  Logger.log("index: %s", idx);
}
