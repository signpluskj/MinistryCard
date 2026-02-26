var SHEET_ID = "1aaQbwBdL0hr_WN47zTtmOuBlq8CbnLwQBnQF93M0ZfY";

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

function mapRows(values) {
  var header = values[0] || [];
  return values.slice(1).map(function (row) {
    var obj = {};
    header.forEach(function (key, idx) {
      obj[key] = row[idx];
    });
    return obj;
  });
}

function headerIndex(values) {
  var header = values[0] || [];
  var map = {};
  header.forEach(function (key, idx) {
    map[key] = idx + 1;
  });
  return map;
}

function formatYYMMDD(value) {
  if (!value) {
    return "";
  }
  var d = new Date(String(value).replace(/\./g, "-").replace(/\//g, "-"));
  if (isNaN(d.getTime())) {
    return value;
  }
  var yy = String(d.getFullYear()).slice(-2);
  var mm = String(d.getMonth() + 1);
  if (mm.length < 2) mm = "0" + mm;
  var dd = String(d.getDate());
  if (dd.length < 2) dd = "0" + dd;
  return yy + "/" + mm + "/" + dd;
}

function findRow(values, headerName, value) {
  var idx = values[0].indexOf(headerName);
  if (idx === -1) {
    return -1;
  }
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idx]) === String(value)) {
      return i + 1;
    }
  }
  return -1;
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || "bootstrap";
  if (action === "bootstrap") {
    var cardsValues = getSheet("구역카드").getDataRange().getValues();
    var areasValues = getSheet("구역번호").getDataRange().getValues();
    var completionsValues = getSheet("완료내역").getDataRange().getValues();
    var visitsRows = getAllVisitRows();
    var evangelistsValues = getSheet("전도인명단").getDataRange().getValues();
    return jsonResponse({
      cards: mapRows(cardsValues),
      areas: mapRows(areasValues),
      completions: mapRows(completionsValues),
      visits: visitsRows,
      evangelists: mapRows(evangelistsValues)
    });
  }
  return jsonResponse({ error: "지원하지 않는 요청입니다." });
}

function doPost(e) {
  var data = e && e.parameter ? e.parameter : {};
  var action = data.action;
  if (action === "login") {
    return handleLogin(data);
  }
  if (action === "startService") {
    return handleStartService(data);
  }
  if (action === "resetRecentVisits") {
    return handleResetRecentVisits(data);
  }
  if (action === "recordVisit") {
    return handleRecordVisit(data);
  }
  if (action === "updateVisit") {
    return handleUpdateVisit(data);
  }
  return jsonResponse({ success: false, message: "지원하지 않는 요청입니다." });
}

function handleLogin(data) {
  var values = getSheet("전도인명단").getDataRange().getValues();
  var rows = mapRows(values);
  var user = rows.find(function (row) {
    return String(row["이름"]) === String(data.name);
  });
  if (!user) {
    return jsonResponse({ success: false, message: "사용자를 찾을 수 없습니다." });
  }
  var role = String(user["역할"] || user["권한"] || "");
  var storedPassword = user["비밀번호"] != null ? String(user["비밀번호"]) : "";
  var inputPassword = data.password != null ? String(data.password) : "";
  if (role === "전도인") {
    if (storedPassword && storedPassword !== inputPassword) {
      return jsonResponse({ success: false, message: "비밀번호가 올바르지 않습니다." });
    }
  } else {
    if (!storedPassword || storedPassword !== inputPassword) {
      return jsonResponse({ success: false, message: "비밀번호가 올바르지 않습니다." });
    }
  }
  return jsonResponse({ success: true, role: role, name: user["이름"] });
}

function handleStartService(data) {
  var sheet = getSheet("구역번호");
  var values = sheet.getDataRange().getValues();
  var headers = headerIndex(values);
  var rowIndex = findRow(values, "구역번호", data.areaId);
  if (rowIndex === -1) {
    sheet.appendRow([data.areaId, formatYYMMDD(data.date), "", data.leaderName]);
  } else {
    if (headers["시작날짜"]) {
      sheet.getRange(rowIndex, headers["시작날짜"]).setValue(formatYYMMDD(data.date));
    }
    if (headers["인도자"]) {
      sheet.getRange(rowIndex, headers["인도자"]).setValue(data.leaderName);
    }
    if (headers["완료날짜"]) {
      sheet.getRange(rowIndex, headers["완료날짜"]).setValue("");
    }
  }
  var cardsSheet = getSheet("구역카드");
  var cardValues = cardsSheet.getDataRange().getValues();
  var cardHeaders = headerIndex(cardValues);
  var areaIdx = cardValues[0].indexOf("구역번호");
  for (var i = 1; i < cardValues.length; i++) {
    if (String(cardValues[i][areaIdx]) === String(data.areaId)) {
      if (cardHeaders["최근방문일"]) {
        cardsSheet.getRange(i + 1, cardHeaders["최근방문일"]).setValue("");
      }
      if (cardHeaders["만남"]) {
        cardsSheet.getRange(i + 1, cardHeaders["만남"]).setValue(false);
      }
      if (cardHeaders["부재"]) {
        cardsSheet.getRange(i + 1, cardHeaders["부재"]).setValue(false);
      }
      if (cardHeaders["초대장"]) {
        cardsSheet.getRange(i + 1, cardHeaders["초대장"]).setValue(false);
      }
    }
  }
  return jsonResponse({ success: true });
}

function handleResetRecentVisits(data) {
  var sheet = getSheet("구역카드");
  var values = sheet.getDataRange().getValues();
  var headers = headerIndex(values);
  var areaIdx = values[0].indexOf("구역번호");
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][areaIdx]) === String(data.areaId)) {
      if (headers["최근방문일"]) sheet.getRange(i + 1, headers["최근방문일"]).setValue("");
      if (headers["만남"]) sheet.getRange(i + 1, headers["만남"]).setValue(false);
      if (headers["부재"]) sheet.getRange(i + 1, headers["부재"]).setValue(false);
      if (headers["초대장"]) sheet.getRange(i + 1, headers["초대장"]).setValue(false);
    }
  }
  return jsonResponse({ success: true });
}

var MAX_VISIT_ROWS = 5000;

function getVisitSheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheets = ss.getSheets();
  var result = [];
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name === "방문내역" || name.indexOf("방문내역") === 0) {
      result.push(sheets[i]);
    }
  }
  result.sort(function (a, b) {
    return a.getName() > b.getName() ? 1 : a.getName() < b.getName() ? -1 : 0;
  });
  return result;
}

function getVisitSheetForAppend() {
  var sheets = getVisitSheets();
  if (sheets.length === 0) {
    return getSheet("방문내역");
  }
  var current = sheets[sheets.length - 1];
  if (current.getLastRow() >= MAX_VISIT_ROWS) {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var base = sheets[0];
    var newName = "방문내역" + (sheets.length + 1);
    var newSheet = ss.insertSheet(newName);
    var headerRange = base.getRange(1, 1, 1, base.getLastColumn());
    newSheet.getRange(1, 1, 1, base.getLastColumn()).setValues(headerRange.getValues());
    return newSheet;
  }
  return current;
}

function getAllVisitRows() {
  var sheets = getVisitSheets();
  var all = [];
  for (var i = 0; i < sheets.length; i++) {
    var values = sheets[i].getDataRange().getValues();
    if (values.length > 1) {
      var rows = mapRows(values);
      all = all.concat(rows);
    }
  }
  return all;
}

function handleRecordVisit(data) {
  var visitsSheet = getVisitSheetForAppend();
  var headerRange = visitsSheet.getRange(1, 1, 1, visitsSheet.getLastColumn());
  var headerValues = headerRange.getValues();
  var visitHeader = headerValues[0] || [];
  var visitRow = {};
  if (visitHeader.length === 0) {
    var fallbackRow = [data.areaId, data.cardNumber, formatYYMMDD(data.visitDate), data.worker, data.result, data.note];
    visitsSheet.appendRow(fallbackRow);
    visitRow = {
      "구역번호": data.areaId,
      "구역카드": data.cardNumber,
      "방문날짜": formatYYMMDD(data.visitDate),
      "전도인": data.worker,
      "결과": data.result,
      "메모": data.note
    };
  } else {
    var visitHeaders = headerIndex(headerValues);
    var row = new Array(visitHeader.length).fill("");
    if (visitHeaders["구역번호"]) row[visitHeaders["구역번호"] - 1] = data.areaId;
    if (visitHeaders["카드번호"]) row[visitHeaders["카드번호"] - 1] = data.cardNumber;
    if (visitHeaders["구역카드"]) row[visitHeaders["구역카드"] - 1] = data.cardNumber;
    if (visitHeaders["방문일"]) row[visitHeaders["방문일"] - 1] = formatYYMMDD(data.visitDate);
    if (visitHeaders["방문날짜"]) row[visitHeaders["방문날짜"] - 1] = formatYYMMDD(data.visitDate);
    if (visitHeaders["전도인"]) row[visitHeaders["전도인"] - 1] = data.worker;
    if (visitHeaders["방문자"]) row[visitHeaders["방문자"] - 1] = data.worker;
    if (visitHeaders["결과"]) row[visitHeaders["결과"] - 1] = data.result;
    if (visitHeaders["방문결과"]) row[visitHeaders["방문결과"] - 1] = data.result;
    if (visitHeaders["메모"]) row[visitHeaders["메모"] - 1] = data.note;
    if (visitHeaders["비고"]) row[visitHeaders["비고"] - 1] = data.note;
    visitsSheet.appendRow(row);
    for (var c = 0; c < visitHeader.length; c++) {
      var key = visitHeader[c];
      if (key) {
        visitRow[key] = row[c];
      }
    }
  }
  var cardsSheet = getSheet("구역카드");
  var cardValues = cardsSheet.getDataRange().getValues();
  var headers = headerIndex(cardValues);
  var areaIdx = cardValues[0].indexOf("구역번호");
  var cardIdx = cardValues[0].indexOf("카드번호");
  var cardUpdate = {
    areaId: data.areaId,
    cardNumber: data.cardNumber,
    recentVisitDate: formatYYMMDD(data.visitDate),
    meet: data.result === "만남",
    absent: data.result === "부재",
    invite: data.result === "초대장",
    sixMonths: data.result === "6개월",
    banned: data.result === "방문금지"
  };
  for (var i = 1; i < cardValues.length; i++) {
    if (String(cardValues[i][areaIdx]) === String(data.areaId) && String(cardValues[i][cardIdx]) === String(data.cardNumber)) {
      if (headers["최근방문일"]) {
        var recent = formatYYMMDD(data.visitDate);
        cardsSheet.getRange(i + 1, headers["최근방문일"]).setValue(recent);
        cardValues[i][headers["최근방문일"] - 1] = recent;
      }
      if (headers["만남"]) {
        var meetBool = data.result === "만남";
        cardsSheet.getRange(i + 1, headers["만남"]).setValue(meetBool);
        cardValues[i][headers["만남"] - 1] = meetBool;
      }
      if (headers["부재"]) {
        var absentBool = data.result === "부재";
        cardsSheet.getRange(i + 1, headers["부재"]).setValue(absentBool);
        cardValues[i][headers["부재"] - 1] = absentBool;
      }
      if (headers["초대장"]) {
        var inviteBool = data.result === "초대장";
        cardsSheet.getRange(i + 1, headers["초대장"]).setValue(inviteBool);
        cardValues[i][headers["초대장"] - 1] = inviteBool;
      }
      if (headers["6개월"]) {
        var sixBool = data.result === "6개월";
        cardsSheet.getRange(i + 1, headers["6개월"]).setValue(sixBool);
        cardValues[i][headers["6개월"] - 1] = sixBool;
      }
      if (headers["방문금지"]) {
        var banBool = data.result === "방문금지";
        cardsSheet.getRange(i + 1, headers["방문금지"]).setValue(banBool);
        cardValues[i][headers["방문금지"] - 1] = banBool;
      }
      break;
    }
  }
  var complete = checkAreaComplete(data.areaId, cardValues);
  if (complete) {
    updateCompletion(data.areaId, formatYYMMDD(data.visitDate), data.leaderName);
  }
  return jsonResponse({ success: true, complete: complete, visit: visitRow, cardUpdate: cardUpdate });
}

function checkAreaComplete(areaId, cardValues) {
  var headers = headerIndex(cardValues);
  var areaIdx = cardValues[0].indexOf("구역번호");
  for (var i = 1; i < cardValues.length; i++) {
    if (String(cardValues[i][areaIdx]) === String(areaId)) {
      var recentValue = headers["최근방문일"] ? cardValues[i][headers["최근방문일"] - 1] : "";
      if (!recentValue) {
        return false;
      }
    }
  }
  return true;
}

function updateCompletion(areaId, completeDate, leaderName) {
  var areaSheet = getSheet("구역번호");
  var areaValues = areaSheet.getDataRange().getValues();
  var areaHeaders = headerIndex(areaValues);
  var areaRowIndex = findRow(areaValues, "구역번호", areaId);
  var startDate = "";
  if (areaRowIndex !== -1) {
    if (areaHeaders["시작날짜"]) {
      startDate = areaSheet
        .getRange(areaRowIndex, areaHeaders["시작날짜"])
        .getValue();
    }
    if (areaHeaders["완료날짜"]) {
      areaSheet.getRange(areaRowIndex, areaHeaders["완료날짜"]).setValue(formatYYMMDD(completeDate));
    }
    if (areaHeaders["인도자"]) {
      areaSheet.getRange(areaRowIndex, areaHeaders["인도자"]).setValue(leaderName);
    }
  }
  var logSheet = getSheet("완료내역");
  logSheet.appendRow([areaId, startDate, formatYYMMDD(completeDate), leaderName]);
}

function normalizeDateText(value) {
  var txt = formatYYMMDD(value);
  return String(txt);
}

function findLatestVisit(areaId, cardNumber) {
  var sheets = getVisitSheets();
  var latest = null;
  var latestResult = "";
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var values = sheet.getDataRange().getValues();
    if (!values.length) {
      continue;
    }
    var headers = headerIndex(values);
    var areaIdx = values[0].indexOf("구역번호");
    var cardIdx = values[0].indexOf("카드번호");
    var cardAltIdx = values[0].indexOf("구역카드");
    var dateIdx =
      values[0].indexOf("방문일") !== -1
        ? values[0].indexOf("방문일")
        : values[0].indexOf("방문날짜") !== -1
        ? values[0].indexOf("방문날짜")
        : values[0].indexOf("방문일자") !== -1
        ? values[0].indexOf("방문일자")
        : values[0].indexOf("날짜") !== -1
        ? values[0].indexOf("날짜")
        : values[0].indexOf("방문일시") !== -1
        ? values[0].indexOf("방문일시")
        : values[0].indexOf("일자");
    var resultIdx =
      values[0].indexOf("결과") !== -1
        ? values[0].indexOf("결과")
        : values[0].indexOf("방문결과");
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      var rowArea = row[areaIdx];
      var rowCard = cardIdx !== -1 ? row[cardIdx] : cardAltIdx !== -1 ? row[cardAltIdx] : "";
      if (String(rowArea) !== String(areaId) || String(rowCard) !== String(cardNumber)) {
        continue;
      }
      var dText = normalizeDateText(row[dateIdx]);
      var parts = dText.split("/");
      if (parts.length !== 3) {
        continue;
      }
      var dateObj = new Date("20" + parts[0] + "-" + parts[1] + "-" + parts[2]);
      if (isNaN(dateObj.getTime())) {
        continue;
      }
      if (latest == null || dateObj.getTime() > latest.getTime()) {
        latest = dateObj;
        latestResult = resultIdx !== -1 ? String(row[resultIdx]) : "";
      }
    }
  }
  return latest
    ? { dateText: formatYYMMDD(latest), resultText: latestResult }
    : { dateText: "", resultText: "" };
}

function updateCardRecentVisit(areaId, cardNumber) {
  var latest = findLatestVisit(areaId, cardNumber);
  var cardsSheet = getSheet("구역카드");
  var cardValues = cardsSheet.getDataRange().getValues();
  var headers = headerIndex(cardValues);
  var areaIdx = cardValues[0].indexOf("구역번호");
  var cardIdx = cardValues[0].indexOf("카드번호");
  for (var i = 1; i < cardValues.length; i++) {
    if (String(cardValues[i][areaIdx]) === String(areaId) && String(cardValues[i][cardIdx]) === String(cardNumber)) {
      if (headers["최근방문일"] && latest.dateText) {
        cardsSheet.getRange(i + 1, headers["최근방문일"]).setValue(latest.dateText);
      } else if (headers["최근방문일"] && !latest.dateText) {
        cardsSheet.getRange(i + 1, headers["최근방문일"]).setValue("");
      }
      if (headers["만남"]) {
        cardsSheet.getRange(i + 1, headers["만남"]).setValue(latest.resultText === "만남");
      }
      if (headers["부재"]) {
        cardsSheet.getRange(i + 1, headers["부재"]).setValue(latest.resultText === "부재");
      }
      if (headers["초대장"]) {
        cardsSheet.getRange(i + 1, headers["초대장"]).setValue(latest.resultText === "초대장");
      }
      if (headers["6개월"]) {
        cardsSheet.getRange(i + 1, headers["6개월"]).setValue(latest.resultText === "6개월");
      }
      if (headers["방문금지"]) {
        cardsSheet.getRange(i + 1, headers["방문금지"]).setValue(latest.resultText === "방문금지");
      }
      break;
    }
  }
  return {
    areaId: areaId,
    cardNumber: cardNumber,
    recentVisitDate: latest.dateText,
    meet: latest.resultText === "만남",
    absent: latest.resultText === "부재",
    invite: latest.resultText === "초대장",
    sixMonths: latest.resultText === "6개월",
    banned: latest.resultText === "방문금지"
  };
}

function handleUpdateVisit(data) {
  var areaId = data.areaId;
  var cardNumber = data.cardNumber;
  var oldDate = data.oldVisitDate;
  var oldWorker = data.oldWorker;
  var newDate = data.newVisitDate || oldDate;
  var newWorker = data.newWorker || oldWorker;
  var newResult = data.newResult || data.result || "";
  var newNote = data.newNote || data.note || "";
  var sheets = getVisitSheets();
  var updatedRowObj = {};
  var found = false;
  for (var i = 0; i < sheets.length && !found; i++) {
    var sheet = sheets[i];
    var values = sheet.getDataRange().getValues();
    if (!values.length) {
      continue;
    }
    var headers = headerIndex(values);
    var areaIdx = values[0].indexOf("구역번호");
    var cardIdx = values[0].indexOf("카드번호");
    var cardAltIdx = values[0].indexOf("구역카드");
    var dateIdx =
      values[0].indexOf("방문일") !== -1
        ? values[0].indexOf("방문일")
        : values[0].indexOf("방문날짜") !== -1
        ? values[0].indexOf("방문날짜")
        : values[0].indexOf("방문일자") !== -1
        ? values[0].indexOf("방문일자")
        : values[0].indexOf("날짜") !== -1
        ? values[0].indexOf("날짜")
        : values[0].indexOf("방문일시") !== -1
        ? values[0].indexOf("방문일시")
        : values[0].indexOf("일자");
    var workerIdx =
      values[0].indexOf("전도인") !== -1
        ? values[0].indexOf("전도인")
        : values[0].indexOf("방문자");
    var resultIdx =
      values[0].indexOf("결과") !== -1
        ? values[0].indexOf("결과")
        : values[0].indexOf("방문결과");
    var noteIdx =
      values[0].indexOf("메모") !== -1
        ? values[0].indexOf("메모")
        : values[0].indexOf("비고");
    var oldDateText = normalizeDateText(oldDate);
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      var rowArea = String(row[areaIdx]);
      var rowCard = String(cardIdx !== -1 ? row[cardIdx] : cardAltIdx !== -1 ? row[cardAltIdx] : "");
      var rowDateText = normalizeDateText(row[dateIdx]);
      var rowWorker = String(workerIdx !== -1 ? row[workerIdx] : "");
      if (rowArea === String(areaId) && rowCard === String(cardNumber) && rowDateText === String(oldDateText) && rowWorker === String(oldWorker)) {
        if (dateIdx !== -1) sheet.getRange(r + 1, dateIdx + 1).setValue(formatYYMMDD(newDate));
        if (workerIdx !== -1) sheet.getRange(r + 1, workerIdx + 1).setValue(newWorker);
        if (resultIdx !== -1) sheet.getRange(r + 1, resultIdx + 1).setValue(newResult);
        if (noteIdx !== -1) sheet.getRange(r + 1, noteIdx + 1).setValue(newNote);
        var header = values[0] || [];
        var obj = {};
        for (var c = 0; c < header.length; c++) {
          var key = header[c];
          if (!key) continue;
          if (c === dateIdx) {
            obj[key] = formatYYMMDD(newDate);
          } else if (c === workerIdx) {
            obj[key] = newWorker;
          } else if (c === resultIdx) {
            obj[key] = newResult;
          } else if (c === noteIdx) {
            obj[key] = newNote;
          } else {
            obj[key] = row[c];
          }
        }
        updatedRowObj = obj;
        found = true;
        break;
      }
    }
  }
  if (!found) {
    return jsonResponse({ success: false, message: "수정할 방문내역을 찾을 수 없습니다." });
  }
  var cardUpdate = updateCardRecentVisit(areaId, cardNumber);
  var cardsValues = getSheet("구역카드").getDataRange().getValues();
  var complete = checkAreaComplete(areaId, cardsValues);
  if (complete && cardUpdate.recentVisitDate) {
    updateCompletion(areaId, cardUpdate.recentVisitDate, data.leaderName || "");
  }
  return jsonResponse({ success: true, visit: updatedRowObj, cardUpdate: cardUpdate, complete: complete });
}
