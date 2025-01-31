function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Complete')
    .addItem('set', 'tabcolor')
    //.addItem('reset', 'reset')
    
    .addToUi();}

function calendarToSheet(){
  var icalUrl_lmshin = '[사용자 icalUrl]';
  var lmshin = "이명신";
  var icalUrl_lurker132 = '[사용자 icalUrl]';
  var lurker132 = "신후철";
  var icalUrl_oper89054 = '[사용자 icalUrl]';
  var oper89054 = "신승섭";
  var icalUrl_arestarvip12 = '[사용자 icalUrl]';
  var arestarvip12 = "조성별";
  var icalUrl_djjozza77 = '[사용자 icalUrl]';
  var djjozza77 = "조규성";

  

  importICalToSheet(icalUrl_lmshin, lmshin);
  importICalToSheet(icalUrl_lurker132, lurker132);
  importICalToSheet(icalUrl_oper89054, oper89054);
  importICalToSheet(icalUrl_arestarvip12, arestarvip12);
  importICalToSheet(icalUrl_djjozza77, djjozza77);
}

function importICalToSheet(Calender_Url,sheetName) {
  // 캘린더에서 데이터를 추출하는 함수 (UrlFetchApp 사용)
  var icalUrl = Calender_Url;
  var response;
  try {
    response = UrlFetchApp.fetch(icalUrl, {muteHttpExceptions: true});
  } catch (e) {
    Logger.log('Error fetching iCal data: ' + e.message);
    return;
  }
  
  var icalData = response.getContentText();
  
  // iCal 데이터를 파싱하는 함수 호출
  var events = parseICalData(icalData);

  // 원하는 sheet 선택
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sheet;
  for (var i = 0; i < sheets.length; i++) {
    sheet = sheets[i];
    if (sheet.getName() == sheetName) {
       break;
    }
  }

  var today = new Date();
  var todayYMD = formatDateToYMD(today)
  
  var twoWeeksLater = new Date(today.getTime() + 14 * 24 * 60 * 60 * 1000); // 2주 후 날짜
  var twoWeeksLaterYMD = formatDateToYMD(twoWeeksLater)

  // 이벤트 데이터를 스프레드시트에 입력
  var rowIndex = 0; // 데이터를 작성할 시작 행 번호
  for (var i = 0; i < events.length; i++) {
    var event = events[i];

    var eventDate = new Date(event.start);
    var eventDateYMD = formatDateToYMD(eventDate);

    //날짜가 현재일 기준으로 2주 안의 데이터일 경우에만 실행
    if( todayYMD <= eventDateYMD && eventDateYMD <= twoWeeksLaterYMD){
      
      var date = event.start
      var keyword = extractData(event.summary)
      var category = keyword.category; //구분
      var company = keyword.company; //고객사명
      var detail = keyword.detail; //업무내용
      var startTime = formatDateTime(new Date(event.start)); //시작시간
      var endTime = formatDateTime(new Date(event.end)); //종료시간

      //시간을 설정하기 않은 경우에 대한 예외처리
      if(startTime==endTime){
        startTime = "9:00"
        endTime = "18:00"
      }

      //sheet에 데이터 입력하기
      rowIndex = findRowByDate(sheet, date);

      if(rowIndex==-1){
        console.log("sheet 상에 data를 집어넣을 양식이 존재하지 않습니다.")
      }
      else if (sheet.getRange(rowIndex, 4).getValue() !== "") {
        continue; // 데이터가 이미 있으면 다음 이벤트로 건너뜀
      }
      else{
        sheet.getRange(rowIndex, 2).setValue(company);
        sheet.getRange(rowIndex, 4).setValue(detail);
        sheet.getRange(rowIndex, 6).setValue(category);
        sheet.getRange(rowIndex, 7).setValue(startTime);
        sheet.getRange(rowIndex, 8).setValue(endTime);

        // 다음 행으로 이동
        rowIndex++;
      }

      
    }
  }
}


// iCal 데이터를 파싱하는 함수
function parseICalData(data) {
  var events = [];
  var currentEvent = null;

  var lines = data.split(/\r?\n/);
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (line.startsWith('BEGIN:VEVENT')) {
      currentEvent = {};
    } 
    else if (line.startsWith('END:VEVENT')) {
      events.push(currentEvent);
      currentEvent = null;
    } 
    else if (currentEvent) {
      var parts = line.split(':');
      var key = parts[0].trim();
      var value = parts.slice(1).join(':').trim();

      if (key === 'SUMMARY') {
        currentEvent.summary = value;
      } else if (key === 'DTSTART' || key === 'DTSTART;VALUE=DATE') {
        currentEvent.start = parseICalDate(value);
      } else if(key === 'DTEND' || key === 'DTEND;VALUE=DATE'){
        currentEvent.end = parseICalDate(value);
      }else if (key === 'DESCRIPTION') {
        currentEvent.description = value;
      }
    }
  }
  return events;
}

// iCal 날짜 문자열을 파싱하는 함수
function parseICalDate(value) {
  if (value.length === 8) { // YYYYMMDD 형식
    return value.replace(/(\d{4})(\d{2})(\d{2})/, '$1-$2-$3');
  } else if (value.length === 16) { // YYYYMMDDTHHmmssZ 형식
    return value.replace(/(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z/, '$1-$2-$3T$4:$5:$6Z');
  }
  return value; // 예상하지 못한 형식의 경우 원래 값 반환
}


// 날짜와 시간을 HH:mm 하는 함수,
function formatDateTime(date) {
  var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'HH:mm');
  return formattedDate;
}

function formatDateToYMD(date) {
    // 'yyyyMMdd' 형식으로 변환
    var year = date.getFullYear();
    var month = date.getMonth() + 1; // 0부터 시작하므로 +1
    var day = date.getDate();

    // 숫자로 변환 (예: 20241223)
    return year * 10000 + month * 100 + day;
}



//항목을 구분하는 함수
function classifyEvent(eventSummary) {
  const keywords = ["휴가", "설치", "점검", "기술","교육","기타","장애"];

  for (const keyword of keywords) {
    if (eventSummary.includes(keyword)) {
      return keyword;
    }
  }

  return "구분외"; // 위에 해당하는 키워드가 없을 경우
}

// 정규식으로 문자열을 구분하는 함수
function extractData(eventSummary) {
  const regex = /^\[(.*?)\]\s*([^ ]*)\s*(.*)$/;
  const match = eventSummary.match(regex);
  
  //연차휴가 예외처리
  if(match==null){
    console.log("형식에 맞지않는 데이터 값이 입력 되었습니다.", eventSummary)
  }
  else if(match[1]=="휴가"){
    match[2] = ""
    match[3] = "연차휴가"
  }
  else if(match[1]=="alt_off"){
    match[1] = "휴가"
    match[2] = ""
    match[3] = "alt_off"
  }
  else{}

  const keywords = ["설치", "점검", "기술","기타","장애","휴가","alt_off"];

  if(match==null){
    console.log("형식에 맞지않는 데이터 값이 입력 되었습니다.")
  }
  else if(match[1] === keywords[0]){
    match[1] = "설치업무";
  }
  else if(match[1] === keywords[1]){
    match[1] = "정기점검";
  }
  else if(match[1] === keywords[2]){
    match[1] = "기술지원";
  }
  else if(match[1] === keywords[3]){
    match[1] = "기타업무";
  }
  else if(match[1] === keywords[4]){
    match[1] = "장애지원";
  }
  else if(match[1]==keywords[5]){
    match[2] = ""
    match[3] = "연차휴가"
  }
  else if(match[1]==keywords[6]){
    match[1] = "휴가"
    match[2] = ""
    match[3] = "alt_off"
  }


  if (match) {
      return {
          category: match[1],
          company: match[2],
          detail: match[3]
      };
  } else {
      return {
          category: "기타",
          company: "미확인",
          detail: eventSummary
      };
  }
}

function findRowByDate(sheet, searchDate) {
  var lastRow = sheet.getLastRow();
  var dataRange = sheet.getRange(1, 1, lastRow, 1); // A열 전체 데이터 가져오기
  var values = dataRange.getValues();
  
  for (var i = 0; i < values.length; i++) {

    var V = new Date(values[i][0]);
    var SD = new Date(searchDate);

    flag = isSameDate(V,SD)

    if (flag) {
      return i + 1; // 행 번호 반환
    }
  }
  return -1; // 찾지 못함
}

function isSameDate(date1, date2){
  return(
    date1.getFullYear() === date2.getFullYear() &&
    date1.getMonth() === date2.getMonth() &&
    date1.getDate() === date2.getDate()
  );
}
