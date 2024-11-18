const activeSheet = liveDataSS.getSheetByName("Active_Courses");
const allData = liveDataSS.getSheetByName("LiveData");
const dowSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

const term = "After-School Term 4 2024";//<--- this much be changed each term
//Monday = 1, Tuesday = 2, Wednesday = 3, Thursday = 4, Friday = 5
//Week Cell Ranges: wk1 = (I)J:N, wk2 = (O)P:T, wk3 = (U)V:Z, wk4 = (AA)AB:AF, wk5 = (AG)AH:AL, wk6 = (AN)AN:AR, wk7 = (AS)AT:AX, wk8 = (AY)AZ:BD, wk9 = (BE)BF:BJ, wk10 = (BK)BL:BP, wk11 = (BQ)BR:BV
//Week 1 Monday = 15/07/2024
const wkCol = ["I","O","U","AA","AG","AM","AS","AY","BE","BK","BQ","BW","CC"];
var wkColNum = [];

/**
 * takes week 1 date, and sets the date for the following weeks
 *    this has to be set manually
 */
function setWeekDate(){
  var wkOneCol = 10;
  var cs = 6;
  var wkCount = 1;
  const wkOne = dowSheet.getRange(2,wkOneCol).getValue();
  console.log(wkOne);
  var lc = dowSheet.getLastColumn();
  while(((cs*wkCount)+9)<lc){
    var dateReset = dowSheet.getRange(2,wkOneCol).getValue();
    var days = wkCount * 7;
    var newDate = add_weeks(dateReset, wkCount)
    var col = (cs*wkCount) + wkOneCol;
    dowSheet.getRange(2,col).setValue(newDate);
    wkCount++;
  }
}

function add_weeks(dt, n) {
  return new Date(dt.setDate(dt.getDate() + (n * 7)));
}

/**
 * removes cancelled courses
 */
function cancelled(){
  var courses = allData.getRange("A2:E").getValues();
  var dowAllRows = dowSheet.getRange("A4:A").getValues();
  var dow = dowAllRows.filter(function (el) {
    return el != null;
  });
  var cancelles = courses.filter(function (c) {
    return c[4] == "cancelled" || c[4] == "deleted" || c[4] == "pending";
  })
  for(i=dow.length; i>=0; i--){
    cancelles.find(function (x) {
      if(x[0]==dow[i]){
        console.log(dow[i]+" is cancelled");
        dowSheet.deleteRow(i+4);
      }
    })
  }
}

/**
 * Resizes column on each week
 */
function setColSize(){
  const partition = 20; const staff = 80; const roll = 40;  const rating = 85;  const comments = 600;
  var lc = dowSheet.getLastColumn();
  let state = 1;
  for(c=9; c<=lc; c++){
    switch (state){
      case 1:
        dowSheet.setColumnWidth(c,20);
        state++;
        break;
      case 2:
        dowSheet.setColumnWidth(c,80);
        state++;
        break;
      case 3:
        dowSheet.setColumnWidth(c,80);
        state++;
        break;
      case 4:
        dowSheet.setColumnWidth(c,40);
        state++;
        break;
      case 5:
        dowSheet.setColumnWidth(c,85);
        state++;
        break;
      case 6:
        dowSheet.setColumnWidth(c,600);
        state=1;
        break;
      default:
        break;
    }
  }
}

/**
 * Get column number based on Letter value
 */
function findColumnNumber() {
  var columnNumbers = [];
  wkCol.forEach(function(c) {
    let x = dowSheet.getRange(c+"1").getColumn();
    columnNumbers.push(x);
  })
  return columnNumbers;
}

/**
 * Convert int to column letter
 */
function columnToLetter(column){
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Get all active couses
 * Seperate course by day of week they run on
 * Gets day of week from sheet name and writes course for that day
 * Course line: lola course link, lola venue link, state, course type, start time, finish time
 */
function coursesByDay() {
  var api = liveDataSS.getSheetByName("Active_Courses");
  var data = api.getRange("A:P").getValues();
  // console.log(data.length);
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var currentSheet = ss.getSheetName();
  var courses = data.filter((c) => {
    return c[7] == term;
  })
  var monday = courses.filter((c) => {
    if(c[14].getDay() == 1) {return c;}
  })
  var tuesday = courses.filter((c) => {
    if(c[14].getDay() == 2) {return c;}
  })
  var wednesday = courses.filter((c) => {
    if(c[14].getDay() == 3) {return c;}
  })
  var thursday = courses.filter((c) => {
    if(c[14].getDay() == 4) {return c;}
  })
  var friday = courses.filter((c) => {
    if(c[14].getDay() == 5) {return c;}
  })
  switch (currentSheet) {
    case 'Monday':
      settingVal(currentSheet, monday);
      break;
    case 'Tuesday':
      settingVal(currentSheet, tuesday);
      break;
    case 'Wednesday':
      settingVal(currentSheet, wednesday);
      break;
    case 'Thursday':
      settingVal(currentSheet, thursday);
      break;
    case 'Friday':
      settingVal(currentSheet, friday);
      break;
    default:
      break;
  }
}

/**
 * Write links to course & venue, state, course type, start time, end time
 */
function settingVal(weekDay, arr) {
  wkColNum = findColumnNumber();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(weekDay);
  var lc = sheet.getLastColumn();
  var wkOneDate = sheet.getRange("J2").getValue();
  var courses = [];
  var links = [];
  // sheet.getRange(4,9,arr.length,(lc-8)).setBackground('#666666')
  arr.forEach(function(obj, i) {
    var courseId = linkBuilder(obj[0],courseLInk,obj[0]);
    var venueName = linkBuilder(obj[3],venueLink,obj[2]);
    links.push([courseId,venueName]);
    var name = courseName(obj[1]);
    var sTime = startTime(obj[0],obj[10]);
    var eTime = endTime(obj[0],obj[10]);
    courses.push([obj[10],name,sTime,eTime]);
  })
  sheet.getRange(4,1,links.length,links[0].length).setRichTextValues(links);
  sheet.getRange(4,3,courses.length,courses[0].length).setValues(courses);
  // console.log(`${links[7]}: ${courses[7][0]} ${courses[7][2]}`);
  // console.log(links.length);
  // console.log(courses.length);
}

/**
 * Write line with set data for each of it's column
 */
function startToEndDate(){
  var api = liveDataSS.getSheetByName("Active_Courses");
  var data = api.getRange("A:P").getValues();
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var currentSheet = ss.getSheetName();
  wkColNum = findColumnNumber();

  var courses = data.filter((c) => {
    return c[7] == term;
  })
  var monday = courses.filter((c) => {
    if(c[14].getDay() == 1) {return c;}
  })
  var tuesday = courses.filter((c) => {
    if(c[14].getDay() == 2) {return c;}
  })
  var wednesday = courses.filter((c) => {
    if(c[14].getDay() == 3) {return c;}
  })
  var thursday = courses.filter((c) => {
    if(c[14].getDay() == 4) {return c;}
  })
  var friday = courses.filter((c) => {
    if(c[14].getDay() == 5) {return c;}
  })
  switch (currentSheet) {
    case 'Monday':
      settingStartToEnd(monday);
      break;
    case 'Tuesday':
      settingStartToEnd(tuesday);
      break;
    case 'Wednesday':
      settingStartToEnd(wednesday);
      break;
    case 'Thursday':
      settingStartToEnd(thursday);
      break;
    case 'Friday':
      settingStartToEnd(friday);
      break;
    default:
      break;
  }
}

function settingStartToEnd(arr){
  var termVals = [];
  wkColNum = findColumnNumber();
  var wkDates = [];
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = sheet.getLastRow();
  var lc = sheet.getLastColumn();
  var daysCourses = sheet.getRange(4,1,(lr-3)).getValues();
  wkColNum.forEach(function(col){
    var tempDate = sheet.getRange(2,col+1).getValue();
    // tempDate.setHours(0,0,0,0);
    wkDates.push(tempDate);
  })
  // console.log(wkDates);
  daysCourses.forEach(function(d){
    var courseWks = [];
    arr.find(function(a){
      if(d[0]==a[0]){
        wkDates.forEach(function(w){
          if(w>=a[14] && w<=a[15]){
            if(a[12]>12){
              courseWks.push("","FALSE","FALSE","FALSE","","");
            }else{courseWks.push("","FALSE","","FALSE","","");}
          }else{courseWks.push("","","","","","");}
        })
      }
    })
    termVals.push(courseWks);
  })
  console.log(termVals[0].length);
  sheet.getRange(4,9,termVals.length,termVals[0].length).setValues(termVals);
}

function insertCheckboxes(){
  wkColNum = findColumnNumber();
  var lr = dowSheet.getLastRow();
  wkColNum.forEach(function(col,week){
    var wkRows = dowSheet.getRange(4,col+1,lr-3,2).getValues();
    var rOne = [], rTwo = [], tempOne = [], tempTwo = [];
    // var rTwo = []; var tempOne = []; var tempTwo = [];
    wkRows.forEach(function(c, i){
      if(c[0] === "" && tempOne.length > 0){
        rOne.push(tempOne);
        tempOne = [];
      } else if (c[0] !== ""){
        tempOne.push(i+4);
      }
      if(c[1] === "" && tempTwo.length > 0){
        rTwo.push(tempTwo);
        tempTwo = [];
      } else if (c[1] !== ""){
        tempTwo.push(i+4);
      }
    })
    if(tempOne.length>0){rOne.push(tempOne);}
    if(tempTwo.length>0){rTwo.push(tempTwo);}
    if(rOne.length>0){
      rOne.forEach(function(r){
        dowSheet.getRange(r[0],col+1,r.length).insertCheckboxes();
        dowSheet.getRange(r[0],col+3,r.length).insertCheckboxes();
        dowSheet.getRange(r[0],col+4,r.length).setDataValidation(ratingDropDown());
      })
    }
    if(rTwo.length>0){
      rTwo.forEach(function(r){
        dowSheet.getRange(r[0],col+2,r.length).insertCheckboxes();
      })
    }
  })
}

function updateStaff(){
  wkColNum = findColumnNumber();
  var lr = dowSheet.getLastRow();
  var daysCourses = dowSheet.getRange(4,1,lr-3).getValues();
  var data = activeSheet.getRange("A:P").getValues();
  wkColNum.forEach(function(col,week){
    var wkRows = dowSheet.getRange(4,col+1,lr-3,2).getValues();
    wkRows.forEach(function(c, i){
      var count = data.find(d => d[0]==daysCourses[i][0])
      if(count[12]>12 && c[0]!=="" && c[1]===""){
        dowSheet.getRange(i+4,col+2).insertCheckboxes();
      } else if (count[12]<=12 && c[0]!=="" && c[1]!=="") {
        dowSheet.getRange(i+4,col+2).removeCheckboxes();
      }
    })
  })
}

/**
 * set background for the week to white when first col of the week is not blank
 */
function bgConFormating(){
  wkColNum = findColumnNumber();
  var ranges = [];
  var rules = dowSheet.getConditionalFormatRules();
  wkColNum.forEach(function(col){
    var r = dowSheet.getRange(4,col+1,100,5);
    var colLetter = columnToLetter(col+1);
    var formula = "=NOT(ISBLANK($"+colLetter+"4))";
    ranges.push(r);
    const isNotBlank = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground("#ffffff")
    .setRanges([r])
    .build();
    rules.push(isNotBlank);
  })
  dowSheet.setConditionalFormatRules(rules);
}

/**
 * Creates the drop down menu for each week for runnning courses
 * Content: Good, Needs attention, Urgent
 */
function ratingDropDown(){
  const listContent = ["Good","Needs attention","Urgent"];
  const dropList = SpreadsheetApp.newDataValidation()
    .requireValueInList(listContent)
    .setAllowInvalid(false)
    .build();
  return dropList;
}

/**
 * Condition formating for all 'Rating' columns.
 * Changes cell colour coresponding to rating set
 * Good: green
 * Needs attention: yellow
 * Urgent: red
 */
function ratingConFortmat(){
  wkColNum = findColumnNumber();
  var ranges = [];
  wkColNum.forEach(function(col){
    var r = dowSheet.getRange(4,col+4,100);
    ranges.push(r);
  })
  const goodRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Good")
    .setBackground("green")
    .setRanges(ranges)
    .build();
  const naRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Needs attention")
    .setBackground("yellow")
    .setRanges(ranges)
    .build()
  const urgentRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Urgent")
    .setBackground("red")
    .setRanges(ranges)
    .build();
  var rules = dowSheet.getConditionalFormatRules();
  rules.push(goodRule,naRule,urgentRule);
}

/**
 * Adjust start and finish times for course in different states
 */
function timeAdj(state, time) {
  let options = {timeStyle: 'short', hour12: true};
  let [h,m,s] = time.split(':');
  let date = new Date();
  let res;
  switch (state) {
    case "SA":
      // options = {timeZone: 'Australia/Adelaide', timeStyle: 'short', hour12: true};
      date.setHours(h,m);
      date.setMinutes(date.getMinutes()+30);
      res = date.toLocaleTimeString("en-AU", options);
      return res;
    case "WA":
      // options = {timeZone: 'Australia/Perth', timeStyle: 'short', hour12: true};
      date.setHours(h,m);
      date.setHours(date.getHours()+2);
      res = date.toLocaleTimeString("en-AU", options);
      return res;
    case "QLD":
      // options = {timeZone: 'Australia/Brisbane', timeStyle: 'short', hour12: true};
      date.setHours(h,m);
      date.setHours(date.getHours()+1);
      res = date.toLocaleTimeString("en-AU", options);
      return res;
    default:
      date.setHours(h,m);
      res = date.toLocaleTimeString("en-AU", options);
      return res;
  }
}

function updateTime() {
  var lr = dowSheet.getLastRow();
  var times = dowSheet.getRange("C4:F"+lr).getValues();
  var newTiems = [];
  times.forEach(function(x){
    var sT, eT;
    sT = updateTimeAdj(x[0], x[2]);
    eT = updateTimeAdj(x[0], x[3]);
    newTiems.push([x[0],x[1],sT,eT]);
  })
  // console.log(newTiems);
  dowSheet.getRange(4,3,newTiems.length,newTiems[0].length).setValues(newTiems);
}
function updateTimeAdj(state, time) {
  let options = {timeStyle: 'short', hour12: true};
  let res;
  switch (state) {
    // case "QLD":
    //   time.setHours(time.getHours()+1);
    //   res = time.toLocaleTimeString("en-AU", options);
    //   return res;
    case "WA":
      time.setHours(time.getHours()+1);
      res = time.toLocaleTimeString("en-AU", options);
      return res;
    default:
      res = time.toLocaleTimeString("en-AU", options);
      return res;
  }
}

/**
 * Shorten course type name for readability
 */
function courseName(fullName) {
  if(fullName == "Code Camp After-School Coding"){return "Coding";}
  else if(fullName == "Little Coders After-School"){return "Little Coders";}
  else if(fullName == "Curious Minds by Code Camp"){return "Curious Minds";}
  else if(fullName == "Minecraft Engineers"){return "Minecraft Engineers";}
  else if(fullName == "Robotics After-School"){return "Robotics";}
  else if(fullName == "Animation After-School"){return "Animation";}
  else if(fullName == "Design After-School"){return "Design";}
  else if(fullName == "Code Camp After-School 3D"){return "3D";}
  else {return fullName;}
}

/**
 * Get all non hidden sheets of GO Sheet
 */
function ALLSHEETNAMES() {
  let ss = SpreadsheetApp.getActive();
  let sheets = ss.getSheets();
  let sheetNames = [];
  sheets.forEach(function (sheet) {
    if(!sheet.isSheetHidden()){sheetNames.push(sheet.getName());}
  });
  sheetNames.pop();
  return sheetNames;
}

/**
 * Create the link and text to be used as a 'Rich Text' data type
 */
function linkBuilder(txt, link, id) {
  var richValue = SpreadsheetApp.newRichTextValue()
    .setText(txt)
    .setLinkUrl(link+id)
    .build();
  return richValue;
}