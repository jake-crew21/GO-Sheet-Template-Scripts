const importSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pulled Check In Data");

//data[][List Id, Staff Id, Staff Name, Channel, Date(DateTime), Checkin Time(DateTime)]

function writeUpdateCI() {
  var data = getTodayCheckInData();
  var lc = importSheet.getLastColumn();
  var lr = importSheet.getLastRow();
  importSheet.getRange(1,1,lr,lc).clearContent();
  importSheet.getRange(1,1,data.length,data[0].length).setValues(data);
}

function getTodayCheckInData() {
  const checkInSheet = checkInSS.getSheetByName("Roll Data");
  var lr = checkInSheet.getLastRow();
  var lc = checkInSheet.getLastColumn();
  var data = checkInSheet.getRange(1,1,lr,lc).getValues();
  var today = new Date();
  today.setHours(0,0,0,0);
  var tdIndx = data.findIndex((x) => +x[3] == +today);
  
  var todaysCourses = data.slice(tdIndx,data.length);
  todaysCourses.unshift(data[0])
  // console.log(+data[lr-1][3]==+today);
  // console.log(todaysCourses.length);
  return todaysCourses;
}

//List ID, Workflow Timestamp,	Value ID,	UTC Date,	State,	Venue,	Venue ID,	Course,	Course ID,	Class,	Staff,	Pending Staff,	Arrived Staff,	Channel,	Completed,	On Site,	On Site Since

/**
 *
 *
 * @return {JSON.stringify(displayData)} 
 */
function getLiveData(){
  var lr = importSheet.getLastRow();
  var lc = importSheet.getLastColumn();

  if (lr === 0 || lc === 0) {return [['No Data Found']];} // Fallback in case the sheet is empty

  var data = importSheet.getRange(1,1,lr,lc).getValues();
  //data[?][4,5,7,10,11,13,14,15]  State, Venue, Course, All Staff, Pending Staff, Channel, Completed (Roll)
  var displayData = [];
  data.forEach(e => {
    var temp = [];
    temp.push(e[4],e[5],e[7],e[10],e[11],e[13],e[14],e[15]);
    displayData.push(temp);
  });
  return JSON.stringify(displayData);
}

function getCheckInData(){
  const checkInSheet = checkInSS.getSheetByName("Roll Data");
  var lr = checkInSheet.getLastRow();
  var lc = checkInSheet.getLastColumn();

  if (lr === 0 || lc === 0) {
    return [['No Data Found']]; // Fallback in case the sheet is empty
  }

  var data = checkInSheet.getRange(1,1,lr,lc).getValues();
  // console.log(data)
  return data;
}

// Function to open the HTML dialog
function openDialog() {
  writeUpdateCI();
  const html = HtmlService.createHtmlOutputFromFile('Index')
                .setWidth(1200)
                .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Check-ins');
}

function testCall() {return "Success"}