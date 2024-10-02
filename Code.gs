const BASE_PATH = "https://finance.yahoo.com/calendar/earnings";
const MAX_TRIES = 5; 
const MILLIS_PER_DAY = 1000*60*60*24;  
const DATE_RANGE = 7; 
const EARNINGS_SHEET = "Earnings Data"
const EARNINGS_RANGE = "A:H"

function getEarningsForDate(date) {
  //Variables for outer loop and URL params
  var offset = 0;
  var count = 0; 
  var html = ""; 
  var response = "";
  var pages = 1;
  var hasResults = true;
  var x = 0; 
  //Variable to store a full days earnings records
  var output = [];
  while(x < pages+1 && hasResults){
    offset = x*100; 
    var URL = BASE_PATH+"?day="+date+"&offset="+offset+"&size=100";
    var success = false
    var tries = 0; 
    //Make sure the URL is received (network errors etc.)
    while (!success && tries < MAX_TRIES){
      try{
        response = UrlFetchApp.fetch(URL);
        Utilities.sleep(500);
        html = response.getContentText()
        success = true
        console.log(URL)
      }
      catch(e){
        console.error("Error fetching URL: "+URL)
        tries++;
        Utilities.sleep(2000);
      }
    }
    //Check number records first time
    if(x == 0){
      results = html.match(/of ([0-9]+) results/g)
      if(!results){
        console.log("No earnings on date: "+date)
        hasResults = false;
        return null;
      }
      records = parseInt(results[0].match(/[0-9]+/g))  
      console.log(records+" records on "+date)
      //Calculate pages based on 100 records per page
      pages = Math.floor(records / 100) 
    }
    
    //Get the earnings table
    table = html.match(/<tbody>.*<\/tbody>/g)
    if(table == null){
      if(output.length < 1){
        return null;
      }
      else{
        return output; 
      }
    }
    table = table[0]
    //List of elements within the table
    values = table.split("<td")
    row = [new Date(date)]//.split("-").reverse().join("/"))]
    //For each td element:
    for(let i = 0; i  < values.length; i++){
      //Skip first blank row, add the text from the element (narrowing down using split)
      if(i>0) row.push(values[i].split("</td>")[0].split("</span>")[0].split("</a><div")[0].split(">").slice(-1)[0])
      //7 columns, start new after each row
      if(i % 7 == 0 && i>1){
        //console.log(row)
        count++;
        output.push(row);
        row = [new Date(date)]//.split("-").reverse().join("/"))]
      }
    }
    //Increment variable for offset to fetch next page
    x++;
  }
  //console.log(count);
  return output;
}

function isoFormat(dateString){
  date = new Date(dateString);
  return Utilities.formatDate(date, "GMT+4:00", "yyyy-MM-dd"); 
}

function sheetDateFormat(dateString){
  date = new Date(dateString);
  return Utilities.formatDate(date, "GMT+4:00", "dd/MM/yyyy"); 
}

function findRowForDate2(date){
  ss = SpreadsheetApp.getActive()
  range = ss.getSheetByName(EARNINGS_SHEET).getRange(EARNINGS_RANGE); 
  vals = range.getValues(); 
  for(let i = vals.length-1; i >= 1; i--){
    if(date != sheetDateFormat(vals[i-1][0] && date == sheetDateFormat(vals[i][0]))){
      console.log(date)
      console.log(sheetDateFormat(vals[i][0]))
      //console.log(i+1)
      return i+1; 
    }
    if(new Date(isoFormat(date)).getTime() < new Date(isoFormat(vals[i][0]))){
      return null; 
    }
  }
  return null; 
}


function findRowForDate(date){
  ss = SpreadsheetApp.getActive()
  range = ss.getSheetByName(EARNINGS_SHEET).getRange(EARNINGS_RANGE); 
  vals = range.getValues(); 
  for(let i = 30000; i < vals.length; i++){
    if(date == sheetDateFormat(vals[i][0])){
      console.log(date)
      console.log(sheetDateFormat(vals[i][0]))
      //console.log(i+1)
      return i+1; 
    }
  }
  return null; 
}

function getLastRow(sheet, column){
  range = sheet.getRange(column+":"+column);
  vals = range.getValues()
  for(let i = vals.length-1; i >= 1; i--){
    if(vals[i-1] != "" && vals[i] == ""){
      return i; 
    }
  }
  return vals.length;
}

function main(){
  sheet = SpreadsheetApp.getActive().getSheetByName(EARNINGS_SHEET);
  lastRow = getLastRow(sheet, "A");
  maxDateInSheet = sheet.getRange("A"+lastRow).getValue()
  today = new Date()
  startDate = new Date(today.getTime() - (MILLIS_PER_DAY*3))
  endDate = new Date(today.getTime() + (MILLIS_PER_DAY*DATE_RANGE))
  console.log(isoFormat(startDate) + " " + sheetDateFormat(startDate))
  console.log(new Date(isoFormat(maxDateInSheet)).getTime() + " " + new Date(isoFormat(startDate)).getTime())
  foundClearPoint = false;
  var d = 0;  
  while (d <= DATE_RANGE*2 && !foundClearPoint && (new Date(isoFormat(maxDateInSheet)).getTime() >= new Date(isoFormat(startDate)).getTime())){
    clearRow = findRowForDate(sheetDateFormat(startDate.getTime()+(d*MILLIS_PER_DAY)))
    if(clearRow != null ){
      console.log("Clear range: "+EARNINGS_RANGE.replace(":",clearRow+":")+lastRow)
      sheet.getRange(EARNINGS_RANGE.replace(":",clearRow+":")+lastRow).clearContent()
      foundClearPoint = true; 
    }
    d++;
  }
  for (let i = 0; i <= DATE_RANGE+3;i++){
    day = isoFormat(startDate.getTime()+(i*MILLIS_PER_DAY));
    daysEarnings = getEarningsForDate(day);
    if(daysEarnings != null){
      lastRow = getLastRow(sheet,"A")+1
      sheet.getRange(EARNINGS_RANGE.replace(":",(lastRow)+":")+(lastRow+daysEarnings.length-1)).setValues(daysEarnings); 
    }
  }
}


function addToCalendar(){
  const TRADING_HOLIDAYS = [
    "2024-01-01",
    "2024-01-15",
    "2024-02-19",
    "2024-03-29",
    "2024-05-27",
    "2024-06-19",
    "2024-07-04",
    "2024-09-02",
    "2024-11-28",
    "2024-12-25",
    "2025-01-01",
    "2025-01-20",
    "2025-02-17",
    "2025-04-18",
    "2025-05-26",
    "2025-06-19",
    "2025-07-04",
    "2025-09-01",
    "2025-11-27",
    "2025-12-25",
    "2026-01-01",
    "2026-01-19",
    "2026-02-16",
    "2026-04-03",
    "2026-05-25",
    "2026-06-19",
    "2026-07-03",
    "2026-09-07",
    "2026-11-26",
    "2026-12-25"
  ]
  var calendar = CalendarApp.getCalendarById("756bc8006f9a3021157df87a3b4fce1d10fcbdb3c23281a5a3d9d382029fd85a@group.calendar.google.com");
  daysAhead = SpreadsheetApp.getActive().getSheetByName("Backtests").getRange("B2").getValue(); 
  upcomingEarnings = SpreadsheetApp.getActive().getSheetByName("Upcoming Earnings").getDataRange().getValues();
  console.log(daysAhead) 

  today = new(Date);
  todaysPlays = []
  for(let i = 1; i < upcomingEarnings.length; i++){
    if(upcomingEarnings[i][0] != ''){
      //console.log(upcomingEarnings[i])
      orderDate = new Date(new Date(upcomingEarnings[i][0]).getTime() + MILLIS_PER_DAY*daysAhead)
      if(TRADING_HOLIDAYS.includes(isoFormat(orderDate))){
        orderDate = new Date(orderDate).getTime() + MILLIS_PER_DAY;
      }
      if(Utilities.formatDate(orderDate,"GMT","u") > 5){
        orderDate = new Date(orderDate).getTime() + MILLIS_PER_DAY * (parseInt(Utilities.formatDate(orderDate,"GMT","u"))-5);
      }
      orderDate = isoFormat(orderDate)
      if(orderDate === isoFormat(today)){
        console.log(upcomingEarnings[i])
        ticker = [upcomingEarnings[i][1]];
        name = new String([upcomingEarnings[i][2]]).replace("&amp;","&");
        surprise = [upcomingEarnings[i][7]];
        pctChange = [upcomingEarnings[i][13]];
        todaysPlays.push([orderDate,ticker,name,surprise,pctChange]);
        
      }
    }
  }

  if(todaysPlays.length == 0){
    return;
  }

  console.log(todaysPlays)

  subject = sheetDateFormat(today) + " Earnings plays: "
  desc = "Companies to trade today based on earnings:\n\n"
  for(let i = 0; i < todaysPlays.length; i++){
    subject += todaysPlays[i][1]
    if(todaysPlays.length-1 != i){
      subject += " | ";
    }
    desc += todaysPlays[i][1] + ": " + todaysPlays[i][2];
    surprise = parseFloat(todaysPlays[i][3])
    if(isNaN(surprise) || surprise == ''){
      desc += " Earnings results not found. "
    }
    else if(surprise > 0){
      desc += " Beat earnings by " + surprise + "%. "
    }
    else if(surprise < 0){
      desc += " Missed earnings by " + surprise + "%. "
    }
    else if(surprise == 0){
      desc += " Matched expecations. "
    }

    pctMove = parseFloat(todaysPlays[i][4])
    if(isNaN(pctMove) || todaysPlays[i][4] == ''){
      desc += " Could not find price data. "
    }
    else if(pctMove > 0){
      desc += " This caused the stock to jump " + (pctMove*100).toFixed(2) + "% on earnings day."
    }
    else if(pctMove < 0){
      desc += " This caused the stock to fall " + (pctMove*100).toFixed(2) + "% on earnings day."
    }
    else if(pctMove == 0){
      desc += " The stock ended flat on the day. "
    }

    desc += "\n\n";

    

  }
  console.log(isoFormat(today))
  yyyy = isoFormat(today).split("-")[0]
  MM = isoFormat(today).split("-")[1] - 1
  dd = isoFormat(today).split("-")[2]
  hh = "14"
  mm = "30"
  eventStartDate = new Date(yyyy,MM,dd,hh,mm);
  eventEndDate = new Date(yyyy,MM,dd,hh,new String(parseInt(mm)+30))
  console.log(eventStartDate)
  console.log(eventEndDate)
  console.log(subject)
  console.log(desc)

  calendar.createEvent(subject,eventStartDate,eventEndDate,{description:desc});



}



