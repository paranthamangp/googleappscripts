
/**
 * @ Copy daily returns value and paste it to below  to store historical data
 *  Sheet: Daily_Log
 */

function copyHistoricalData() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:P1').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Daily_Log'), true);
  spreadsheet.getRange('21:21').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('A21').activate();
  spreadsheet.getRange('A3:J3').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

}

/**
 * @get data from google sheet and store to array
 *  Sheet: Daily_Log
 */
function getData() {
  var okPortfolioLogs = [];
  var dailyLog ={};

  //get OK portfolio return values using named range
  var portfolioReturn = SpreadsheetApp.getActive().getSheetByName("Daily_Log").getRange("OK_Portfolio_Returns").getValues();
  portfolioReturn.shift(); //remove headers
  portfolioReturn.forEach(function(value) {
    dailyLog.date = value[0];
    dailyLog.totalCost = value[2];
    dailyLog.totalCurrentValue = value[3];
    dailyLog.overallProfitRs = value[4];
    dailyLog.overallProfitPercent = value[5];
    dailyLog.TodayProfitRs = value[6];
    dailyLog.TodayProfitPercent = value[7];
  })

  //get Index return values using named range
  var indexReturn = SpreadsheetApp.getActive().getSheetByName("Daily_Log").getRange("Index_Returns").getValues();
  indexReturn.shift(); //remove headers
  indexReturn.forEach(function(value) {
    dailyLog.market = value[0];
    dailyLog.nifty50 = value[1];
    dailyLog.niftynext50 = value[2];
    dailyLog.midcap = value[3];
    dailyLog.smallcap = value[4];
    dailyLog.nasdaq100 = value[5];
    dailyLog.gold = value[6];
    dailyLog.silver = value[7];
    dailyLog.vix = value[8];
  })

  okPortfolioLogs.push(dailyLog);
  //Logger.log(okPortfolioLogs)
  return okPortfolioLogs;
}

/**
 * @form email template
 *  Sheet: Daily_Log
 */

function getEmailText(dailyLogData) {
  
  var text = "";

  dailyLogData.forEach(function(dailyLog) {

      //if market is open send the portfolio returns else sen
      if(dailyLog.market == "Open"){
      
        //Form dynamic messages based on todays return
        var dynamicMessage = "";
        if( dailyLog.TodayProfitRs > 0 ){
          dynamicMessage = dynamicMessage + " increased by "
        } else {
          dynamicMessage = dynamicMessage + " decreased by "
        }

        text = text +
              "Todays OK portfolio" +  dynamicMessage + (Math.round(dailyLog.TodayProfitPercent *10000)/100)+ '%'+ "\n"+
              "------------------------------------------------------------\n" +
              "OK Portfolio Returns	: " +"\n" +
              "Total Portfolio Cost (Rs.)	: " + "Rs. "+ Math.round(dailyLog.totalCost) + "\n" + 
              "Total Portfolio Value (Rs.) : " + "Rs. "+ Math.round(dailyLog.totalCurrentValue) + "\n" + 
              "Overall Profit / Loss (Rs.) : " + "Rs. "+ Math.round(dailyLog.overallProfitRs) + "\n" + 
              "Overall Profit / Loss (%) : " + (Math.round(dailyLog.overallProfitPercent*10000)/100)+ '%' + "\n" + 
              "Today Profit / Loss (Rs.) : " + "Rs. "+ Math.round(dailyLog.TodayProfitRs) + "\n" + 
              "Today Profit / Loss  (%) : " + (Math.round(dailyLog.TodayProfitPercent *10000)/100)+ '%'+ "\n" + 
              "------------------------------------------------------------\n"+
              "Nifty 50 Index: "  + (Math.round(dailyLog.nifty50 *10000)/100)+ '%'+ "\n" + 
              "Nifty Next 50 Index: " + (Math.round(dailyLog.niftynext50 *10000)/100)+ '%'+ "\n" + 
              "Midcap Index : " + (Math.round(dailyLog.midcap *10000)/100)+ '%'+ "\n" + 
              "Smallcap Index : " + (Math.round(dailyLog.smallcap *10000)/100)+ '%'+ "\n" + 
              "NASDAQ 100 Index : " + (Math.round(dailyLog.nasdaq100 *10000)/100)+ '%'+ "\n" + 
              "Gold : " + (Math.round(dailyLog.gold *10000)/100)+ '%'+ "\n" + 
              "Silver : " + (Math.round(dailyLog.silver *10000)/100)+ '%'+ "\n" + 
              "India VIX : " + (Math.round(dailyLog.vix *10000)/100)+ '%'+ "\n" + 
              "------------------------------------------------------------\n"+
            
              "Automatic script execution at "+ dailyLog.date +"\n";
      }else{

        //else if market is closed today
        text = text + "Market is closed today " +"\n" +
        "------------------------------------------------------------\n"+
        "Automatic script execution at "+ dailyLog.date +"\n";
      }

      });

  Logger.log(text)
  return text;
}

/**
 * @trigger email request - Scheduled to trigger at 4PM IST daily
 */

function sendEmail() {
  
  //copy current day data to historical data table
  copyHistoricalData();
  var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy")
  var dailyLogsData = getData();
  var body = getEmailText(dailyLogsData);

  //hide unwanted sheet so that it doesnt gets stored in attachment
  //SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Read This First").hideSheet();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Watchlist").hideSheet();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary_OSV").hideSheet();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Notes").hideSheet();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions_OSV").hideSheet();
  //SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Daily_Log").hideSheet();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ChartData-DONTEDIT").hideSheet();


  var recipient = "********@gmail.com";
  var subject = "Ok Portfolio Returns for " + date;
  var options = {
    cc: "**************@gmail.com",
    //bcc: "bcc@example.com",
    //replyTo: "help@example.com"
    attachments: [SpreadsheetApp.getActiveSpreadsheet().getAs(MimeType.PDF).setName("Portfolio Summary "+ date +" .pdf")]
  }
  MailApp.sendEmail(recipient, subject, body, options);


  //activate all hidden sheets back
  //SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Read This First").showSheet();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Watchlist").showSheet();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary_OSV").showSheet();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Notes").showSheet();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions_OSV").showSheet();
  //SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Daily_Log").showSheet();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ChartData-DONTEDIT").showSheet();
}


