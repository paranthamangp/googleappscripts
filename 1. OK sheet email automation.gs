/** 
  * OK Portfolio Sheet Automation
  *  Change log V1
  *    - 11/08/2024 - Base version of Ok Portfolio Sheet 
  *    - 17/08/2024 - Added other index comparison and NAV calculation
  *    - 26/12/2024 - Added XIRR Calculation
  *    - 27/12/2024 - SortOKIndexByDailyReturn() - Sort Ok Index sheet by daily returns
  *    - 09/01/2025 - copyATHData() - Copy ATH data value to ATH sheet
  *    - 09/01/2025 - isMarketOpen() - Seperate method to check if market is open (or) closed
  *  Future requirements
  *    - Mail at start of the day to inform opportunities ATH -10% & < 30 days
  *    - Mail at start of the day to inform opporunities in invested stocks
  */

/**
  * @ Check if market is open (or) closed
  *  Sheet: Daily_Log
  */
function isMarketOpen() {

   var marketStatus;
   //get Index return values using named range
   var indexReturn = SpreadsheetApp.getActive().getSheetByName("Daily_Log").getRange("Index_Returns").getValues();
   indexReturn.shift(); //remove headers
   indexReturn.forEach(function(value) {
      marketStatus = value[0];
   });

   return marketStatus;

};

/**
  * @ Sort Ok Index sheet by daily returns at end of the day before sending email
  *  Sheet: Ok_Index
  */
function sortOKIndexByDailyReturn() {

   var spreadsheet = SpreadsheetApp.getActive();
   var sheet = spreadsheet.getSheetByName('Ok_Index');
   var lastRow = sheet.getLastRow();
   //getRange(row, column, numRows, numColumns) 
   var range = sheet.getRange(2, 1, lastRow - 2, 13);
   // sort by column 8 which contains % today's change (Change - Today's)
   range.sort({
      column: 8,
      ascending: false
   });

};

/**
  * @ Copy daily ATH data and store it in ATH_Data sheet
  *  Sheet: ATH_Data
  */
function copyATHData() {
   //navigate to ATH data sheet
   var spreadsheet = SpreadsheetApp.getActive();
   spreadsheet.getRange('J17').activate();
   spreadsheet.setActiveSheet(spreadsheet.getSheetByName('ATH_Data'), true);
   spreadsheet.getRange('J:J').activate();
   //validate if we are on correct sheet before we edit the page
   var sheet = SpreadsheetApp.getActiveSheet();
   //Logger.log(sheet.getSheetName());
   if (sheet.getSheetName() == 'ATH_Data') {

      spreadsheet.getActiveSheet().insertColumnsBefore(spreadsheet.getActiveRange().getColumn(), 1);
      spreadsheet.getActiveRange().offset(0, 0, spreadsheet.getActiveRange().getNumRows(), 1).activate();
      spreadsheet.getRange('H:H').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

   } else {
      Logger.log("Error: Not in ATH_Data sheet");
   }

};

/**
  * @ Copy daily ATH data and store it in ATH_Data sheet
  *  Sheet: ATH_Data
  */
function copy1DChangeData() {
   //navigate to ATH data sheet
   var spreadsheet = SpreadsheetApp.getActive();
   spreadsheet.getRange('J17').activate();
   spreadsheet.setActiveSheet(spreadsheet.getSheetByName('1DChange_Data'), true);
   spreadsheet.getRange('J:J').activate();
   //validate if we are on correct sheet before we edit the page
   var sheet = SpreadsheetApp.getActiveSheet();
   //Logger.log(sheet.getSheetName());
   if (sheet.getSheetName() == '1DChange_Data') {

      spreadsheet.getActiveSheet().insertColumnsBefore(spreadsheet.getActiveRange().getColumn(), 1);
      spreadsheet.getActiveRange().offset(0, 0, spreadsheet.getActiveRange().getNumRows(), 1).activate();
      spreadsheet.getRange('H:H').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

   } else {
      Logger.log("Error: Not in 1DChange_Data sheet");
   }

};

/**
  * @ Copy daily returns value and paste it to below  to store historical data
  *  Sheet: Daily_Log
  */
function copyHistoricalData() {
   //navigate to daily log sheet
   var spreadsheet = SpreadsheetApp.getActive();
   spreadsheet.getRange('A1:L1').activate();
   spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Daily_Log'), true);
   spreadsheet.getRange('50:50').activate();
   //validate if we are on correct sheet
   var sheet = SpreadsheetApp.getActiveSheet();
   if (sheet.getSheetName() == 'Daily_Log') {
      spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
      spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
      spreadsheet.getRange('A50').activate();
      spreadsheet.getRange('A3:R3').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

      // enter formula to calculate NAV data
      spreadsheet.getRange('Q50').activate();
      spreadsheet.getCurrentCell().setFormula('=A50');
      spreadsheet.getRange('R50').activate();
      spreadsheet.getCurrentCell().setFormula('=R51+(R51*I50)');
      spreadsheet.getRange('S50').activate();
      spreadsheet.getCurrentCell().setFormula('=S51+(S51*J50)');
      spreadsheet.getRange('T50').activate();
      spreadsheet.getCurrentCell().setFormula('=T51+(T51*K50)');
      spreadsheet.getRange('U50').activate();
      spreadsheet.getCurrentCell().setFormula('=U51+(U51*L50)');
      spreadsheet.getRange('V50').activate();
      spreadsheet.getCurrentCell().setFormula('=V51+(V51*M50)');
   } else {
      Logger.log("Error: Not in Daily_Log sheet");
   }

}

/**
  * @get data from google sheet and store to array
  *  Sheet: Daily_Log
  */
function getData() {
   var okPortfolioLogs = [];
   var dailyLog = {};

   //get OK portfolio return values using named range
   var portfolioReturn = SpreadsheetApp.getActive().getSheetByName("Daily_Log").getRange("OK_Portfolio_Returns").getValues();
   portfolioReturn.shift(); //remove headers
   portfolioReturn.forEach(function(value) {
      dailyLog.date = value[0];
      dailyLog.totalCost = value[2];
      dailyLog.totalCurrentValue = value[3];
      dailyLog.overallProfitRs = value[4];
      dailyLog.overallProfitPercent = value[5];
      dailyLog.xirr = value[6];
      dailyLog.TodayProfitRs = value[7];
      dailyLog.TodayProfitPercent = value[8];
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
      if (dailyLog.market == "Open") {

         //Form dynamic messages based on todays return
         var dynamicMessage = "";
         if (dailyLog.TodayProfitRs > 0) {
            dynamicMessage = dynamicMessage + " increased by "
         } else {
            dynamicMessage = dynamicMessage + " decreased by "
         }

         text = text + "Todays OK portfolio" + dynamicMessage + (Math.round(dailyLog.TodayProfitPercent * 10000) / 100) + '%' + "\n" + "------------------------------------------------------------\n" + "OK Portfolio Returns	: " + "\n" + "Total Portfolio Cost (Rs.)	: " + "Rs. " + Math.round(dailyLog.totalCost) + "\n" + "Total Portfolio Value (Rs.) : " + "Rs. " + Math.round(dailyLog.totalCurrentValue) + "\n" + "Overall Profit / Loss (Rs.) : " + "Rs. " + Math.round(dailyLog.overallProfitRs) + "\n" + "Overall Profit / Loss (%) : " + (Math.round(dailyLog.overallProfitPercent * 10000) / 100) + '%' + "\n" + "Portfolio XIRR (%) : " + (Math.round(dailyLog.xirr * 10000) / 100) + '%' + "\n" + "Today Profit / Loss (Rs.) : " + "Rs. " + Math.round(dailyLog.TodayProfitRs) + "\n" + "Today Profit / Loss  (%) : " + (Math.round(dailyLog.TodayProfitPercent * 10000) / 100) + '%' + "\n" + "------------------------------------------------------------\n" + "Nifty 50 Index: " + (Math.round(dailyLog.nifty50 * 10000) / 100) + '%' + "\n" + "Nifty Next 50 Index: " + (Math.round(dailyLog.niftynext50 * 10000) / 100) + '%' + "\n" + "Midcap Index : " + (Math.round(dailyLog.midcap * 10000) / 100) + '%' + "\n" + "Smallcap Index : " + (Math.round(dailyLog.smallcap * 10000) / 100) + '%' + "\n" + "NASDAQ 100 Index : " + (Math.round(dailyLog.nasdaq100 * 10000) / 100) + '%' + "\n" + "Gold : " + (Math.round(dailyLog.gold * 10000) / 100) + '%' + "\n" + "Silver : " + (Math.round(dailyLog.silver * 10000) / 100) + '%' + "\n" + "India VIX : " + (Math.round(dailyLog.vix * 10000) / 100) + '%' + "\n" + "------------------------------------------------------------\n" +

         "Automatic script execution at " + dailyLog.date + "\n";
      } else {

         //else if market is closed today
         text = text + "Market is closed today " + "\n" + "------------------------------------------------------------\n" + "Automatic script execution at " + dailyLog.date + "\n";
      }

   });

   Logger.log(text);
   return text;
}

/**
  * @form email request
  */
function sendDailyReturnsEmail() {

   // send email only if market is open

   var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy") var dailyLogsData = getData();
   var body = getEmailText(dailyLogsData);

   //hide unwanted sheet so that it doesnt gets added to attachment
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Notes").hideSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Watchlist").hideSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions_OSV").hideSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Wallet").hideSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ATH_Data").hideSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("1DChange_Data").hideSheet();

   // permanently hide everytime
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary_OSV").hideSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Static_Data").hideSheet();

   var recipient = "emailId1id1@gmail.com";
   var subject = "Ok Portfolio Returns for " + date;
   var options = {
      //cc: "emailid2@gmail.com",
      bcc: "emailid3@gmail.com",
      //replyTo: "help@example.com"
      attachments: [SpreadsheetApp.getActiveSpreadsheet().getAs(MimeType.PDF).setName("OK Portfolio Summary " + date + " .pdf")]
   }
   MailApp.sendEmail(recipient, subject, body, options);

   //activate all hidden sheets that are required
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Notes").showSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Watchlist").showSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions_OSV").showSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Wallet").showSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ATH_Data").showSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("1DChange_Data").showSheet();
};

/**
  * @Trigger scheduler for Ok Sheet -  Scheduled to trigger at 4PM IST daily
  */
function triggerSchedulerOKSheet() {

   //copy current day data to historical data table and ATH values
   copyHistoricalData();

   // run below functions only if market is open 
   if (isMarketOpen() == "Open") {
      // copy ATH data to sheet
      copyATHData();
      // copy 1D change data to sheet
      copy1DChangeData();
      //sort OK Index sheet by daily returns before sending report
      sortOKIndexByDailyReturn();
      // send the daily report as email with attachment
      sendDailyReturnsEmail();
   }

};