/** 
 * Stock Alert Portfolio Sheet Automation
 * Used to automate the stock laerts when it reaches support zone and when super trend indicator gives signal
 * Change log V1
 *      - 25/03/2025 - Base version of stock alert sheet
 *      - 28/03/2025 - Seperated super trend and stock alert methods
 */

/** -------------------------------- Super Trend related methods -------------------------------- */

/**
 * @ Copy daily returns value and paste it to below  to store historical data
 *  Sheet: 1D_ST_Data
 */

function CopySuperTrendData() {
   var spreadsheet = SpreadsheetApp.getActive();
   spreadsheet.getRange('H2').activate();
   spreadsheet.setActiveSheet(spreadsheet.getSheetByName('1D_ST_Data'), true);
   spreadsheet.getRange('F:F').activate();

   //validate if we are on correct sheet before we edit the page
   var sheet = SpreadsheetApp.getActiveSheet();
   if (sheet.getSheetName() == '1D_ST_Data') {

      spreadsheet.getActiveSheet().insertColumnsBefore(spreadsheet.getActiveRange().getColumn(), 1);
      spreadsheet.getActiveRange().offset(0, 0, spreadsheet.getActiveRange().getNumRows(), 1).activate();
      spreadsheet.getRange('D:D').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

   } else {
      Logger.log("Error: Not in 1D_ST_Data sheet");
   }

};

/**
 * @check if there is alert to be triggered
 *  Sheet: SuperTrend-Index
 */
function isSTAlertTriggered() {

   //set default flag to no
   var isAlertTriggered = "No";
   var superTrendData = SpreadsheetApp.getActive().getSheetByName("SuperTrend-Index").getRange("SuperTrendData").getValues();
   superTrendData.shift(); //remove headers
   superTrendData.forEach(function(value) {
      //save value to log only if Send alert column [9] has value as "yes"
      if (value[9] == "Yes") {
         isAlertTriggered = "Yes";
      };
   })

   return isAlertTriggered;
}

/**
 * @get data from google sheet and store to array
 *  Sheet: SuperTrend-Index
 */
function getSTData() {

   var superTrendLogs = [];
   //get data from sheet using named range
   var superTrendData = SpreadsheetApp.getActive().getSheetByName("SuperTrend-Index").getRange("SuperTrendData").getValues();
   superTrendData.shift(); //remove headers
   superTrendData.forEach(function(value) {
      //save value to log only if Send alert column [9] has value as "yes"
      if (value[9] == "Yes") {
         var dailyLog = {};
         dailyLog.symbol = value[1];
         dailyLog.indexName = value[2];
         dailyLog.superTrend = value[7];
         dailyLog.sendAlert = value[9];
         superTrendLogs.push(dailyLog);
      };
   })

   //Logger.log(superTrendLogs);
   return superTrendLogs;

};

/**
 * @form email template
 */
function getSTEmailText(dailyLogData) {

   var bodyText = "Alert triggered for MA Super Trend Indicator in below Index " + "\n";

   dailyLogData.forEach(function(dailyLog) {
      bodyText = bodyText + "\n" + "------------------------------------------------------------\n" + "Index Symbol	: " + dailyLog.symbol + "\n" + "Index Name	: " + dailyLog.indexName + "\n" + "SuperTrend Status	: " + dailyLog.superTrend + "\n" + "\n" + "------------------------------------------------------------\n";

   });

   Logger.log(bodyText);
   return bodyText;
};

/**
 * @hide unwanted sheets in attachements
 */
function hideSTSheets() {

   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("1D_ST_Data").hideSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ST_TradeLog").hideSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("StockAlerts").hideSheet();

};

/**
 * @ unhide sheets after attaching in email
 */
function unHideSTSheets() {

   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("1D_ST_Data").showSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ST_TradeLog").showSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("StockAlerts").showSheet();

};

/**
 * @form email request
 */
function sendSTEmailAlert() {

   var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy") var dailyLogsData = getSTData();
   var body = getSTEmailText(dailyLogsData);

   hideSTSheets();

   var recipient = "emailId1@gmail.com" + "," + "emailId2@gmail.com";
   var subject = "Alert triggered for MA Super Trend Indicator " + date;
   var options = {
      //cc: "emailId3@gmail.com",
      bcc: "emailId4@gmail.com",
      //replyTo: "help@example.com"
      attachments: [SpreadsheetApp.getActiveSpreadsheet().getAs(MimeType.PDF).setName("Super trend index data " + date + " .pdf")]
   }
   MailApp.sendEmail(recipient, subject, body, options);

   unHideSTSheets();

};

/**
 * @Trigger scheduler for Trading Portfolio Sheet -  Scheduled to trigger at 10AM ,4PM IST daily
 */
function triggerSchedulerTradingPortfolio() {

   //check if the ST alert needs to be triggered and if triggered send email
   var checkIfSTAlertTriggered = isSTAlertTriggered();
   if (checkIfSTAlertTriggered == "Yes") {
      sendSTEmailAlert();
   } else {
      Logger.log("No alert triggered");
   };

   //copy current day data to historical data table
   CopySuperTrendData();

};

/** -------------------------------- Stock Alerts related methods -------------------------------- */

/**
 * @check if there is alert to be triggered
 *  Sheet: StockAlerts
 */
function isStockAlertTriggered() {

   //set default flag to no
   var isAlertTriggered = "No";
   var stockAlertData = SpreadsheetApp.getActive().getSheetByName("StockAlerts").getRange("StockAlertsData").getValues();
   stockAlertData.shift(); //remove headers
   stockAlertData.forEach(function(value) {
      //save value to log only if Send alert column [9] has value as "yes"
      if (value[8] == "Yes" && value[9] != "Off") {
         isAlertTriggered = "Yes";
      };
   })

   Logger.log(isAlertTriggered);
   return isAlertTriggered;
}

/**
 * @get data from google sheet and store to array
 *  Sheet: SuperTrend-Index
 */
function getStockAlertsData() {

   var StockAlertLogs = [];
   //get data from sheet using named range
   var stockAlertData = SpreadsheetApp.getActive().getSheetByName("StockAlerts").getRange("StockAlertsData").getValues();
   stockAlertData.shift(); //remove headers
   stockAlertData.forEach(function(value) {
      //save value to log only if Send alert column [9] has value as "yes"
      if (value[8] == "Yes" && value[9] != "Off") {
         var dailyLog = {};
         dailyLog.stockSymbol = value[2];
         dailyLog.stockName = value[3];
         dailyLog.currentPrice = value[4];
         dailyLog.target = value[6];
         dailyLog.ltp = value[7];
         StockAlertLogs.push(dailyLog);
      };
   })

   //Logger.log(StockAlertLogs);
   return StockAlertLogs;

};

/**
 * @form email template
 */
function getStockAlertEmailText(dailyLogData) {

   var bodyText = "Stock alert triggered for the below stocks " + "\n";

   dailyLogData.forEach(function(dailyLog) {
      bodyText = bodyText + "\n" + "------------------------------------------------------------\n" + "Stock symbol	: " + dailyLog.stockSymbol + "\n" + "Stock name	: " + dailyLog.stockName + "\n" + "Stock target price	: " + dailyLog.ltp + " " + dailyLog.target + "\n" + "Stock current price	: " + dailyLog.currentPrice + "\n" +

      "\n" + "------------------------------------------------------------\n";

   });

   Logger.log(bodyText);
   return bodyText;
};

/**
 * @hide unwanted sheets in attachements
 */
function hideStockAlertSheets() {

   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("1D_ST_Data").hideSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ST_TradeLog").hideSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SuperTrend-Index").hideSheet();

};

/**
 * @ unhide sheets after attaching in email
 */
function unHideStockAlertSheets() {

   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("1D_ST_Data").showSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ST_TradeLog").showSheet();
   SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SuperTrend-Index").showSheet();

};

/**
 * @form email request
 */
function sendStockAlertEmail() {

   var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy") var dailyLogsData = getStockAlertsData();
   var body = getStockAlertEmailText(dailyLogsData);

   hideStockAlertSheets();

   var recipient = "emailId1@gmail.com" + "," + "emailId2@gmail.com";
   var subject = "Stock alert triggered for the below monitored stocks " + date;
   var options = {
      //cc: "emailId3@gmail.com",
      bcc: "emailId4@gmail.com",
      //replyTo: "help@example.com"
      attachments: [SpreadsheetApp.getActiveSpreadsheet().getAs(MimeType.PDF).setName("Stock alerts data " + date + " .pdf")]
   }
   MailApp.sendEmail(recipient, subject, body, options);

   unHideStockAlertSheets();

};

/**
 * @Trigger scheduler for Trading Portfolio Sheet -  Scheduled to trigger at 10AM ,4PM IST daily
 */
function triggerSchedulerStockAlerts() {

   //check if the stock alert needs to be triggered and if triggered send email
   var checkIfStockAlertTriggered = isStockAlertTriggered();
   if (checkIfStockAlertTriggered == "Yes") {
      sendStockAlertEmail();
   } else {
      Logger.log("No alert triggered");
   };

};

/** -------------------------------- future methods -------------------------------- */

//end of macro
