/* ==================================================================================================

For tech spreadsheet:  DO NOT DELETE!
https://docs.google.com/spreadsheets/d/1t39_Xq2XnzR7a16J0zva9Lx_aLXj6gM6Saf4a344SNY/edit?gid=1787496996#gid=1787496996

================================================================================================== */

//////////////////////////////////   SET VARIABLES   /////////////////////////////////////////////

const SHEETNAME   = "NO_SCHOOL_TASK";
const SUMMERNAME  = "SUMMER_TASK";
const VARSSHEET   = "ADMIN"; 
const EMAILADMIN  = 'me@email.com'; // owner email
const EMAILON     = true;

const date = new Date();
const currentDay = date.toLocaleString('en-US', { month:'numeric', day:'numeric', year:'numeric' });

//////////////////////////////////   PULL VARIABLES   /////////////////////////////////////////////

let spsh = SpreadsheetApp.getActiveSpreadsheet();
let actsh = spsh.getSheetByName(VARSSHEET);

const TASKNAMECOL = 0; // task name column [A]
const PRIORITYCOL = 1; // column [B]
const OWNERCOL    = 2; // column [C]
const STATUSCOL   = 3; // status column [D]
const STARTDATECOL= 4; // start date column [E]
const ENDDATECOL  = 5; // end date column [F]
const NOTESCOL    = 6; // column [G]
const REFSTATUSCOL= 7; // ref status column [H]
const INPROGVAL   = actsh.getRange(3,1).getValue(); // value of new status
const ONHOLDVAL   = actsh.getRange(4,1).getValue(); // on hold value
const COMPLETEVAL = actsh.getRange(5,1).getValue(); // value of status

///////////////////////////////////////////////////////////////////////////////////////////////////


/* =================================================================================================== */

function RunAllTasks() {

    ClearCompleted();

    UpdateStatusByDate();

};

function RunSummerTasks() {

    ClearSummerCompleted();

    UpdateSummerStatusByDate();

};

/* ===================================================================================================== */

function ClearCompleted() {
  
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheetByName(SHEETNAME);

      SpreadsheetApp.setActiveSheet(sh);

      var rows = sh.getDataRange();
      var numRows = rows.getNumRows();
      var values = rows.getValues();
      var eLastRow = rows.getLastRow();
      
      Logger.log('Checking for completed task...');
      
      var rowsDeleted = 0;

      for (var i=0; i <= numRows - 1; i++) {

          var row = values[i];
          
          if (row[STATUSCOL] == COMPLETEVAL || row[STATUSCOL] == '') { // searches for completed tasks & delete them
            
              // Insert blank row at end of last row BEFORE deletion loop, otherwise deletion will crash! 
              if(eLastRow < 3) sh.insertRowsAfter(eLastRow, 1);
              
              sh.deleteRow((parseInt(i) + 1) - rowsDeleted);
              rowsDeleted++;
              
              Logger.log('Deleted row: ' + i);
              Logger.log('Num Deleted: ' + rowsDeleted);
              
          };

      };

      Logger.log('Cleanup completed');
  
}; //end function


function ClearSummerCompleted() {
  
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheetByName(SUMMERNAME);

      SpreadsheetApp.setActiveSheet(sh);

      var rows = sh.getDataRange();
      var numRows = rows.getNumRows();
      var values = rows.getValues();
      var eLastRow = rows.getLastRow();
      
      Logger.log('Checking for completed task...');
      
      var rowsDeleted = 0;

      for (var i=0; i <= numRows - 1; i++) {

          var row = values[i];
          
          if (row[STATUSCOL] == COMPLETEVAL || row[STATUSCOL] == '') { // searches for completed tasks & delete them
            
              // Insert blank row at end of last row BEFORE deletion loop, otherwise deletion will crash! 
              if(eLastRow < 3) sh.insertRowsAfter(eLastRow, 1);
              
              sh.deleteRow((parseInt(i) + 1) - rowsDeleted);
              rowsDeleted++;
              
              Logger.log('Deleted row: ' + i);
              Logger.log('Num Deleted: ' + rowsDeleted);
              
          };

      };

      Logger.log('Cleanup completed');
  
}; //end function

/* ===================================================================================================== */

function DeleteBlankRows() {

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheetByName(SHEETNAME);
      var su = ss.getSheetByName(SUMMERNAME);

      SpreadsheetApp.setActiveSheet(sh);
      
      var maxRows = sh.getMaxRows(); 
      var lastRow = sh.getLastRow();

      if (maxRows + lastRow > 3) {
            
            sh.deleteRows(lastRow + 1, maxRows - lastRow);
            Logger.log('NST - Empty rows found & deleted');
      
      }; // end if

      SpreadsheetApp.setActiveSheet(su);
      
      var maxRows = su.getMaxRows(); 
      var lastRow = su.getLastRow();

      if (maxRows + lastRow > 3) {
            
            su.deleteRows(lastRow + 1, maxRows - lastRow);
            Logger.log('SUM - Empty rows found & deleted');
      
      }; // end if

}; // end function

/* ===================================================================================================== */

function UpdateStatusByDate() {

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheetByName(SHEETNAME);

      SpreadsheetApp.setActiveSheet(sh);

      var rows = sh.getDataRange().getValues();
      // Check each row data in the balance sheet. Ignore header row
      
      for (var i=0; i < rows.length; i++) {
        
          var start = rows[i][STARTDATECOL];
          var end   = rows[i][ENDDATECOL];
          var ref   = rows[i][REFSTATUSCOL];
          var task  = rows[i][TASKNAMECOL];
          var email = rows[i][OWNERCOL];

          var startDate = start.toLocaleString('en-US',{ month:'numeric', day:'numeric', year:'numeric' });
          var endDate = end.toLocaleString('en-US', { month:'numeric', day:'numeric', year:'numeric' });

          // convert dates to a numeric value
          var currDay = date.valueOf();
          var startVal = start.valueOf();
          var endVal = end.valueOf();

          Logger.log('currentDay: ' + currentDay);
          Logger.log('currDay: ' + currDay);
          Logger.log('start: ' + startDate);
          Logger.log('startVal: ' + startVal);
          Logger.log('end: ' + endDate);
          Logger.log('endVal: ' + endVal);


          // Start check process
          switch(true) {

                case (currDay > endVal) :

                      // End date has expired!
                      sh.getRange(i+1, STATUSCOL + 1).setValue(ONHOLDVAL);
                      Logger.log('Found outdated task... Updated as ' + ONHOLDVAL);

                      if(ref != 2) sh.getRange(i+1, REFSTATUSCOL + 1).setValue('0');

                      // Send Email out
                      if(EMAILON) {
                          
                            if(ref == 1 || ref == 2) { Logger.log("No emails sent."); }

                            else { 
                              
                                  // sendEmail(recipient, subject, body, options)   
                                  MailApp.sendEmail(email, `On-Hold Task is PAST DUE! - ${endDate}`, 'Hello!', 
                                                    { name: 'Action Support domain.org',
                                                      noReply: true,
                                                      bcc: EMAILADMIN,
                                                      htmlBody: `[This is an auto-generated message]<br>
                                                                ====================================<br>
                                                                <font size="+1">
                                                                You have a task that is marked PAST DUE, and needs your attention!<br>
                                                                Please refer to the spreadsheet for details!<br>
                                                                Task: ${task}<br>
                                                                Start Date: ${startDate}<br>
                                                                Due Date: ${endDate}<br>
                                                                </font>` 
                                                    });
                                  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
                                  Logger.log("Remaining email quota: " + emailQuotaRemaining);

                                  sh.getRange(i+1, REFSTATUSCOL + 1).setValue('2');
                                  Logger.log('Emails sent.');

                            }; // end if/else

                      }; // end email on

                break;
                
                case (currDay >= startVal) :
                      
                      // Today is start date!
                      sh.getRange(i+1, STATUSCOL + 1).setValue(INPROGVAL);
                      Logger.log('Found tasks for you today... Updated as ' + INPROGVAL);

                      // Send Email out
                      if(EMAILON) {

                            if(ref == 1) { Logger.log("No emails sent."); }

                            else { 
                                  // sendEmail(recipient, subject, body, options)   
                                  MailApp.sendEmail(email, `NEW Out of School Task Is Ready! - ${startDate}`, 'Hello!', 
                                                    { name: 'Action Support domain.org',
                                                      noReply: true,
                                                      htmlBody: `[This is an auto-generated message]<br>
                                                                ====================================<br>
                                                                <font size="+1">
                                                                A new "out of school" task is ready!<br>
                                                                Please refer to the spreadsheet for details!<br>
                                                                Task: ${task}<br>
                                                                Start Date: ${startDate}<br>
                                                                Due Date: ${endDate}<br>
                                                                </font>` 
                                                    });
                                  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
                                  Logger.log("Remaining email quota: " + emailQuotaRemaining);

                                  sh.getRange(i+1, REFSTATUSCOL + 1).setValue('1');
                                  Logger.log('Emails sent.');

                            }; // end if/else

                      }; // end email on

                break;
                
                default : 
                      Logger.log('No emails sent.');
                break;

          }; // end switch case 

      }; // end loop

}; // end function

/* ===================================================================================================== */

function UpdateSummerStatusByDate() {

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheetByName(SUMMERNAME);

      SpreadsheetApp.setActiveSheet(sh);

      var rows = sh.getDataRange().getValues();
      // Check each row data in the balance sheet. Ignore header row
      
      for (var i=0; i < rows.length; i++) {
        
          var start = rows[i][STARTDATECOL];
          var end   = rows[i][ENDDATECOL];
          var ref   = rows[i][REFSTATUSCOL];
          var task  = rows[i][TASKNAMECOL];
          var email = rows[i][OWNERCOL];

          var startDate = start.toLocaleString('en-US',{ month:'numeric', day:'numeric', year:'numeric' });
          var endDate = end.toLocaleString('en-US', { month:'numeric', day:'numeric', year:'numeric' });

          // convert dates to a numeric value
          var currDay = date.valueOf();
          var startVal = start.valueOf();
          var endVal = end.valueOf();

          Logger.log('currentDay: ' + currentDay);
          Logger.log('currDay: ' + currDay);
          Logger.log('start: ' + startDate);
          Logger.log('startVal: ' + startVal);
          Logger.log('end: ' + endDate);
          Logger.log('endVal: ' + endVal);


          // Start check process
          switch(true) {

                case (currDay > endVal) :

                      // End date has expired!
                      sh.getRange(i+1, STATUSCOL + 1).setValue(ONHOLDVAL);
                      Logger.log('Found outdated task... Updated as ' + ONHOLDVAL);

                      if(ref != 2) sh.getRange(i+1, REFSTATUSCOL + 1).setValue('0');

                      // Send Email out
                      if(EMAILON) {
                          
                            if(ref == 1 || ref == 2) { Logger.log("No emails sent."); }

                            else { 
                              
                                  // sendEmail(recipient, subject, body, options)   
                                  MailApp.sendEmail(email, `On-Hold Summer Task is PAST DUE! - ${endDate}`, 'Hello!', 
                                                    { name: 'Action Support domain.org',
                                                      noReply: true,
                                                      bcc: EMAILADMIN,
                                                      htmlBody: `[This is an auto-generated message]<br>
                                                                ====================================<br>
                                                                <font size="+1">
                                                                You have a summer task that is marked PAST DUE, and needs your attention!<br>
                                                                Please refer to the spreadsheet for details!<br>
                                                                Task: ${task}<br>
                                                                Start Date: ${startDate}<br>
                                                                Due Date: ${endDate}<br>
                                                                </font>` 
                                                    });
                                  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
                                  Logger.log("Remaining email quota: " + emailQuotaRemaining);

                                  sh.getRange(i+1, REFSTATUSCOL + 1).setValue('2');
                                  Logger.log('Emails sent.');

                            }; // end if/else

                      }; // end email on

                break;
                
                case (currDay >= startVal) :
                      
                      // Today is start date!
                      sh.getRange(i+1, STATUSCOL + 1).setValue(INPROGVAL);
                      Logger.log('Found tasks for you today... Updated as ' + INPROGVAL);

                      // Send Email out
                      if(EMAILON) {

                            if(ref == 1) { Logger.log("No emails sent."); }

                            else { 
                                  // sendEmail(recipient, subject, body, options)   
                                  MailApp.sendEmail(email, `NEW Summer Task Is Ready! - ${startDate}`, 'Hello!', 
                                                    { name: 'Action Support domain.org',
                                                      noReply: true,
                                                      htmlBody: `[This is an auto-generated message]<br>
                                                                ====================================<br>
                                                                <font size="+1">
                                                                A new SUMMER task is ready!<br>
                                                                Please refer to the spreadsheet for details!<br>
                                                                Task: ${task}<br>
                                                                Start Date: ${startDate}<br>
                                                                Due Date: ${endDate}<br>
                                                                </font>` 
                                                    });
                                  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
                                  Logger.log("Remaining email quota: " + emailQuotaRemaining);

                                  sh.getRange(i+1, REFSTATUSCOL + 1).setValue('1');
                                  Logger.log('Emails sent.');

                            }; // end if/else

                      }; // end email on

                break;
                
                default : 
                      Logger.log('No emails sent.');
                break;

          }; // end switch case 

      }; // end loop

}; // end function
