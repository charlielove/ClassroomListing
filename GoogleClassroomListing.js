//get the first empty row in the sheet
function getFirstEmptyRow(sheet) {
  
  var row = sheet.getRange("L1").getValue(); 
  return (row);
}


//get the pageSize value to use in the sheet
function getPageSize(sheet) {
  var row = sheet.getRange("M1").getValue(); 
  return (row);
}

Date.prototype.yyyymmdd = function() {
    var mm = this.getMonth() + 1; // getMonth() is zero-based
    var dd = this.getDate();

     return [this.getFullYear(),
          (mm>9 ? '' : '0') + mm,
          (dd>9 ? '' : '0') + dd
         ].join('');
};


function listClasses(){

  //by Charlie Love
  //charlielove.org tw: @charlie_love

  //let fire up the batch execution of the script
  startOrResumeContinousExecutionInstance("listClasses");

  //open the spreadsheet
  var date = new Date();
  var dateString = date.yyyymmdd();

  var my_ss = dateString + "ClassroomListing";
  var files = DriveApp.getFilesByName(my_ss);
  var file = !files.hasNext() ? SpreadsheetApp.create(my_ss) : files.next();
  var ss = SpreadsheetApp.openById(file.getId())
  try 
  {
     ss.setActiveSheet(ss.getSheetByName(my_sheet));
  } catch (e){;} 
  
  //set the activeSheet
  var sheet = ss.getActiveSheet();
  
  //if we don't have a batch key set then we are at the start so clear the sheet, add the headings and set the start to the first row.
  var startRow = getFirstEmptyRow(sheet);

  if(startRow == '') {
    sheet.clear();
    sheet.appendRow(["No.","Class Owner","Organization","Creation Date","Last Updated","Course State","Course Section","Course Name","Enrollment Code"]);  
    //start at row 0
    var startRow = 0;
    // it isn't set, start with an empty token
    var nextPageToken = '';
    setBatchKey("listClasses", nextPageToken);
  } else {
    //let's get a page of results for classroom listings 
    //get the token we are using to resume the batches
    //this is the nextPageToken
    var nextPageToken = getBatchKey("listClasses");
  }
 
  //get the pageSizeValue stored in the sheet
  var pageSizeValue = getPageSize(sheet);
  //if it isn't set then start at 400
  if (pageSizeValue == '') {
    pageSizeValue = 400;
    var cell = sheet.getRange("M1");
    cell.setValue(pageSizeValue);
  } 
 
  //and now we'll loop arround retrieving a batch of results, 
  //writing them to the spreadsheet and then getting the next batch etc.
  var errorflag = false;
  
  do {
    //get list of course details
    var optionalArgs = {
      pageSize: pageSizeValue,
      pageToken: nextPageToken
    };
    try {   var courses = Classroom.Courses.list(optionalArgs);
            var nextPageToken = courses.nextPageToken;
  
            //loop round
            for ( var i= 0, len = courses.courses.length; i < len; i++) {
                 var courseName =  courses.courses[i].name;
                 var courseCreation = courses.courses[i].creationTime;
                 var courseUpdated = courses.courses[i].updateTime;
                 var courseSection = courses.courses[i].section;
                 var courseCode = courses.courses[i].enrollmentCode
                 if (courseSection == null) {
                     courseSection = "";
                 }
                 var courseState = courses.courses[i].courseState;
                 var owner = courses.courses[i].ownerId;
      
                 try{
                    ownerObj = AdminDirectory.Users.get(owner);}
                    catch(err){owner += ": " +err.message; }
                    owner = ownerObj.name.fullName;
                    var ou = ownerObj.orgUnitPath;
    
                    ss.getSheets()[0].appendRow([startRow+1,owner,ou,courseCreation,courseUpdated,courseState,courseSection, courseName.toString(),courseCode]);
                    startRow++;  //we've written a row, so add one to start row.
        
                    //write the row value into the sheet to read later 
                    var cell = sheet.getRange("L1");
                    cell.setValue(startRow);
            }
            setBatchKey("listClasses", nextPageToken); 
            
        } catch (e ){ //error thown because we've fewer classrooms than the page size
               
               //if the page size is already one then we've run out of classrooms
               if (pageSizeValue == 1) {
                 errorflag = true;
               }
               //let's set the pageSize to 1 and get the last few one at a time
               pageSizeValue = 1;
               //write the pageSizeValue into the Sheet
               var cell = sheet.getRange("M1");
               cell.setValue(pageSizeValue);
      }

    } while ((nextPageToken != undefined) && (isTimeRunningOut("listClasses") != true) && (errorflag != true )); //and do this until there are no more pages of results to get
    
    //we've run out of classrooms
    if ((nextPageToken == undefined)||(errorflag == true)){
       endContinuousExecutionInstance(listClasses, "YourEmailAddress", "Classroom");
    }
  
  
}


/**
 *  ---  Continous Execution Library ---
 *
 *  Copyright (c) 2013 Patrick Martinent
 *
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 *  distributed under the License is distributed on an "AS IS" BASIS,
 *  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *  See the License for the specific language governing permissions and
 *  limitations under the License.
 */


/*************************************************************************
* Call this function at the start of your batch script
* it will create the necessary UserProperties with the fname
* so that it can keep managing the triggers until the batch
* execution is complete. It will store the start time for the
* email it sends out to you when the batch has completed
*
* @param {fname} str The batch function to invoke repeatedly.
*/
function startOrResumeContinousExecutionInstance(fname){
  var userProperties = PropertiesService.getUserProperties();
  var start = userProperties.getProperty('GASCBL_' + fname + '_START_BATCH');
  if (start === "" || start === null)
  {
    start = new Date();
    userProperties.setProperty('GASCBL_' + fname + '_START_BATCH', start);
    userProperties.setProperty('GASCBL_' + fname + '_KEY', "");
  }
  
  userProperties.setProperty('GASCBL_' + fname + '_START_ITERATION', new Date());
  
  deleteCurrentTrigger_(fname);
  enableNextTrigger_(fname); 
}

/*************************************************************************
* In order to be able to understand where your batch last executed you
* set the key ( or counter ) everytime a new item in your batch is complete
* when you restart the batch through the trigger, use getBatchKey to start 
* at the right place
*
* @param {fname} str The batch function we are continuously triggering.
* @param {key} str The batch key that was just completed.
*/
function setBatchKey(fname, key){
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('GASCBL_' + fname + '_KEY', key);
}

/*************************************************************************
* This function returns the current batch key, so you can start processing at
* the right position when your batch resumes from the execution of the trigger
*
* @param {fname} str The batch function we are continuously triggering.
* @returns {string} The batch key which was last completed.
*/
function getBatchKey(fname){
  var userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty('GASCBL_' + fname + '_KEY');
}


/*************************************************************************
* When the batch is complete run this function, and pass it an email and
* custom title so you have an indication that the process is complete as
* well as the time it took
*
* @param {fname} str The batch function we are continuously triggering.
* @param {emailRecipient} str The email address to which the email will be sent.
* @param {customTitle} str The custom title for the email.
*/
function endContinuousExecutionInstance(fname, emailRecipient, customTitle){
  var userProperties = PropertiesService.getUserProperties();
  var end = new Date();
  var start = userProperties.getProperty('GASCBL_' + fname + '_START_BATCH');
  var key = userProperties.getProperty('GASCBL_' + fname + '_KEY');

  var emailTitle = customTitle + " : Continuous Execution Script";
  var body = "Started : " + start + "<br>" + "Ended :" + end + "<br>" + "LAST KEY : " + key; 
  MailApp.sendEmail(emailRecipient, emailTitle, "", {htmlBody:body});
  
  deleteCurrentTrigger_(fname);
  userProperties.deleteProperty('GASCBL_' + fname + '_START_ITERATION');
  userProperties.deleteProperty('GASCBL_' + fname + '_START_BATCH');
  userProperties.deleteProperty('GASCBL_' + fname + '_KEY');
  userProperties.deleteProperty('GASCBL_' + fname);
}

/*************************************************************************
* Call this function when finishing a batch item to find out if we have
* time for one more. if not exit elegantly and let the batch restart with
* the trigger
*
* @param {fname} str The batch function we are continuously triggering.
* @returns (boolean) wether we are close to reaching the exec time limit
*/
function isTimeRunningOut(fname){
  var userProperties = PropertiesService.getUserProperties();
  var start = new Date(userProperties.getProperty('GASCBL_' + fname + '_START_ITERATION'));
  var now = new Date();
  
  var timeElapsed = Math.floor((now.getTime() - start.getTime())/1000);
  return (timeElapsed > 270);
}

/*
* Set the next trigger, 7 minutes in the future
*/
function enableNextTrigger_(fname) {
  var userProperties = PropertiesService.getUserProperties();
  var nextTrigger = ScriptApp.newTrigger(fname).timeBased().after(9 * 60 * 1000).create();
  var triggerId = nextTrigger.getUniqueId();

  userProperties.setProperty('GASCBL_' + fname, triggerId);
}

/*
* Deletes the current trigger, so we don't end up with undeleted
* time based triggers all over the place
*/
function deleteCurrentTrigger_(fname) {
  var userProperties = PropertiesService.getUserProperties();
  var triggerId = userProperties.getProperty('GASCBL_' + fname);
  var triggers = ScriptApp.getProjectTriggers();
  for (var i in triggers) {
    if (triggers[i].getUniqueId() === triggerId)
      ScriptApp.deleteTrigger(triggers[i]);
  }
  userProperties.setProperty('GASCBL_' + fname, "");
}
