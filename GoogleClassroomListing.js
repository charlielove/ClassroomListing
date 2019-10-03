/*********************************
* set your admin email adress here
*/
function getEmailRecipient() {
  return "admin@yourdomain.org";
}

// consider using Cahce Service to store values and arrays?
// https://stackoverflow.com/questions/7854573/exceeded-maximum-execution-time-in-google-apps-script

// get the pageSize value to use in the sheet
function getPageSize(sheet) { 
  return (sheet.getRange("M1").getValue());
}

// set the pageSize value to use in the sheet
function setPageSize(value, sheet) {
    sheet.getRange("M1").setValue(value);
}

// function for date format
Date.prototype.yyyymmdd = function() {
    var mm = this.getMonth() + 1; // getMonth() is zero-based
    var dd = this.getDate();

     return [this.getFullYear(),
          (mm>9 ? '' : '0') + mm,
          (dd>9 ? '' : '0') + dd
         ].join('');
};

// main function for listing courses
function listClasses(){

  // lets fire up the batch execution of the script
  startOrResumeContinousExecutionInstance("listClasses");

  // open the spreadsheet
  var date = new Date();
  var dateString = date.yyyymmdd();
  var my_ss = dateString + "ClassroomListing";
  var files = DriveApp.getFilesByName(my_ss);
  var file = !files.hasNext() ? SpreadsheetApp.create(my_ss) : files.next();
  var ss = SpreadsheetApp.openById(file.getId())
  var sheet = ss.getSheets()[0];
  
  // setup and rename needed sheets
  // rename first sheet if not already renamed
  try {
    ss.setActiveSheet(ss.getSheetByName('ClassroomList'));
  } catch (e) {
    ss.renameActiveSheet('ClassroomList');
  }
  
  // create a second sheet for storing id:s that we already have "translated" into names
  try {
    ss.setActiveSheet(ss.getSheetByName('idToName'));
  } catch (e) {
    ss.insertSheet('idToName', 1);
  }
  var ownerSheet = ss.getSheetByName('idToName');
  
  // read from ownerSheet into an array
  if (ownerSheet.getLastRow() != 0) {
    var ownerArray = ownerSheet.getRange(1,1,ownerSheet.getLastRow(),3).getValues();
  } else {
    var ownerArray = [];
  }
  
  // if we don't have a batch key set then we are at the start so clear the sheet
  // add the headings and set the start to the first row.
  var startRow = sheet.getLastRow();
  if(startRow == 0) {
    sheet.clear();
    sheet.appendRow(["No.","Class Owner","Primary Email","Organization","Creation Date","Last Updated",
                     "Course State","Course Section","Course Name","Enrollment Code","Course Id"]);  
    // it isn't set, start with an empty token
    var nextPageToken = '';
    setBatchKey("listClasses", nextPageToken);
  } else {
    // subtract one from startrow (adjust to header counting as one row)
    startRow -= 1;
    // get the token we are using to resume the batches
    var nextPageToken = getBatchKey("listClasses");
  }
 
  // get the pageSizeValue stored in the sheet
  var pageSizeValue = getPageSize(sheet);
  // if it isn't set then start at 300
  if (pageSizeValue == '') {
    pageSizeValue = 300;
    setPageSize(pageSizeValue, sheet);
  }
  
  // if resuming the batch - we must add at least one row below all results
  if (sheet.getLastRow() != 0) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  }
 
  // and now we'll loop arround retrieving a batch of results, 
  // writing them to the spreadsheet and then getting the next batch etc.
  var errorflag = false;
  
  do {
    var batchWrite = [];
    // get list of course details
    // use "fields" to narrow down the size of the request to the specific fields we actually want
    var optionalArgs = {
      pageSize: pageSizeValue,
      pageToken: nextPageToken,
      fields: "nextPageToken,courses(id,name,ownerId,courseState,creationTime,updateTime,section,enrollmentCode)"
    };
    try {
      var courses = Classroom.Courses.list(optionalArgs);
      var nextPageToken = courses.nextPageToken;
      // loop round the result page
      // todo - get pageProgress for current pageToken, and continue from this point
      for ( var i= 0, len = courses.courses.length; i < len; i++) {
        var courseName =  courses.courses[i].name;
        var courseCreation = courses.courses[i].creationTime;
        var courseUpdated = courses.courses[i].updateTime;
        var courseSection = courses.courses[i].section;
        if (courseSection == null) {
            courseSection = "";
        }
        var courseCode = courses.courses[i].enrollmentCode
        var courseId = courses.courses[i].id
        var courseState = courses.courses[i].courseState;
        var ownerId = courses.courses[i].ownerId;
        
        // check if we have a "stored" ownerId
        var lookUp = ownerArray.filter(function(v,i) {
          return v[0] === ownerId;
        });
        
        var owner = "";
        var ou = "";
        var email = "";
        
        // if we get a hit through lookUp we can match the owner without an API call
        if (lookUp[0]) {
          owner = lookUp[0][1];
          email = lookUp[0][2];          
          ou = lookUp[0][3];
        } else {
          try {
            var ownerObj = AdminDirectory.Users.get(ownerId);
            owner = ownerObj.name.fullName;
            email = ownerObj.primaryEmail;            
            ou = ownerObj.orgUnitPath;
            // push result to ownerArray
            ownerArray.push([ownerId,owner,email,ou]);
          } catch(err) { // if we get an error here - the owner might be deleted!
            ou = "None";
            owner = ownerId + " (" + err.message + ")";
            ownerArray.push([ownerId,owner,email,ou]);
          }
        }
        startRow++;
        batchWrite.push([startRow,owner,email,ou,courseCreation,courseUpdated,
                         courseState,courseSection,courseName.toString(),courseCode,courseId]);
      }
      // wait until loop finishes to write to the sheet,
      // instead of writing each iteration
      var row = sheet.getLastRow()+1;
      var len = startRow + 1;
      var range = "A" + row + ":K"+len;
      sheet.getRange("L1").setValue(startRow);
      sheet.getRange(range).setValues(batchWrite);
      setBatchKey("listClasses", nextPageToken);
      } catch (e) {
        // Gradually slow things down - don't go to pageSize = 1 right away
        if (pageSizeValue > 1) {
          pageSizeValue /= 2;
          pageSizeValue = parseInt(pageSizeValue);
        }
        
        /* if the number of results don't match the page size, an error shouldn't occur,
        *  but instead a corrupt classroom could cause an error like this -
        *  indicating that your domain could have moore classsrooms that just can't be
        *  listed until google resolves some bugs:
        *  https://issuetracker.google.com/issues/36760244
        */
        // if we come down to a pageSizeValue of "1", and end up here - its an error, and we can quit!
        if (pageSizeValue == 1) {
          errorflag = true;
        }
        
        // write the pageSizeValue into the Sheet
        sheet.getRange("M1").setValue(pageSizeValue);
      }
    } while ((isTimeRunningOut("listClasses") != true) && (nextPageToken != undefined) && (errorflag != true ));
    // and do this until there are no more pages of results to fetch
  
  // empty ownerSheet and rewrite the ownerArray when we have run out of time
  ownerSheet.clear();
  var ownerLen = ownerArray.length;
  var ownerRange =  "A1:D" + ownerLen;
  if (ownerLen > ownerSheet.getMaxRows()) {
    var addRows = ownerLen - ownerSheet.getMaxRows();
    ownerSheet.insertRowsAfter(ownerSheet.getMaxRows(), addRows);
  }  
  ownerSheet.getRange(ownerRange).setValues(ownerArray);
  
  // we've run out of classrooms
  if ((nextPageToken == undefined)||(errorflag == true)){
    var emailRecipient = getEmailRecipient();
    endContinuousExecutionInstance("listClasses", emailRecipient, "Classroom");
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
    // store the individual pageProgress
    //userProperties.setProperty('GASCBL_' + fname + '_PROGRESS', 0);
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
  // uncomment to log how long each "page" takes to process:
  //Logger.log(timeElapsed);
  return (timeElapsed > 270);
}

/*
* Set the next trigger, 7 minutes in the future
*/
function enableNextTrigger_(fname) {
  var userProperties = PropertiesService.getUserProperties();
  var nextTrigger = ScriptApp.newTrigger(fname).timeBased().after(7 * 60 * 1000).create();
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
