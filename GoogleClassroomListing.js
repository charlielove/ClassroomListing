/**
 * Generates a list of all ACTIVE Google Classrooms with their creator teacher and OU path
 * Written by Charlie Love
 * Twitter: @charlie_love
 * Blog: https://charlielove.org
 **/

//function for the date
Date.prototype.yyyymmdd = function() {
    var mm = this.getMonth() + 1; // getMonth() is zero-based
    var dd = this.getDate();

    return [this.getFullYear(),
        (mm > 9 ? '' : '0') + mm,
        (dd > 9 ? '' : '0') + dd
    ].join('');
};

/**
 * Generates a list of all ACTIVE Google Classrooms with their creator teacher and OU path
 */
function listClasses() {
    //get setup. start with the date
    var now = new Date();
    //get the data as a string for the filename later
    var endTime = now.toISOString();
    //we're going to grab classroom details in batches of 500 at a time
    var pageSizeValue = 500;
    //setup an array to dump the class list into, we won't write it to the sheet until we are done as its faster to use memory
    var rows = [];
    //first page being read so this is empty - 
    var nextPageToken = '';
    //First record about a classroom is at Row 1
    var startRow = 1;
    //We're going to write the teacher and OU details into a separate array as a quick lookup of teacher details
    //(so we only have to read the each teacher only once from the Admin panel - so its faster!)
    var ownerArray = [];
    //start a loop which will keep going until all the Classrooms are listed
    do {
        //Let's setup the arguments so that we can run the query to get 500 classroom records
        var optionalArgs = {
            pageSize: pageSizeValue,
            pageToken: nextPageToken,
            fields: "nextPageToken,courses(id,name,ownerId,courseState,creationTime,updateTime,section,enrollmentCode)",
        };
        //write the results into an array object cls
        var cls = Classroom.Courses.list(optionalArgs);
        //point to the next 500 to get, nextPageToken bookmarks the next set of records to get.
        var nextPageToken = cls.nextPageToken;
        // loop round the result page
        // go through all the results in the array object to grab each classroom list and write it to an array
        for (var i = 0, len = cls.courses.length; i < len; i++) {

            //what is the state of the class, we're only interested in ACTIVE classrooms
            var courseState = cls.courses[i].courseState;
            //only write ACTIVE ones
            if (courseState = 'ACTIVE') {
                //get the classroom name, creationTime, last time it was updated and other info about it 
                var courseName = cls.courses[i].name;
                var courseCreation = cls.courses[i].creationTime;
                var courseUpdated = cls.courses[i].updateTime;
                var courseSection = cls.courses[i].section;
                if (courseSection == null) {
                    courseSection = "";
                }
                var courseCode = cls.courses[i].enrollmentCode
                var courseId = cls.courses[i].id
                
                //grab the ownerID for the classroom
                var ownerId = cls.courses[i].ownerId;

                //setup the owner variables to get the classroom owner info
                var owner = "";
                var ou = "";
                var primaryEmail = "";
              
                // check if we have a "stored" ownerId in the array of owners
                var lookUp = ownerArray.filter(function(v, i) {
                    return v[0] === ownerId;
                });

                // if we get a hit through lookUp we can match the owner without an API call
                if (lookUp[0]) {
                    owner = lookUp[0][1];
                    ou = lookUp[0][2];
                    primaryEmail = lookUp[0][3];
                } else {
                    try {
                        //grab the teacher's details from the Admin Directory and write into the owner array for a quick loop up
                        var ownerObj = AdminDirectory.Users.get(ownerId);
                        owner = ownerObj.name.fullName;
                        ou = ownerObj.orgUnitPath;
                        primaryEmail = ownerObj.primaryEmail;
                        // push result to ownerArray
                        ownerArray.push([ownerId, owner, ou, primaryEmail]);
                    } catch (err) { // if we get an error here - the owner might be deleted!
                        ou = "None";
                        owner = ownerId + " (" + err.message + ")";
                        primaryEmail = "Not Set";
                        ownerArray.push([ownerId, owner, ou, primaryEmail]);
                    }
                }
  
                //we've got all the information so write that to an array call row.
                row = [startRow, owner, primaryEmail, ou, courseCreation, courseUpdated,
                    courseState, courseSection, courseName.toString(), courseCode, courseId, ou
                ];
                //write the row to our rows array which holds all the records retrieved
                rows.push(row);
            }
        }
    } while (nextPageToken);
   
 
    //open a spreadsheet for this
    //set a name of the spreadsheet
    var my_ss = "GoogleClassroomList".concat(endTime);
    var files = DriveApp.getFilesByName(my_ss);
    var file = !files.hasNext() ? SpreadsheetApp.create(my_ss) : files.next();
    var ss = SpreadsheetApp.openById(file.getId())
    var sheet = ss.getSheets()[0];

    //get the first sheet in the workbook
    var sheet = ss.getSheets()[0];

    // setup and rename needed sheets
    // rename first sheet if not already renamed
    try {
        ss.setActiveSheet(ss.getSheetByName('ClassroomList'));
    } catch (e) {
        ss.renameActiveSheet('ClassroomList');
    }

    sheet.clear();
    var headers = ["No.", "Class Owner", "OwnerId", "Organization", "Creation Date", "Last Updated",
        "Course State", "Course Section", "Course Name", "Enrollment Code", "Course Id", "OrgUnitPath"
    ];
    sheet.appendRow(headers);

    // Append the results.
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

    Logger.log('Report spreadsheet created: %s', ss.getUrl());

}
