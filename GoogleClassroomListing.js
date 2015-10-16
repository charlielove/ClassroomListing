function listClasses(){

  //by Charlie Love
  //charlielove.org tw: @charlie_love
  
  //open a new spreadsheet
  var my_ss = "Classroom Listing";
  var files = DriveApp.getFilesByName(my_ss);
  var file = !files.hasNext() ? SpreadsheetApp.create(my_ss) : files.next();
  var ss = SpreadsheetApp.openById(file.getId())
  try 
  {
     ss.setActiveSheet(ss.getSheetByName(my_sheet));
  } catch (e){;} 
  var sheet = ss.getActiveSheet();
  sheet.clear();
  sheet.appendRow(["No.","Class Owner","Organization","Creation Date","Course State","Course Section","Course Name"]);
  //start at row 0
  var startRow = 0;

  //let's get the first page of results for classroom listings 
  //(there is a limit to how many you can get at one go!)
  var nextPageToken = ""; 
  
  //and now we'll loop arround retrieving a batch of results, 
  //writing them to the spreadsheet and then getting the next batch etc.
  do {
    //get list of course details
    var optionalArgs = {
      pageSize: 400,
      pageToken: nextPageToken
    };
    var courses = Classroom.Courses.list(optionalArgs);
    var nextPageToken = courses.nextPageToken;
  
    //loop round
    for ( var i= 0, len = courses.courses.length; i < len; i++) {
      var courseName =  courses.courses[i].name;
      var courseCreation = courses.courses[i].creationTime;
      var courseUpdated = courses.courses[i].updateTime;
      var courseSection = courses.courses[i].section;
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
    
        ss.getSheets()[0].appendRow([startRow+1,owner,ou,courseCreation,courseState,courseSection, courseName.toString()]);
        startRow++;  //we've written a row, so add one to start row.
      }

  } while (nextPageToken != undefined); //and do this until there are no more pages of results to get

}

