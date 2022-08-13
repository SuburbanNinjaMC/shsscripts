/**
 * PREREQUISITES: Make sure that the Original Sign Up Genius sheet is loaded up FIRST. THEN, upload the mentor sheet. THEN, the registration sheet.
 * WHEN CREATING THE SIGNUPGENIUS REPORT, CHECK "SHOW SELECT FIELDS" AND UNCHECK "TIMESTAMP"
 */
function createTheSpreadsheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var signUpSheet = spreadsheet.getSheets()[0]; // first sheet
  var mentorSheet = spreadsheet.getSheets()[1]; // second sheet
  var registrationSheet = spreadsheet.getSheets()[2]; // third sheet
  SpreadsheetApp.setActiveSheet(registrationSheet); // temp placeholder

  if (signUpSheet.getRange("A1").getValue() === "Sign Up") {
    signUpSheet.deleteRow(1);
  }

  var sheetPlaceholder = spreadsheet.getSheetByName("Final Mentor-Student Assignments");
  if (sheetPlaceholder != null) {
    spreadsheet.deleteSheet(sheetPlaceholder);
  }
  
  // Creates the Result Spreadsheet without us doing anything
  spreadsheet.insertSheet("Final Mentor-Student Assignments");
  var resultSheet = spreadsheet.getSheets()[3];
  SpreadsheetApp.setActiveSheet(resultSheet);

  // ======== Formatting Instruments ====================

  var SULastRow = signUpSheet.getLastRow() - 2; // last row is just "sign ups with no time zone"
  var INSTRUMENT_COLUMN = 1;
  var originalinstrumentColumn = signUpSheet.getRange(1, INSTRUMENT_COLUMN, SULastRow);
  var instrumentColumnValues = originalinstrumentColumn.getValues();

  var row; // loop counter

  for (row = 0; row < instrumentColumnValues.length; row++) {
    instrumentColumnValues[row][0] = instrumentColumnValues[row][0].substring(instrumentColumnValues[row][0].indexOf("-") + 2); // remove "Practice Partners" from beginning.
  }

  var formattedInstrumentColumn = resultSheet.getRange(1, INSTRUMENT_COLUMN, SULastRow);
  formattedInstrumentColumn.setValues(instrumentColumnValues); // Col 1

  // ======== Formatting Times =====================

  var START_TIME_COLUMN = 2;
  var originalStartTimeColumn = signUpSheet.getRange(1, START_TIME_COLUMN, SULastRow, 4); // the 4 is to extend length of array
  var timeValues = originalStartTimeColumn.getValues();

  for (row = 0; row < timeValues.length; row++) {
    var retrievedDate = new Date(timeValues[row][0]);
    retrievedDate.setHours(retrievedDate.getHours() - 3); // for some reason it's in Pacific Time...
    var retrievedProperDate = (retrievedDate.getMonth() + 1) + "/" + retrievedDate.getDate() + "/" + retrievedDate.getFullYear();
    var retrievedTime = retrievedDate.toLocaleTimeString(['en-US'], {hour: '2-digit', minute:'2-digit'});

    timeValues[row][0] = retrievedProperDate;

    var daysOfWeek = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    timeValues[row][1] = daysOfWeek[retrievedDate.getDay()];

    // Start and End Times
    timeValues[row][2] = retrievedTime;
    retrievedDate.setMinutes(retrievedDate.getMinutes() + 20); // end time
    timeValues[row][3] = retrievedDate.toLocaleTimeString(['en-US'], {hour: '2-digit', minute:'2-digit'});
  }

  var formattedStartTimeAndDateColumn = resultSheet.getRange(1, START_TIME_COLUMN, SULastRow, 4);
  formattedStartTimeAndDateColumn.setValues(timeValues); // Col 2 = date, 3 = time, 4 = end time

  // ========= Formatting Parent Names, Student Names, and Emails =========
  var PARENT_FIRST_NAME_COL = 6;
  var originalParentInfoColumn = signUpSheet.getRange(1, PARENT_FIRST_NAME_COL, SULastRow, 6);
  var originalParentValues = originalParentInfoColumn.getValues();

  for (row = 0; row < originalParentValues.length; row++) {

    // Capitalize First and Last Names of Parents and Students
    originalParentValues[row][0] = originalParentValues[row][0].charAt(0).toUpperCase() + originalParentValues[row][0].substring(1);
    originalParentValues[row][1] = originalParentValues[row][1].charAt(0).toUpperCase() + originalParentValues[row][1].substring(1);

    originalParentValues[row][4] = originalParentValues[row][4].charAt(0).toUpperCase() + originalParentValues[row][4].substring(1);
    if (originalParentValues[row][4].includes(" ")) {
      var tsfn = originalParentValues[row][4].substring(0, originalParentValues[row][4].indexOf(" "));
      var tsln = originalParentValues[row][4].substring(originalParentValues[row][4].indexOf(" ") + 1, originalParentValues[row][4].length);

      while (tsln.charAt(0) == " ") {
        tsln = tsln.substring(1);
      }

      tsln = tsln.charAt(0).toUpperCase() + tsln.substring(1);
      originalParentValues[row][4] = tsfn  + " " + tsln;
    }

    var fullParentName = originalParentValues[row][0] + " " + originalParentValues[row][1];
    //                   First Name                           Last Name

    originalParentValues[row][0] = fullParentName;
    originalParentValues[row][1] = originalParentValues[row][2]; // shift the emails over by one column
    originalParentValues[row][2] = originalParentValues[row][4]; // shift Student Name into the column next to parent email
    originalParentValues[row][3] = "";
    originalParentValues[row][4] = "";
  }

  var formattedParentColumn = resultSheet.getRange(1, PARENT_FIRST_NAME_COL, SULastRow, 6);
  
  // ======== Getting Student Grades, Proficiencies, and Goals, and Parent Formatting ===========
  var REGISTRATION_PARENT_INFO_BEGIN_COL = 2;
  var STUDENT_NAME_COLUMN = 10;
  var REGISTRATION_STUDENT_NAME_COLUMN = 6;
  var REGISTRATION_STUDENT_GRADE_COLUMN = 7;
  var REGISTRATION_STUDENT_PROFICIENCY_COLUMN = 15;
  var REGISTRATION_STUDENT_GOALS_COLUMN = 16;

  var registrationSheetValues = registrationSheet.getDataRange().getValues();
  var registrationRow; // loop counter
  
  var finalGradeColumns = resultSheet.getRange(1, 9, SULastRow, 3);
  var finalGradeValues = finalGradeColumns.getValues();
  for (row = 0; row < originalParentValues.length; row++) {
    var studentSearchName = originalParentValues[row][2];

    var firstStudentSearchName = studentSearchName.toLowerCase();
    if (studentSearchName.indexOf(" ") != -1) {
      firstStudentSearchName = studentSearchName.substring(0, studentSearchName.indexOf(" ")).toLowerCase();
    }

    fullParentName = originalParentValues[row][0];
    fullParentEmail = originalParentValues[row][1];
    fullParentName = fullParentName.toLowerCase();
    fullParentEmail = fullParentEmail.toLowerCase();

    for (registrationRow = 1; registrationRow < registrationSheetValues.length; registrationRow++) {
      var tempName = registrationSheetValues[registrationRow][REGISTRATION_STUDENT_NAME_COLUMN - 1];
      var exactMatch = (tempName.indexOf(studentSearchName) != -1);// Easy case: searched student name is the one in preregistration sheet

      // Harder case: check the first names of the kids and then check their parents. If they match up, match them.
      var firstName = tempName.toLowerCase();
      if (tempName.indexOf(" ") != -1) {
        firstName = tempName.substring(0, tempName.indexOf(" ")).toLowerCase();
      } 

      var registrationParent = registrationSheetValues[registrationRow][REGISTRATION_PARENT_INFO_BEGIN_COL - 1] + " " + registrationSheetValues[registrationRow][REGISTRATION_PARENT_INFO_BEGIN_COL]; // parent first and last
      var registrationParentName = registrationParent.toLowerCase();
      
      var registrationParentEmail = registrationSheetValues[registrationRow][REGISTRATION_PARENT_INFO_BEGIN_COL + 1];
      registrationParentEmail = registrationParentEmail.toLowerCase();

      var almostExactMatch = ((registrationParentName.indexOf(fullParentName) != -1 || registrationParentEmail.indexOf(fullParentEmail) != -1) && firstStudentSearchName.indexOf(firstName) != -1);

      if (exactMatch || almostExactMatch) { // Check if the times align!
        var grade = registrationSheetValues[registrationRow][REGISTRATION_STUDENT_GRADE_COLUMN - 1];
        var proficiency = registrationSheetValues[registrationRow][REGISTRATION_STUDENT_PROFICIENCY_COLUMN - 1];
        var goals = registrationSheetValues[registrationRow][REGISTRATION_STUDENT_GOALS_COLUMN - 1];

        finalGradeValues[row][0] = grade;
        finalGradeValues[row][1] = proficiency;
        finalGradeValues[row][2] = goals;
      }
    }

    // Also, while we're here, let's check if the parent inputted their student's name for parent name.
    if (fullParentName.includes(firstStudentSearchName) && !(firstStudentSearchName === "")) {
      for (registrationRow = 1; registrationRow < registrationSheetValues.length; registrationRow++) {
        // Look for the student name and get email
        if (registrationSheetValues[registrationRow][REGISTRATION_STUDENT_NAME_COLUMN - 1].includes(studentSearchName) && registrationSheetValues[registrationRow][REGISTRATION_PARENT_INFO_BEGIN_COL + 1].toLowerCase() === fullParentEmail) {
          originalParentValues[row][0] = registrationSheetValues[registrationRow][REGISTRATION_PARENT_INFO_BEGIN_COL - 1] + " " + registrationSheetValues[registrationRow][REGISTRATION_PARENT_INFO_BEGIN_COL];
        }
      }
    }
  }
  formattedParentColumn.setValues(originalParentValues); // Col 5 = parent name, col 6 = parent email, col 7 = student
  finalGradeColumns.setValues(finalGradeValues); // col 8 - grade, col 9 - prof, col 10 - goals

  // FUN PART: MENTOR-STUDENT ASSIGNMENTS

  var beganTime = new Date();
  
  var MENTOR_NAME_COL = 1;
  var MENTOR_INSTRUMENT_COL = 4;
  var MENTOR_PRIMARY_DAY_COL = 5;
  var MENTOR_START_TIMES_COL = 6;
  var MENTOR_SECONDARY_DAY_COL = 7;
  var MENTOR_SECONDARY_START_TIMES_COL = 8;

  var finalMentorColumn = resultSheet.getRange(1, 12, SULastRow, 3);
  var finalMentorValues = finalMentorColumn.getValues();

  var mentorInformation = mentorSheet.getDataRange().getValues();

  var mentorTimesMentored = [[""]];

  var mentorRow;
  for (row = 0; row < originalParentValues.length; row++) {
    var truncatedTime = timeValues[row][2]; // Timestamp
    truncatedTime = truncatedTime.substring(1, truncatedTime.indexOf(" "));

    var studentInstrument = formattedInstrumentColumn.getValues()[row][0];
    var studentWeekday = formattedStartTimeAndDateColumn.getValues()[row][1];

    var allStudentInformation = [studentInstrument, studentWeekday, truncatedTime]; // ["Instrument", "Weekday", "Start Time"]

    var onlyOne = function(instrument, day, time) {
      var totalPossibleMentors = 0;
  
      var mentorInformation = mentorSheet.getDataRange().getValues();

      for (testRow = 0; testRow < mentorInformation.length; testRow++) {
        var testMentorInstrument = mentorInformation[testRow][MENTOR_INSTRUMENT_COL - 1];
        var testMentorDay = mentorInformation[testRow][MENTOR_PRIMARY_DAY_COL - 1];
        var testMentorTime = "" + mentorInformation[testRow][MENTOR_START_TIMES_COL - 1];

        if (testMentorInstrument.includes(instrument) && testMentorDay.includes(day) && testMentorTime.includes(time)) {
          totalPossibleMentors++;
        }

        var testMentorSecondaryDay = mentorInformation[testRow][MENTOR_SECONDARY_DAY_COL - 1];
        var testMentorSecondaryTime = "" + mentorInformation[testRow][MENTOR_SECONDARY_START_TIMES_COL - 1];

        if (testMentorInstrument.includes(instrument) && testMentorSecondaryDay.includes(day) && testMentorSecondaryTime.includes(time)) {
          totalPossibleMentors++;
        }
      }

      if (totalPossibleMentors <= 1) {
        return true;
      }

      return false;
    }

    var counter = 0;
    for (mentorRow = 0; mentorRow < mentorInformation.length; mentorRow++) {
      var allMentorPrimaryInformation = [mentorInformation[mentorRow][MENTOR_INSTRUMENT_COL - 1], mentorInformation[mentorRow][MENTOR_PRIMARY_DAY_COL - 1], mentorInformation[mentorRow][MENTOR_START_TIMES_COL - 1]];
      var allMentorSecondaryInformation = [mentorInformation[mentorRow][MENTOR_INSTRUMENT_COL - 1], mentorInformation[mentorRow][MENTOR_SECONDARY_DAY_COL - 1], mentorInformation[mentorRow][MENTOR_SECONDARY_START_TIMES_COL - 1]];

      var doPrimaryDatesAlign = (allStudentInformation[0].indexOf(allMentorPrimaryInformation[0]) != -1 && allStudentInformation[1].indexOf(allMentorPrimaryInformation[1]) != -1 && allMentorPrimaryInformation[2].indexOf(allStudentInformation[2]) != -1);

      var doSecondaryDatesAlign = (allStudentInformation[0].indexOf(allMentorSecondaryInformation[0]) != -1 && allStudentInformation[1].indexOf(allMentorSecondaryInformation[1]) != -1 && allMentorSecondaryInformation[1] != "" && allMentorSecondaryInformation[2] != "");

      studentSearchName = formattedParentColumn.getValues()[row][2];

      var mentorName = mentorInformation[mentorRow][MENTOR_NAME_COL - 1];
      var mentorGrade = mentorInformation[mentorRow][MENTOR_NAME_COL];
      var mentorEmail = mentorInformation[mentorRow][MENTOR_NAME_COL + 1];

      var mentorPresent = function() {
          for (var testRow = 0; testRow < mentorTimesMentored.length; testRow++) {
            if (mentorTimesMentored[testRow][0] === mentorName) {
                return true;
            }
          }
          return false;
        };

      var isMentorAlreadyHere = mentorPresent();

      var numTimes = 0;
      for (var r = 0; r < mentorTimesMentored.length; r++) {
        if (mentorTimesMentored[r].indexOf(mentorName) != -1) {
          numTimes = mentorTimesMentored[r][2];
        }
      }

      var align = [doPrimaryDatesAlign, doSecondaryDatesAlign];

      // Dates Align
      for (var i = 0; i < 2; i++) { // just repeat this 2 times
        if (align[i] && finalMentorValues[row][0] === "") {
          // If you've mentored less than 3 different students or if you are the only person available for a time
          var conditions = [(numTimes <= 3 || onlyOne(studentInstrument, studentWeekday, truncatedTime)), true];
          if (conditions[counter] && finalMentorValues[row][0] === "") {
            finalMentorValues[row][0] = mentorName; // Mentor Name
            finalMentorValues[row][1] = mentorGrade; // Grade
            finalMentorValues[row][2] = mentorEmail; // Email;

            if (!isMentorAlreadyHere && mentorTimesMentored[mentorTimesMentored.length - 1][0] == "") {
              mentorTimesMentored[mentorTimesMentored.length - 1][0] = mentorName;
              mentorTimesMentored[mentorTimesMentored.length - 1].push([studentSearchName]);
              mentorTimesMentored[mentorTimesMentored.length - 1].push(1);

              mentorTimesMentored.push([""]); // add a new space for future mentors
              isMentorAlreadyHere = mentorPresent();
            }

            for (var mentorTimesRow = 0; mentorTimesRow < mentorTimesMentored.length; mentorTimesRow++) {
              if (mentorTimesMentored[mentorTimesRow].length > 1 && mentorName == mentorTimesMentored[mentorTimesRow][0] && !mentorTimesMentored[mentorTimesRow][1].includes(studentSearchName)) {
                mentorTimesMentored[mentorTimesRow][1].push(studentSearchName);
                mentorTimesMentored[mentorTimesRow][2]++;
                break;
              }
            }
          }
        }
      }

      // Basically, what we just did filled out the slots for everyone that fulfilled the criteria. Next is picking up the scraps -- filling out who was left behind.
      if (mentorRow == mentorInformation.length - 1 && counter == 0) {
        counter++;
        mentorRow = -1;
      }
    }
  }
  finalMentorColumn.setValues(finalMentorValues); // col 11 - mentor name, col 12 - grade, col 13 - email
  
  // Final Checks - for ANY BLANKS
  var allResultValues = resultSheet.getRange(1, 1, resultSheet.getLastRow(), 13).getValues();
  for (row = 0; row < allResultValues.length; row++) {
    for (var column = 0; column < allResultValues[row].length; column++) {
      if (allResultValues[row][column] === "") {
        Logger.log("There is information missing for " + allResultValues[row][7] + " for the session given on " + allResultValues[row][1].toLocaleDateString() + " at " + allResultValues[row][3].toLocaleTimeString(['en-US']) + ".");
      }
    }
  }
  
  // =========== Reorder Columns and Format Nicely (FINAL PART) ==================

  var endedTime = new Date();
  var elapsedTime = endedTime.getTime() - beganTime.getTime();
  Logger.log("Mentor assignments took " + elapsedTime + " ms, which is " + (elapsedTime/1000) + " sec");

  rearrangeColumns(spreadsheet.getActiveSheet(), [0, 1, 2, 3, 4, 11, 12, 13, 7, 8, 5, 6, 9, 13]);
  resultSheet.insertRowBefore(1);
  var titleRow = resultSheet.getRange(1, 1, 1, 14).setFontWeight("bold");
  var titleValues = [["Instrument", "Date", "Weekday", "Start Time", "End Time", "Mentor Name", "Mentor Grade", "Mentor Email", "Student Name", "Student Grade", "Parent Name", "Parent Email", "Proficiency", "Goals"]];
  titleRow.setValues(titleValues);

  spreadsheet.getActiveSheet().autoResizeColumns(1, 15);

}
