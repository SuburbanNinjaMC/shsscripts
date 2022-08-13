function getRows() {
  var range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getLastRow();
  return range;
}

/**
 * Randomizes assignments for Monday and Wednesday.
 * Do twice for full homework help quota.
 */
function randomizer() {
  // Note to self: adapt capability as to work for both my formatted spreadsheet and the normal type
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var tutoredAlreadyRow = spreadsheet.getRange('G:G');
  var tutoredAlreadyStudents = tutoredAlreadyRow.getValues();

  // Monday Randomizer
  var mondayRow = spreadsheet.getRangeByName('E:E');
  var mondayStudentFrees = mondayRow.getValues();
  var wednesdayRow = spreadsheet.getRangeByName('F:F');
  var wednesdayStudentFrees = wednesdayRow.getValues();
  
  var chosenMondayFrees, chosenWednesdayFrees;

  // Monday Pick
  do {
    var randomRowValue = parseInt(Math.random() * mondayStudentFrees.length);
    chosenMondayFrees = mondayStudentFrees[randomRowValue]; // gets an array
    
    Logger.log("MONDAY DEBUG: Student #" + randomRowValue + " " + chosenMondayFrees[0] + " " + tutoredAlreadyStudents[randomRowValue] + chosenMondayFrees[0].includes("7th"));

  } while (tutoredAlreadyStudents[randomRowValue] == "YES" || chosenMondayFrees[0] == "" || !chosenMondayFrees[0].includes("7th"));

  // Wednesday Pick
  do {
    var randomRowValue2 = parseInt(Math.random() * wednesdayStudentFrees.length);
    chosenWednesdayFrees = wednesdayStudentFrees[randomRowValue2];

    Logger.log("WEDNESDAY DEBUG: Student #" + randomRowValue2 + " " + chosenWednesdayFrees[0] + " " + tutoredAlreadyStudents[randomRowValue2] + chosenWednesdayFrees[0].includes("7th"));

  } while (tutoredAlreadyStudents[randomRowValue2] == "YES" || randomRowValue == randomRowValue2 || chosenWednesdayFrees[0] == "" || !chosenWednesdayFrees[0].includes("7th"));

  var names = spreadsheet.getRange('A:B').getValues();
  var mondayFinalist = names[randomRowValue][0];
  var wednesdayFinalist = names[randomRowValue2][0];
  var mondayEmail = names[randomRowValue][1];
  var wednesdayEmail = names[randomRowValue2][1];
  Logger.log("RANDOMIZED PICKS:\nMonday - " + mondayFinalist + ", " + mondayEmail + "\nWednesday - " + wednesdayFinalist + ", " + wednesdayEmail);

  // Alter Spreadsheet

  spreadsheet.getRange("G" + (randomRowValue + 1) + "").setValue("YES");
  spreadsheet.getRange("G" + (randomRowValue2 + 1) + "").setValue("YES");

  var emailArray = [mondayFinalist, wednesdayFinalist, mondayEmail, wednesdayEmail];
  var officerArray = ["jng22@scarsdaleschools.org", "ehersch22@scarsdaleschools.org", "swong22@scarsdaleschools.org"];
  /** 
  sendEmail(emailArray);
  sendEmailConfirmationToMe(emailArray, officerArray);
  */
}

function sendEmail(recipientArray) {
  var daysOfWeek = ["Sunday", "Monday", "Tuesday", "Wednesday","Thursday", "Friday", "Saturday"];
  var currentDate = new Date();

  // set i = 2 to get wherever the email is.
  for (i = 2; i < recipientArray.length; i++) {
    var d = new Date();
    d.setDate(d.getDate() + ((((2 * i - 3) + 7 - d.getDay()) % 7) || 7));

    var day = d.getDate();
    var month = d.getMonth() + 1;
    var year = d.getFullYear();
    var dayOfWeekForDay = daysOfWeek[d.getDay()];
    var assignedDate = dayOfWeekForDay + ", " + month + "/" + day + "/" + year;

    var recipient = recipientArray[i];
    var firstNameOfRecipient = recipientArray[i - 2].substring(0, recipientArray[i - 2].indexOf(" "));
    var subject = "Signifer Homework Help Center Assignment";
    var message = "Dear " + firstNameOfRecipient + ",\nWe hope all is well! We just want to let you know that you were randomly selected to work the Homework Help Center on " + assignedDate  + " during 7th period in the library. This center will be run in the reference room in the library (first floor next to the desk). You are required to sit here during the period and help students with homework, studying, etc. Please let us know if you're able to make it! \n\nThanks, \nSignifer Officers\n\n\n" + currentDate; 
    MailApp.sendEmail(recipient, subject, message);
  }
}

function sendEmailConfirmationToMe(emailArray, officerArray) {
  var currentDate = new Date();
  for (i = 0; i < officerArray.length; i++) {
    var recipient = officerArray[i];
    var subject = "CONFIRMATION: Signifer Homework Help Center Assignment";
    var message = emailArray[0] + " and " + emailArray[1] + " have been assigned to the Homework Help Center on " + currentDate + " on Monday and Wednesday of this week respectively.";
    MailApp.sendEmail(recipient, subject, message);
  }
}
