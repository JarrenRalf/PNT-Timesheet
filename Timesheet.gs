/**
* This onEdit function trigger ensures that the user makes valid changes to the start time, end time and break time ranges.
*
* @param {Object} e Event object.
*/
function onEdit(e)
{
  if (!e)
    throw new Error('Please do not run the script in the script editor window. It runs automatically when you edit the spreadsheet.');
  
  const START_TIME_COL =  4;
  const   END_TIME_COL =  5;
  const      BREAK_COL =  6;
  const MIN_BREAK_TIME = 30;
  const NUM_HEADER_ROWS = 10;
  const   EMPLOYEE_SIGNATURE_COL =  7;
  const SUPERVISOR_SIGNATURE_ROW = 27;
  
  var spreadsheet = e.source;      // A Spreadsheet object, representing the Google Sheets file to which the script is bound.
  var    oldValue = e.oldValue;    // Cell value prior to the edit, if any. Only available if the edited range is a single cell. Will be undefined if the cell had no previous content.
  var       value = e.value;       // The value of the edited cell
  var       range = e.range;       // A Range object, representing the cell or range of cells that were edited.
  var ui = SpreadsheetApp.getUi(); // A User Interface object
  var okButton = ui.ButtonSet.OK;  // The OK-button for the user interface
  var row    = range.rowStart;     // The first row of the edited range
  var col    = range.columnStart;  // The first column of the edited range
  var rowEnd = range.rowEnd;       // The last row of the edited range
  var colEnd = range.columnEnd;    // The last column of the edited range
  var sheet = spreadsheet.getActiveSheet();
  var dateAndTime = sheet.getRange(row, 2, 1, 4).getValues();
  var       day = dateAndTime[0][0];
  var   weekDay = dateAndTime[0][1];
  var startTime = dateAndTime[0][2];
  var   endTime = dateAndTime[0][3];
  
  if (row == rowEnd && row > NUM_HEADER_ROWS) // Ensure a single row is being edited
  {
    if (isColumnEdited(BREAK_COL, col, colEnd)) // The break column is edited 
    {
      if (isBlank(startTime) || isBlank(endTime)) // User is editing a break time range they didnt work
      { 
        if (isBlank(day) || isBlank(weekDay)) // This means this is the 16th day in a 15 day pay period
          ui.alert("Timesheet Mistake", "This pay period doesn't contain a 16th day.", ui.ButtonSet.OK);
        else
          ui.alert("Timesheet Mistake", "You didn\'t work on " + weekDay + " the " + day + getSuffix(day) + ".", okButton);
        
        undoEdit(oldValue, range); 
      }
      else if (!isBreakTimeValid(MIN_BREAK_TIME, value))
      {
        ui.alert("Invalid Input", "For break time, please enter a whole number that is a multiple of " + MIN_BREAK_TIME + " (Minutes).", okButton);
        undoEdit(oldValue, range);
      }
    }
    else if (isColumnEdited(START_TIME_COL, col, colEnd) || isColumnEdited(END_TIME_COL, col, colEnd)) // The start OR end time column is edited
    {
      if (row == SUPERVISOR_SIGNATURE_ROW)
      {
        ui.alert("Invalid Input", "Please don't edit the signature range.", okButton);
        undoEdit(oldValue, range);
      }
      else if (hasContent(value) && startTime >= endTime)
      {
        var timeRange = sheet.getRange(row, START_TIME_COL, 1, 2);
        ui.alert("Timesheet Mistake", "The end time must be later than the start time.", okButton);
        timeRange.setNumberFormats([["h:mm am/pm", "h:mm am/pm"]]); // Reset the number formats
        timeRange.setValues([["9:00 AM", "5:00 PM"]]);              // Change the start time and end time back to default
      }
    }
    else if (isColumnEdited(EMPLOYEE_SIGNATURE_COL, col, colEnd)) // The signature / hours column is edited
    {
      ui.alert("Invalid Input", "Please don't edit the signature range.", okButton);
      undoEdit(oldValue, range);
    }
  }
}

/**
* This function places the supervisor's signature on the timesheet, representing approval of the employees hours. 
* After which it sends the hours approval email, then the trigger to send the unnapproved timesheet email is deleted.
*
* @author Jarren Ralf
*/
function approveHours()
{
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActive();
  var timesheet = spreadsheet.getSheetByName('Timesheet');
  
  // If the employee has not clicked on the "Get Approval" button, then their signature is not on the timesheet, and thus haven't endorsed their own hours
  if (!hasEmployeeSigned(timesheet))
    throw new Error('Timesheet can\'t be approved until the employee has signed it.');
  
  var response = ui.alert('Give Approval of Hours Worked', 'You are about to send an email to the payroll manager '+
                          'authorizing the hours that your staff member has worked. ' +
                          'Do you want to proceed?', ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.YES)
  {
    var signatureRange = timesheet.getRange(27, 4);
    var isVacationPayRequested = checkVacationPay(timesheet);

    signatureRange.setFormula("=SupervisorSignature"); // Set the formula to import the supervisor signature
    signatureRange.activate();                         // Take the supervisor to the timesheet (so they can observe their signature)
    (isVacationPayRequested[0]) ? hoursApprovalEmail('HoursApproval_withVacationPay', spreadsheet) : hoursApprovalEmail('HoursApproval', spreadsheet); // Send email
  }
}

/**
* This function checks if the employee's triggers are created by checking if the trigger array is empty or not.
*
* @param  {Object[]} arr The given trigger array
* @return {Boolean}  Whether the array is empty or not
* @author Jarren Ralf
*/
function areEmployeesTriggersOn(arr)
{
  return arr.length != 0;
}

/**
* This function sets up the the reminder email trigger when no arguments are sent to it.
* When the function has received an argument, the unapproved email trigger is initiated.
*
* @author Jarren Ralf
*/
function autoSendEmail()
{
  var year, month, startDay, endDay, today, emailDay, reminderDay;

  // Determine which pay period we are in and set the appropriate values
  [year, month, startDay, endDay,, today] = determinePayPeriod();
  [,, emailDay, reminderDay] = generateDates(year, month + 1, startDay, endDay);
  emailDay = emailDay.substring(3, 5);       // Extract only the day from the string
  reminderDay = reminderDay.substring(3, 5); 
  
  // 1 hour before the start of the first day of the pay period. Only the auto trigger should be accessing the sheet before this time
  var triggerControlDate = new Date(year, month, startDay, 8);

  if (today < triggerControlDate) // If the trigger is running this function, reset the missing protections and set up the reminder email trigger
  {
    setProtections(); // Reset the missing protections
    ScriptApp.newTrigger("reminderEmail").timeBased().onMonthDay(reminderDay).atHour(9).create();
  }
  else if (arguments[0]) // Vacation pay
    ScriptApp.newTrigger("unapprovedTimesheet_withVacationPayEmail").timeBased().onMonthDay(emailDay).atHour(9).create();
  else if (arguments.length >= 1) // Otherwise, NO vacation pay
    ScriptApp.newTrigger("unapprovedTimesheetEmail").timeBased().onMonthDay(emailDay).atHour(9).create();
  else // If someone runs the function manually, this will turn on the reminder email trigger
    ScriptApp.newTrigger("reminderEmail").timeBased().onMonthDay(reminderDay).atHour(9).create();
}

/**
* Calculates Easter in the Gregorian/Western (Catholic and Protestant) calendar 
* based on the algorithm by Oudin (1940) from http://www.tondering.dk/claus/cal/easter.php
* @returns {array} [int month, int day]
*/
function calculateGoodFriday(year)
{
	var f = Math.floor,
		// Golden Number - 1
		G = year % 19,
		C = f(year / 100),
		// related to Epact
		H = (C - f(C / 4) - f((8 * C + 13)/25) + 19 * G + 15) % 30,
		// number of days from 21 March to the Paschal full moon
		I = H - f(H/28) * (1 - f(29/(H + 1)) * f((21-G)/11)),
		// weekday for the Paschal full moon
		J = (year + f(year / 4) + I + 2 - C + f(C / 4)) % 7,
		// number of days from 21 March to the Sunday on or before the Paschal full moon
		L = I - J,
		month = 3 + f((L + 40)/44),
		day = L + 28 - 31 * f(month / 4) - 2;
  
    // If the day is negative, make the appropriate changes to the values of month and day
    if (day < 0) 
    {
      month--;
      day = 31 + day
    }

	return [month - 1, day];
}

/**
* This function checks if all of the hours worked are valid, by checking for negative numbers, more than 24 hours worked in a day, and
* for the hours being a multiple of 0.50. It also checks for any negative break times.
*
* @param {Object[][]} hours A double array containing the hours worked along with break time and some blanks
* @return {Boolean} [areAllHoursValid, areAllBreakTimesValid] Returns true if all the hours are valid and false if atleast one is invalid
* @author Jarren Ralf
*/
function checkHours(hours)
{
  var areAllHoursValid = true, areAllBreakTimesValid = true;
  
  for (var i = 0; i < hours.length; i++)
  {

    if (!Number.isInteger(Math.round(100*hours[i][1])/50) || hours[i][1] < 0 || hours[i][1] >= 24) // One of the hours is invalid
    {
      areAllHoursValid = false;
      break;
    }
    
    if (hours[i][0] < 0) // One of the break times is invalid
    {
      areAllBreakTimesValid = false;
      break;
    }
  }
  return [areAllHoursValid, areAllBreakTimesValid];
}

/**
* This function checks whether the employee has requested vacation pay to be added to their upcoming paycheck.
* It also checks if both vacation pay options have been selected.
*
* @return {Boolean[]} [isVacationPayRequested, areBothOptionsSelected] Whether vacation pay is requested, Whether both options are selected
* @author Jarren Ralf
*/
function checkVacationPay(timesheet)
{
  var values = timesheet.getRange(8, 5, 1, 3).getValues();
  var isFullAmountRequested = values[0][0]; // A boolean representing whether the full amount of vacation pay is requested or not
  var          customAmount = values[0][2]; // A value representing a custom amount of vacation pay 
  var isVacationPayRequested, areBothOptionsSelected;
  
  // Throw an error if the user has not entered a valid quantity for vacation pay 
  if (!isBlank(customAmount) && !isNumber(customAmount))
    throw new Error('Vacation Pay Requested must be a valid quantity.');
  
  // Set the boolean variables appropriately
  if (!isFullAmountRequested && isBlank(customAmount))
  {
    isVacationPayRequested = false;
    areBothOptionsSelected = false;
  }
  else if (isFullAmountRequested && !isBlank(customAmount))
  {
    isVacationPayRequested = true;
    areBothOptionsSelected = true;
  }
  else if (isFullAmountRequested || !isBlank(customAmount))
  {
    isVacationPayRequested = true;
    areBothOptionsSelected = false;
  }

  return [isVacationPayRequested, areBothOptionsSelected];
}

/**
* This is a function that ouputs a string for the date in the format MM/DD/YYYY.
* It also prints a '0' infront of single digit months and days for consistent formatting.
*
* @param  {Number} month The chosen month
* @param  {Number} day   The chosen day
* @param  {Number} year  The chosen year
* @return {String} date  A string that represents the date
* @author Jarren Ralf
*/
function createDateString(month, day, year)
{ 
  var date  = (month < 10) ? '0' + month.toString() + '/'                   : month.toString() + '/';
      date += ( day  < 10) ? '0' +   day.toString() + '/' + year.toString() :   day.toString() + '/' + year.toString();
  
  return date;
}

/**
* This function deactivates the spreadsheet by deleting all of the triggers. 
* This means no more automatic emails will be sent and the hours will not reset each pay period.
*
* @author Jarren Ralf
*/
function deactivateTimesheet()
{
  var ui = SpreadsheetApp.getUi();
  var firstResponse = ui.alert('Deactivate Timesheet', 'You are about to delete all of the ' +
                               'triggers that make this spreadsheet function properly. '     +
                               'Do you want to proceed?', ui.ButtonSet.YES_NO);
  
  if (firstResponse == ui.Button.YES)
  {
    var secondResponse = ui.alert('Are you sure?', ui.ButtonSet.YES_NO);
    
    if (secondResponse == ui.Button.YES)
    {
      updateStatusMessage(false);
      deleteAllTriggers();
    }
  }
}

/**
 * Deletes the reminder email trigger.
 *
 * @author Jarren Ralf
 */
function deleteReminderEmailTrigger()
{
  removeProtections(); // Removes certain protections so that the employee can run the setDaysOfPayPeriod function on a trigger
  deleteTriggers(false, "reminderEmail");
}

/**
 * Deletes all of the triggers
 *
 * @author Jarren Ralf
 */
function deleteAllTriggers()
{
  deleteTriggers(true);
}

/**
 * This function either deletes all of the triggers associated with a given trigger handle or all of the project triggers.
 *
 * @param {Boolean} isDeactivated If someone clicked on the deactivation button or not
 * @param {String}  triggerHandle The name of the function to delete the associated triggers for
 * @author Jarren Ralf
 */
function deleteTriggers(isDeactivated, triggerHandle)
{
  var triggers = ScriptApp.getProjectTriggers(); // Get all of the triggers in the current project
  
  if (isDeactivated)
  {
    for (var i = 0; i < triggers.length; i++)
      ScriptApp.deleteTrigger(triggers[i]);
  }
  else
  {
    for (var i = 0; i < triggers.length; i++)
    {
      if (triggers[i].getHandlerFunction() == triggerHandle)
        ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Deletes the unapproved timesheet email trigger.
 *
 * @author Jarren Ralf
 */
function deleteUnapprovedTimesheetEmailTrigger()
{
  deleteTriggers(false, "unapprovedTimesheetEmail");
  deleteTriggers(false, "unapprovedTimesheet_withVacationPayEmail");
}

/**
* This function determines which is the current Pay Period, based on the current year, month and day.
* The period is either 1-15 or 16-[end of month]. 
*
* @return {[...Number[], Date]} An array containing the year, month, startDay, endDay, firstDayOfWeek, and today related to the current Pay Period
* @author Jarren Ralf
*/
function determinePayPeriod()
{
  var today = new Date();           // A date object representing today's date
  var currentDay = today.getDate(); // An integer corresponding to today's day (from 1 - 31)
  var year = today.getFullYear();   // An integer representing the current year
  var month = today.getMonth();     // An integer between 0-11 representing the month
  var startDay, endDay, firstDayOfWeek;
  
  const START_SECOND_PAY_PERIOD = 16;

  //Check which pay period we are in, (either 1-15 or 16-[end of month]) and then set the appropriate values
  (currentDay < START_SECOND_PAY_PERIOD) ? (startDay = 1, endDay = 15) : (startDay = START_SECOND_PAY_PERIOD, endDay = getDaysInMonth(month, year));
  firstDayOfWeek = new Date(year, month, startDay).getDay(); // Returns 0-6 mapping to sun-sat
  
  return [year, month, startDay, endDay, firstDayOfWeek, today];
}

/**
* This is a function that generates the payPeriod, payDay, timesheet-emailDay, and reminder-emailDay for a 
* given year, month, and pay period start and end days.
*
* @param  {number} year     The chosen year
* @param  {number} month    The chosen month
* @param  {number} startDay The start day of the period
* @param  {number} endDay   The end day of the period
* @return {String[]} An array containing the payPeriod, payDay, emailDay, and reminderDay
* @author Jarren Ralf
*/
function generateDates(year, month, startDay, endDay)
{  
  const    SUNDAY =  0;
  const    MONDAY =  1;
  const   TUESDAY =  2;
  const  THURSDAY =  4;
  const    FRIDAY =  5;
  const  SATURDAY =  6;
  const  FEBRUARY =  2;
  const     MARCH =  2;
  const     APRIL =  3;
  const   OCTOBER = 10;
  const  NOVEMBER = 11;
  const    TEN_AM = 10;
  const ONE_BUSINESS_DAY  = 1; // The reminder day will be sent one business day before the timesheet-emailDay
  const TWO_BUSINESS_DAYS = 2;

  var numBusinessDays        = parseInt(SpreadsheetApp.getActive().getSheetByName('Timesheet').getRange(7, 4).getValue().match(/\d+/g)); // Extracts only numbers from a string
  var numBusinessDaysPlusOne = numBusinessDays + ONE_BUSINESS_DAY;
  var is_EmailDay_EffectedByHoliday = false, is_ReminderDay_EffectedByHoliday = false;
  var dayOfWeek, payPeriodString, payDateString, payDay, payWeekDay, emailDateString, reminderDateString, emailDate;
  
  payPeriodString = createDateString(month, startDay, year) + ' - ' + createDateString(month, endDay, year); // Get the pay period
  dayOfWeek = new Date(year, month - 1, endDay).getDay(); // Get the day of week for the end of period

  // Check if it is a holiday that will effect the pay period dates and then make the relevant changes
  if (month == FEBRUARY && startDay == 1) // Family Day
  {
    if (SpreadsheetApp.getActive().getSheetByName("Holidays").getRange(6, 4).getValue() == 15) // If Family Day is on the 15th
      dayOfWeek = 0, endDay -= ONE_BUSINESS_DAY;
  }
  else if (month == OCTOBER && startDay == 1) // Thanksgiving Day
  {
    if (dayOfWeek == THURSDAY)
      is_ReminderDay_EffectedByHoliday = true;
    else if (dayOfWeek > MONDAY && dayOfWeek < THURSDAY)
      is_EmailDay_EffectedByHoliday = true;
  }
  else if (month == NOVEMBER && startDay == 1) // Remembrance Day
  {
    var remembranceDay = SpreadsheetApp.getActive().getSheetByName("Holidays").getRange(13, 4, 1, 2).getValues(); // Day | Day of Week
    
    if (remembranceDay[0][1] == TUESDAY || (remembranceDay[0][1] == FRIDAY && remembranceDay[0][0] == 10) || (remembranceDay[0][1] == MONDAY && remembranceDay[0][0] == 12))
      is_ReminderDay_EffectedByHoliday = true;
    else if (remembranceDay[0][1] > TUESDAY && remembranceDay[0][1] < SATURDAY)
      is_EmailDay_EffectedByHoliday = true;
  }
  else if ((month == MARCH + 1 && startDay == 16) || (month == APRIL + 1 && startDay == 1)) // Good Friday
  {
    var goodFriday = SpreadsheetApp.getActive().getSheetByName("Holidays").getRange(7, 3, 1, 2).getValues(); // Month | Day
    const LAST_DAY_IN_PAY_PERIOD = (goodFriday[0][0] == APRIL) ? 15 : 31;
    const  EARLIEST_REMINDER_DAY = LAST_DAY_IN_PAY_PERIOD - 5;
    
    // Check if the month is a match for the holiday and if the days might effect the pay period dates
    if ( month - 1 == goodFriday[0][0] && (goodFriday[0][1] >= EARLIEST_REMINDER_DAY && goodFriday[0][1] <= LAST_DAY_IN_PAY_PERIOD) )
    {    
      if (goodFriday[0][1] == EARLIEST_REMINDER_DAY)
        is_ReminderDay_EffectedByHoliday = true;
      else if (goodFriday[0][1] == LAST_DAY_IN_PAY_PERIOD || dayOfWeek == SUNDAY || dayOfWeek == SATURDAY)
        endDay -= ONE_BUSINESS_DAY;
      else if (goodFriday[0][1] > EARLIEST_REMINDER_DAY && goodFriday[0][1] < LAST_DAY_IN_PAY_PERIOD)
        is_EmailDay_EffectedByHoliday = true;
    }
  }

  // If the end of pay period falls on a weekend, the pay day needs to roll back to the previous friday
  if (dayOfWeek == SATURDAY)
    endDay -= ONE_BUSINESS_DAY;
  else if (dayOfWeek == SUNDAY)
    endDay -= TWO_BUSINESS_DAYS;
  
  // Extract only the day from the pay date string and then get the day of the week
  payDateString = createDateString(month, endDay, year);
  payDay = parseInt(payDateString.substring(3, 5));
  payWeekDay = new Date(year, month - 1, payDay).getDay();
  
  if (is_EmailDay_EffectedByHoliday)
    endDay -= (payWeekDay - numBusinessDays >= TUESDAY) ? ONE_BUSINESS_DAY : numBusinessDaysPlusOne;
  else if (payWeekDay - numBusinessDays <= SUNDAY)
    endDay -= numBusinessDays;
  
  // Send the timesheet email the chosen number of business days before pay day
  endDay -= numBusinessDays;
  emailDateString = createDateString(month, endDay, year);
  emailDate = new Date(year, month - 1, endDay, TEN_AM); 
  
  // If the emailDay is not effected by a holiday
  if (!is_EmailDay_EffectedByHoliday)
  {   
    if (is_ReminderDay_EffectedByHoliday)
      endDay -= (payWeekDay - numBusinessDaysPlusOne >= TUESDAY) ? ONE_BUSINESS_DAY : numBusinessDaysPlusOne; // If the reminder day is effected, subtract an extra business day
    else if (payWeekDay - numBusinessDaysPlusOne == SUNDAY) // The reminder day falls on a weekend, therefore subtract two business days
      endDay -= TWO_BUSINESS_DAYS;
  }
  
  // Send the reminder email one business day before the timesheet needs to be submitted
  endDay -= ONE_BUSINESS_DAY;
  reminderDateString = createDateString(month, endDay, year);
  
  return [payPeriodString, payDateString, emailDateString, reminderDateString, emailDate];
}

/**
* This function sends the supervisor an email asking for approval of hours and turns on the automatic trigger to 
* send the unnapproved timesheet of the employee two business days before the pay day.
*
* @author Jarren Ralf
*/
function getApproval()
{
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActive();
  var timesheet = spreadsheet.getSheetByName('Timesheet');
  var hours = timesheet.getRange(11, 6, 16, 2).getValues(); // All the hours and break times
  var isVacationPayRequested, areBothOptionsSelected, areAllHoursValid, areAllBreakTimesValid, year, month, startDay, endDay, today, emailDate, response;
  
  [areAllHoursValid, areAllBreakTimesValid] = checkHours(hours);
  [isVacationPayRequested, areBothOptionsSelected] = checkVacationPay(timesheet);
  
  if (!isActive(spreadsheet))
  {
    var triggers = ScriptApp.getProjectTriggers(); // Get all of the triggers in the current project
    var employeesTriggerStatusRange = timesheet.getRange(15, 9);
    var areTriggersOn = areEmployeesTriggersOn(triggers);
    
    if (areTriggersOn)
    {
      ui.alert('Spreadsheet Not Active', 'I\'m sorry but the payroll manager forgot to activate your timesheet.' + 
               ' Please contact them and tell them to click on the green happy face on the Control Sheet.', ui.ButtonSet.OK);
      employeesTriggerStatusRange.check();
    }
    else
    {
      var firstResponse = ui.alert('Activate Employee Triggers?', 'Are you currently being instructed to turn on the employee triggers?', ui.ButtonSet.YES_NO);
      
      if (firstResponse == ui.Button.YES)
      {
        ScriptApp.newTrigger("setDaysOfPayPeriod").timeBased().onMonthDay( 1).atHour(4).create();
        ScriptApp.newTrigger("setDaysOfPayPeriod").timeBased().onMonthDay(16).atHour(4).create();
        employeesTriggerStatusRange.check();
        spreadsheet.toast('You have successfully turned on the triggers that will refresh your timesheet each week.', 'Triggers Created', 10)   
      }
      else
        employeesTriggerStatusRange.uncheck();
    }    
  }
  else
  {
    if (!areAllHoursValid)
      ui.alert('Not All Hours are Valid', 'Please modify the Start Time, End Time and Break Time columns in order to produce hours that are a multiple of 0.50.', ui.ButtonSet.OK);
    else if (!areAllBreakTimesValid)
      ui.alert('Not All Break Times are Valid', 'Please modify the Break Time column to ensure there are no negatives.', ui.ButtonSet.OK);
    else
    {
      if (areBothOptionsSelected)
        ui.alert('Please select ONLY ONE vacation pay request type.');
      else
      {
        [year, month, startDay, endDay,, today] = determinePayPeriod();
        [,,,, emailDate] = generateDates(year, month + 1, startDay, endDay);
        
        if (today > emailDate) // The timesheet is late (the email date has passed)
        {
          response = ui.alert('Your Timesheet is Past Due', 'You are about to send an email directly to the payroll manager with ' +
                              'an unnapproved copy of your timesheet attached. Do you want to proceed?', ui.ButtonSet.YES_NO);
          
          if (response == ui.Button.YES)
          {
            timesheet.getRange(28, 7).setFormula("=EmployeeSignature"); // Set the formula to import the employee signature
            (isVacationPayRequested) ? unapprovedTimesheet('UnapprovedTimesheet_withVacationPay', spreadsheet) : unapprovedTimesheet('UnapprovedTimesheet', spreadsheet); 
          }
        }
        else
        {
          response = ui.alert('Get Approval of Your Hours Worked', 'You are about to send an email to your supervisor ' +
                              'requesting approval of your hours. Do you want to proceed?', ui.ButtonSet.YES_NO);
          
          if (response == ui.Button.YES)
          {
            timesheet.getRange(28, 7).setFormula("=EmployeeSignature"); // Set the formula to import the employee signature
            getApprovalEmail();                                         // Send the supervisor the email asking for approval of hours
            (isVacationPayRequested) ? autoSendEmail(isVacationPayRequested) : autoSendEmail(0); // Turn on the trigger for sending the unnapproved email
            spreadsheet.toast('Requesting Approval from your Supervisor', 'Email Sent', 10);
          }
        }
      }
    }
  }
}

/**
* This function sends an email to the supervisor asking them to review and approve the hours of the employee.
*
* @author Jarren Ralf
*/
function getApprovalEmail()
{
  var employeeFullName, recipientFullName, supervisorFirstName, employeeEmail, supervisorEmail, payPeriod;
  [employeeFullName, recipientFullName,,,, supervisorFirstName, employeeEmail,, supervisorEmail, payPeriod] = getEmailVariables();
  
  // Read in and set the appropriate variables on the html template
  var templateHtml = HtmlService.createTemplateFromFile('GetApproval');
  templateHtml.recipientName  = recipientFullName;
  templateHtml.employeeName   = employeeFullName;
  templateHtml.supervisorName = supervisorFirstName;
  templateHtml.payPeriod      = payPeriod;
  templateHtml.sheetURL       = getSheetUrl();
  
  var emailSubject = employeeFullName + ' hours approval for pay period ' + payPeriod; // The subject of the email
  var message = templateHtml.evaluate().getContent();                                  // Get the contents of the html document
  
  // Fire an email with following chosen parameters
  MailApp.sendEmail({         to: supervisorEmail, 
                         replyTo: employeeEmail,
                            name: employeeFullName,
                         subject: emailSubject, 
                        htmlBody: message});
}

/**
* This is a function I found online that returns the total number of days in a month. Expected input for 
* Date() is Date(YYYY, MMM, DD). YYYY, MMM, DD are all integers. YYYY > 0, 0 <= MM <= 11, 1 <= DD <= 31.
* Notice month + 1 gets the following month. But for DD, the input is 0, which is not in the interval 
* 1 <= DD <= 31. So infact the day jumps back to the previous month and grabs the last day.
*
* @param  {Number} month The chosen month
* @param  {Number} year  The chosen year
* @return {Number} The number of days in the chosen month
*/
function getDaysInMonth(month, year)
{
  return new Date(year, month + 1, 0).getDate();
}

/**
* This functions gets all of the possible email parameters available to the user. Different combinations of these
* variables can be used to construct the varying styles of emails that will be sent by this script. 
*
* @return {...String[], Date, Sheet} An array of all possible email parameters, such as: employeeName, recipientFullName, 
*         supervisorFullName, recipientFirstName, supervisorFirstName, employeeEmail, recipientEmail, supervisorEmail, 
*         payPeriod, emailDay, today and timesheet.
* @author Jarren Ralf
*/
function getEmailVariables()
{
  var spreadsheet = SpreadsheetApp.getActive();
  var timesheet = spreadsheet.getSheetByName('Timesheet');
  timesheet.hideColumns(9); // Hide the signature column
  Utilities.sleep(1000);    // Pause so that the column gets hidden prior to the email being sent
  
  var supervisorSheet = spreadsheet.getSheetByName('Control Sheet');
  var       timesheetVariables =       timesheet.getRange(2, 4, 5).getValues();
  var supervisorSheetVariables = supervisorSheet.getRange(3, 2, 2).getValues();
  var   employeeEmail = timesheetVariables[1][0];
  var  recipientEmail = timesheetVariables[4][0];
  var supervisorEmail = supervisorSheetVariables[1][0];
  var    employeeFullName  = timesheetVariables[0][0];
  var        employeeNames = employeeFullName.split(" ");
  var   employeeFirstName  = employeeNames[0];
  var   recipientFullName  = timesheetVariables[3][0];
  var       recipientNames = recipientFullName.split(" ");
  var  recipientFirstName  = recipientNames[0];
  var  supervisorFullName  = supervisorSheetVariables[0][0];
  var      supervisorNames = supervisorFullName.split(" ");
  var supervisorFirstName  = supervisorNames[0];
  var year, month, startDay, endDay, today, dates, payPeriod, emailDate;
  
  // Check which pay period we are in, (Either 1-15 or 16-[end of month]) and then set the appropriate values
  [year, month, startDay, endDay,, today] = determinePayPeriod(); 
  [payPeriod,,,, emailDate] = generateDates(year, month + 1, startDay, endDay);

  
  return [employeeFullName,  recipientFullName,  supervisorFullName, 
          employeeFirstName, recipientFirstName, supervisorFirstName, 
          employeeEmail,     recipientEmail,     supervisorEmail,
          payPeriod,         emailDate,          today, 
          timesheet];
}

/**
* This function calculates the day that New Years Day, Canada Day, Remembrance Day, and Christmas Day, is observed on for the giving year and month. 
*
* @param  {Number}  year The given year
* @param  {Number} month The given month
* @return {Number}   day The day of the Holiday for the particular year and month
* @author Jarren Ralf
*/
function getDay(year, month)
{
  const JANUARY  =  0;
  const JULY     =  6;
  const NOVEMBER = 10;
  const DECEMBER = 11;
  const SUNDAY   =  0;
  const SATURDAY =  6;
  
  if (month == JANUARY || month == JULY || month == DECEMBER) // New Years Day or Canada Day or Christmas Day
  {
    var holiday = (month == DECEMBER) ? new Date(year, month, 25) : new Date(year, month);
    var dayOfWeek = holiday.getDay();
    var day = (dayOfWeek == SATURDAY) ? holiday.getDate() + 2 : ( (dayOfWeek == SUNDAY) ? holiday.getDate() + 1 : holiday.getDate() ); // Rolls over to the following Monday
  }
  else if (month == NOVEMBER) // Remembrance Day
  {
    var holiday = new Date(year, month, 11);
    var dayOfWeek = holiday.getDay();
    var day = (dayOfWeek == SATURDAY) ? holiday.getDate() - 1 : ( (dayOfWeek == SUNDAY) ? holiday.getDate() + 1 : holiday.getDate() ); // Rolls back to Friday, or over to Monday
  }
  
  return day;
}

/**
* This function calculates what the nth Monday in the given month is for the given year. This function is used for determining the holidays in a given year.
* Victoria Day is an exception to the rule since it is defined to be the preceding Monday before May 25th. The fourth Boolean parameter handles this scenario.
*
* @param  {Number}              n : The nth Monday the user wants to be calculated
* @param  {Number}          month : The given month
* @param  {Number}           year : The given year
* @param  {Boolean} isVictoriaDay : Whether it is Victoria Day or not
* @return {Number} The day of the month that the nth Monday is on (or that Victoria Day is on)
* @author Jarren Ralf
*/
function getMonday(n, month, year, isVictoriaDay)
{
  const NUM_DAYS_IN_WEEK = 7;
  var firstDayOfMonth = new Date(year, month).getDay();
  
  if (isVictoriaDay)
    n = (firstDayOfMonth % (NUM_DAYS_IN_WEEK - 1) < 2) ? 4 : 3; // Corresponds to the Monday preceding May 25th 
  
  return ((NUM_DAYS_IN_WEEK - firstDayOfMonth + 1) % NUM_DAYS_IN_WEEK) + NUM_DAYS_IN_WEEK*n - 6;
}

/**
* This function gets the spreadsheets URL.
*
* @return {String} url The url of the current spreadsheet
*/
function getSheetUrl()
{
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var ss = SS.getActiveSheet();
  var url = '';
  url += SS.getUrl();
  url += '#gid=';
  url += ss.getSheetId(); 
  return url;
}

/**
* This function adds the proper suffix to the number.
*
* @param  {Number} The inputed number representing the day of the month
* @return {String} The suffix which is either "st", "nd", "rd", or "th"
* @author Jarren Ralf
*/
function getSuffix(n)
{
  return (n == 1 || n == 21 || n == 31) ? "st" : ( (n == 2 || n == 22) ? "nd" : ( (n == 3 || n == 23) ? "rd" : "th" ) );
}

/**
* This function creates a copy of the current spreadsheet and saves it in the same root folder as the original in the google drive.
* Then it converts the sheets to pdfs and only keeps the timesheet.
*
* @return {Blob} timeSheetPDF A pdf copy of the timesheet
*/
function getTimesheetPDF()
{
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var   timesheet = SpreadsheetApp.getActive().getSheetByName('Timesheet');
  var   sheetName = timesheet.getSheetName();
  var     parents = DriveApp.getFileById(spreadSheet.getId()).getParents();          // For finding the root folder of the save destination 
  var      folder = (parents.hasNext()) ? parents.next() : DriveApp.getRootFolder(); // Get folder containing spreadsheet to save pdf in
  var  copySpreadsheet = SpreadsheetApp.open(DriveApp.getFileById(spreadSheet.getId()).makeCopy("tmp_convert_to_pdf", folder)); // Copy whole spreadsheet
  var           sheets = copySpreadsheet.getSheets(); // Get all the sheets
  var destinationSheet = sheets[0]; // Get the sheet that we are going to copy all content to
  
  // Delete redundant sheets.
  for (var i = 0; i < sheets.length; i++)
  {
    if (sheets[i].getSheetName() != sheetName) // Only keep the Timesheet
      copySpreadsheet.deleteSheet(sheets[i]);  // Delete the rest
  }
  
  var sourceValues = timesheet.getRange(1, 1, timesheet.getMaxRows(), timesheet.getMaxColumns()).getValues();
  destinationSheet.getRange(1, 1, destinationSheet.getMaxRows(), destinationSheet.getMaxColumns()).setValues(sourceValues); // Replace cell values with text
  var timeSheetPDF = copySpreadsheet.getBlob().getAs('application/pdf').setName(sheetName + ".pdf");                        // Save a copy of the sheet as a pdf.
  DriveApp.getFileById(copySpreadsheet.getId()).setTrashed(true); // Delete the temporary sheet.
  
  return timeSheetPDF;
}

/**
* This function checks if the given value has content or not. A cell value may have been deleted, in which case the updated value is undefined for example.
*
* @param  {Object}  value The inputted value
* @return {Boolean} Whether the input paramater has content or not.
* @author Jarren Ralf
*/
function hasContent(value)
{
  return value !== undefined;
}

/**
* This function checks if the employee has signed their timesheet (clicked the getApproval button), by way of checking if there is a formula in the employee signature range.
*
* @param   {Sheet}  sheet The timesheet
* @return {Boolean} Whether the Employee has signed the timesheet or not.
* @author Jarren Ralf
*/
function hasEmployeeSigned(sheet)
{
  return !isBlank(sheet.getRange(28, 7).getFormula());
}

/**
* This function sends an email to the "recipient" or the payroll manager, from the supervisor. The email states that the
* employee's hours have been approved. If the supervisor has clicked the approval button before the timesheet-email day, 
* then an email is sent with an attached pdf copy of the employee's hours. Otherwise, an automatic email of the 
* unnapproved hours of the employee has already been sent, therefore the approval email will NOT have an attached copy
* of the employee's hours, which is intended to reduce duplicates for the payroll manager.
* 
* @param {String}        htmlDoc   The email template to be created
* @param {Spreadsheet} spreadsheet The active spreadsheet
* @author Jarren Ralf
*/
function hoursApprovalEmail(htmlDoc, spreadsheet)
{ 
  var employeeFullName, supervisorFullName, employeeFirstName, recipientFirstName, employeeEmail, recipientEmail, 
      supervisorEmail, payPeriod, emailDate, today, message, templateHtml_1, templateHtml_2, timeSheetPDF;
  [employeeFullName,, supervisorFullName, employeeFirstName, recipientFirstName,, employeeEmail, 
   recipientEmail, supervisorEmail, payPeriod, emailDate, today] = getEmailVariables();
  
  var emailSubject = 'APPROVED: ' + employeeFullName + ' Timesheet for pay period ' + payPeriod; // The subject of the email
  
  if (today > emailDate) // If it is past the emailDay, then don't attach the pdf
  {
    templateHtml_1 = HtmlService.createTemplateFromFile('HoursApprovalNoReply');
    templateHtml_1.recipientName    = recipientFirstName;
    templateHtml_1.supervisorName   = supervisorFullName;
    templateHtml_1.employeeFullName = employeeFullName;
    templateHtml_1.payPeriod        = payPeriod; 
    message = templateHtml_1.evaluate().getContent();
    
    MailApp.sendEmail({      to: recipientEmail, 
                        replyTo: employeeEmail,
                           name: supervisorFullName,
                             cc: supervisorEmail, 
                        subject: emailSubject, 
                       htmlBody: message});
    
    spreadsheet.toast('Since the timesheet is past due, you have just sent a brief hours approval email with no attachment.', 'Email Sent', 10);
  }
  else // Otherwise attach the pdf
  {
    templateHtml_2 = HtmlService.createTemplateFromFile(htmlDoc);
    templateHtml_2.recipientName     = recipientFirstName;
    templateHtml_2.supervisorName    = supervisorFullName;
    templateHtml_2.employeeFullName  = employeeFullName;
    templateHtml_2.employeeFirstName = employeeFirstName;
    templateHtml_2.employeeEmail     = employeeEmail;
    templateHtml_2.payPeriod         = payPeriod;
    message = templateHtml_2.evaluate().getContent();
    timeSheetPDF = getTimesheetPDF();
    
    MailApp.sendEmail({      to: recipientEmail, 
                        replyTo: employeeEmail,
                           name: supervisorFullName,
                             cc: supervisorEmail, 
                        subject: emailSubject, 
                       htmlBody: message,
                    attachments: timeSheetPDF});
    
    spreadsheet.toast('Approved Timesheet Attached', 'Email Sent', 10);
  }
}

/**
* This function checks if the spreadsheet is active (are all the triggers on?)
*
* @param  {Spreadsheet} ss The spreadsheet
* @return   {Boolean}   Whether the spreadsheet is active or not
* @author Jarren Ralf
*/
function isActive(ss)
{
  return ss.getSheetByName('Control Sheet').getRange(1, 3).getValue();
}

/**
* This function checks if the employee's hours have been approved by the supervisor, by way of checking if there is a formula in the supervisor signature range.
*
* @param   {Sheet}  sheet The timesheet
* @return {Boolean} Whether the supervisor has approved the hours or not.
* @author Jarren Ralf
*/
function isApproved(sheet)
{
  return !isBlank(sheet.getRange(27, 4).getFormula());
}

/**
* This function checks if the given value is blank or not.
*
* @param  {String}  value The inputted string
* @return {Boolean} Whether the input paramater is blank or not
* @author Jarren Ralf
*/
function isBlank(value)
{
  return value == '';
}

/**
* This function checks if the new edited value is a multiple of minBreakTime or not, excluding when the cell value is changed to blank (deleted) and negative. 
*
* @param  {Number}  minBreakTime The minimum amount of break time you can take
* @param  {Object}     value     The new value of the edited cell
* @return {Boolean} Whether the break time is valid or not
* @author Jarren Ralf
*/
function isBreakTimeValid(minBreakTime, value)
{
  return !hasContent(value) || value % minBreakTime === 0 && value >= 0;
}

/**
* This function checks if a particular column is being edited or not. It also ensure that it is the only column being edited.
*
* @param  {String}  col      The column that is being checked for an edit
* @param  {String}  colStart The start column of the edited range
* @param  {String}  colEnd   The end column of the edited range
* @return {Boolean} Whether the given column is being edited or not
* @author Jarren Ralf
*/
function isColumnEdited(col, colStart, colEnd)
{
  return colStart == col && colEnd == col;
}

/**
* This function checks if the givennumber is even or not.
*
* @param  {Number}  num The given number
* @return {Boolean} Whether the input number is even or not
* @author Jarren Ralf
*/
function isEven(num)
{
  return num % 2 == 0;
}

/**
* This function checks if the given input is a number or not.
*
* @param  {Object}  num The inputted argument, assumed to be a number.
* @return {Boolean} Whether the input paramater is a number or not
* @author Jarren Ralf
*/
function isNumber(num)
{
  return !(isNaN(parseInt(num)));
}

/**
* This function recativates the spreadsheet by creating all of the appropriate triggers. 
* This means automatic emails will be sent and the hours will reset each pay period.
*
* @author Jarren Ralf
*/
function reactivateTimesheet()
{
  var ui = SpreadsheetApp.getUi();
  var firstResponse = ui.alert('Reactivate Timesheet', 'You are about to create all of the triggers that will make this ' +
                               'spreadsheet fully functional and automatically send emails to the appropriate people. '   +
                               'Do you want to proceed?', ui.ButtonSet.YES_NO);
  
  if (firstResponse == ui.Button.YES)
  {
    var secondResponse = ui.alert('Are you sure?', ui.ButtonSet.YES_NO);
    
    if (secondResponse == ui.Button.YES)
    {
      if (SpreadsheetApp.getActive().getSheetByName("Timesheet").getRange(15, 9).getValue())
      {
        updateStatusMessage(true);
        reset();
      }
      else
        ui.alert('Employee\'s Triggers Not Created', 'Turn off the protections on cell A15 and tell the employee to click the '
                 + 'Get Approval button in order to turn on their project triggers.', ui.ButtonSet.YES_NO);
    }
  }
}

/**
* This function sends an email to the employee to remind them to review their hours and send them in for approval.
*
* @author Jarren Ralf
*/
function reminderEmail()
{
  var recipientFullName, supervisorFullName,   employeeEmail,   payPeriod;
  [, recipientFullName, supervisorFullName,,,, employeeEmail,,, payPeriod] = getEmailVariables();

  // Read in and set the appropriate variables on the html template
  var templateHtml = HtmlService.createTemplateFromFile('Reminder');
  templateHtml.recipientName  = recipientFullName;
  templateHtml.supervisorName = supervisorFullName;
  templateHtml.payPeriod      = payPeriod;
  templateHtml.sheetURL       = getSheetUrl();
  
  var emailSubject = 'REMINDER: Get your hours approved for pay period ' + payPeriod; // The subject of the email
  var message = templateHtml.evaluate().getContent(); // Get the contents of the html document
  var sendersName = "TIMESHEET REMINDER";             // Set the senders name
  
  // Fire an email with following chosen parameters
  MailApp.sendEmail({         to: employeeEmail, 
                            name: sendersName,
                         subject: emailSubject, 
                        htmlBody: message});
}

/**
* This function removes the range protections that disallow the employee to run the 'setDaysOfPayPeriod' function. It will be run 
* via owner's trigger at 3 A.M. prior to the employee's trigger which runs at 4 A.M. at the beginning of each pay period.
*
* @author Jarren Ralf
*/
function removeProtections()
{
  var protections = SpreadsheetApp.getActive().getProtections(SpreadsheetApp.ProtectionType.RANGE); // All of the range protections on the sheet
  var protectionsToRemove = ['Dates', 'Header # 1', 'Hours', 'Supervisor Signature Range'];         // The descriptions of the range protections
  for (var i = 0; i < protections.length; i++)
  {
    if (protectionsToRemove.indexOf(protections[i].getDescription()) !== -1)
      protections[i].remove();
  }
}

/**
* This function runs eight triggers at the start of each pay period. The fifth trigger runs at the start of every month.
*
* @author Jarren Ralf
*/
function reset()
{
  const START_OF_FIRST_PAY_PERIOD  =  1;
  const START_OF_SECOND_PAY_PERIOD = 16;
  
  ScriptApp.newTrigger("shouldDatesReset").timeBased().onMonthDay(START_OF_FIRST_PAY_PERIOD).atHour(2).create();
  
  ScriptApp.newTrigger("deleteReminderEmailTrigger").timeBased().onMonthDay(START_OF_FIRST_PAY_PERIOD).atHour(3).create();
  ScriptApp.newTrigger("deleteReminderEmailTrigger").timeBased().onMonthDay(START_OF_SECOND_PAY_PERIOD).atHour(3).create();
  
  ScriptApp.newTrigger("autoSendEmail").timeBased().onMonthDay(START_OF_FIRST_PAY_PERIOD).atHour(5).create();  // Create Reminder email Trigger
  ScriptApp.newTrigger("autoSendEmail").timeBased().onMonthDay(START_OF_SECOND_PAY_PERIOD).atHour(5).create(); // Create Reminder email Trigger

  setDaysOfPayPeriod(); // Reset the timesheet with the current pay period
  autoSendEmail();      // Turn on the trigger to send the reminder email
}

/**
* This function sets all of the pay periods, pay days and timesheet-email days for a given year and prints
* them on the Pay Periods page. 
*
* @author Jarren Ralf
*/
function setDates()
{
  var year = new Date().getFullYear(); // An integer corresponding to today's year
  var startDay, endDay, month;
  var data = [["Pay Period", "Pay Day", "Email Day", "Reminder Day"]];
  
  const ROW_START = 2;
  const COL_START = 1;
  const NUM_COLUMNS = 4;
  const NUM_PAY_PERIODS = 24; // The total number of pay periods in a year
  
  for (var i = 0; i < NUM_PAY_PERIODS; i++)
  {
    if (isEven(i)) // If the index is even
    {
         month = i/2 + 1;
      startDay = 1;
        endDay = 15;
    }
    else // If the index is odd
    {
         month = (i + 1)/2;
      startDay = 16;
        endDay = getDaysInMonth(month - 1, year);
    }
    data[i + 1] = generateDates(year, month, startDay, endDay); // Add the generated dates to the next row of the output array
  }
  var output = data.filter(value => value.splice(NUM_COLUMNS)); // Remove the last column
  SpreadsheetApp.getActive().getSheetByName('Pay Periods').getRange(ROW_START, COL_START, NUM_PAY_PERIODS + 1, NUM_COLUMNS).setValues(output);
}

/**
* This function clears the appropriate cells in the timesheet then calculates and sets the correct days for 
* the pay period. It also fills in the default hours of the employee as Mon-Friday 9-5, as well as printing
* the pay period on the sheet.
*
* @author Jarren Ralf
*/
function setDaysOfPayPeriod()
{  
  const START_TIME = '9:00 AM'; // Set your general startTime
  const   END_TIME = '5:00 PM'; // Set your general endTime
  const  START_ROW = 11; // The row number where the data begins 
  const  START_COL =  2; // The col number where the data begins 
  const   SATURDAY =  6;
  const     SUNDAY =  0;
  const   NUM_COLS =  6;
  const NUM_DAYS_IN_WEEK =  7;
  const     MAX_NUM_DAYS = 16;
  const LIGHT_GREEN = '#ecefe9';
  const WEEK_DAYS = ['Sunday', 'Monday','Tuesday', 'Wednesday','Thursday', 'Friday' ,'Saturday'];
  
  var sheet = SpreadsheetApp.getActive().getSheetByName('Timesheet');
  var rangesToClear = ['G8:G8', 'D27:D27', 'G28:G28'];
  var  timeRange = ['D11:E26'];
  var breakRange = ['F11:F26'];
  var data = new Array(MAX_NUM_DAYS).fill(null).map(() => new Array(NUM_COLS)); // A multiarray of null values
  var year, month, startDay, endDay, firstDayOfWeek, payPeriod, dayOfWeek;
  var whiteBackgroundRanges = [], greenBackgroundRanges = [];
  
  // Since the employee creates the trigger that runs setDaysOfPayPeriod, they will delete the UnapprovedTimesheetEmailTrigger every pay period
  deleteUnapprovedTimesheetEmailTrigger();
  
  // Fill up the array of Ranges for setting the Background colour sceme
  for (var c = 0; c < MAX_NUM_DAYS; c += 2)
  {
    whiteBackgroundRanges.push('B' + (11 + c) + ':G' + (11 + c));
    greenBackgroundRanges.push('B' + (12 + c) + ':G' + (12 + c));
  }
  
  // Clear all of the specified ranges, reset the background colours, uncheck the Request Vacation Pay checkbox, and reset the number format
  sheet.getRange( timeRange).setNumberFormat("h:mm am/pm");
  sheet.getRange(breakRange).setNumberFormat("0");
  sheet.getRangeList(rangesToClear).clearContent();
  sheet.getRangeList(whiteBackgroundRanges).setBackground('white');
  sheet.getRangeList(greenBackgroundRanges).setBackground(LIGHT_GREEN);
  sheet.getRange(8, 5).uncheck();
  
  // Determine and set the pay period
  [year, month, startDay, endDay, firstDayOfWeek] = determinePayPeriod();
  payPeriod = generateDates(year, month + 1, startDay, endDay);
  sheet.getRange(4, 4).setValue(payPeriod); // Set the pay period in the chosen cell
  var numDays = endDay - startDay + 1;      // Total number of days in the given pay period
  
  // Set the days, start time and end time for the current pay period in the appropriate columns along with the formulas
  for (var i = 0; i < numDays; i++, firstDayOfWeek++)
  {
    dayOfWeek  = firstDayOfWeek % NUM_DAYS_IN_WEEK; // The day of the week represented by a number 0-6
    data[i][0] = startDay + i;
    data[i][1] = WEEK_DAYS[dayOfWeek]; // Prints the day of week
    data[i][5]  = "=IF(OR(ISBLANK(StartTime),ISBLANK(EndTime)),\"\",(EndTime-StartTime)*24-BreakTime/60)"; // Calculate the hours worked that day
    
    if (dayOfWeek != SUNDAY && dayOfWeek != SATURDAY) // No Weekends
    {
      data[i][2] = START_TIME;
      data[i][3] = END_TIME;
    }
  }
  sheet.getRange(START_ROW, START_COL, MAX_NUM_DAYS, NUM_COLS).setValues(data);
}

/**
* This function sets all of the holidays for the current year.
*
* @author Jarren Ralf
*/
function setHolidays()
{
  const NUM_COLUMNS  =  6;
  const NUM_HEADERS  =  4;
  const NUM_HOLIDAYS = 10; // The total number of holidays in a year
  const             DATE = 1; // These are the array indices for readability
  const              DAY = 3;
  const         WEEK_DAY = 4;
  const NAME_OF_WEEK_DAY = 5;
  const JAN =  0, FEB =  1, MAY =  4, JUL =  6, AUG =  7, SEP =  8, OCT =  9, NOV = 10, DEC = 11;
  const YEAR = new Date().getFullYear(); // An integer corresponding to today's year
  const WEEK_DAYS = ['Sunday', 'Monday','Tuesday', 'Wednesday','Thursday', 'Friday' ,'Saturday'];
  var MMM, DD, dayOfWeek;
  
  [MMM, DD] = calculateGoodFriday(YEAR);

  var output = [["Holidays",                                                                                                                         "", "", "", "", ""],
                ["Are all of these Holidays correct?\n\nIf not, please contact the payroll manager immediately and tell them to update this sheet!", "", "", "", "", ""],
                ["https://www2.gov.bc.ca/gov/content/employment-business/employment-standards-advice/employment-standards/statutory-holidays"      , "", "", "", "", ""],
                ["Name", "Date (Observed)", "Month - 1", "Day", "Day of Week", "Day of Week"],
                ["New Year's Day",   new Date(YEAR, JAN, getDay(YEAR, JAN)),          JAN],
                ["Family Day",       new Date(YEAR, FEB, getMonday(3, FEB, YEAR)),    FEB],
                ["Good Friday",      new Date(YEAR, MMM, DD),                         MMM],
                ["Victoria Day",     new Date(YEAR, MAY, getMonday(0, MAY, YEAR, 1)), MAY], // Victoria Day is on the Monday preceding May 25. The fourth parameter controls this scenario
                ["Canada Day",       new Date(YEAR, JUL, getDay(YEAR, JUL)),          JUL],
                ["BC Day",           new Date(YEAR, AUG, getMonday(1, AUG, YEAR)),    AUG],
                ["Labour Day",       new Date(YEAR, SEP, getMonday(1, SEP, YEAR)),    SEP],
                ["Thanksgiving Day", new Date(YEAR, OCT, getMonday(2, OCT, YEAR)),    OCT],
                ["Remembrance Day",  new Date(YEAR, NOV, getDay(YEAR, NOV)),          NOV],
                ["Christmas Day",    new Date(YEAR, DEC, getDay(YEAR, DEC)),          DEC]];
  
  for (var i = 0; i < NUM_HOLIDAYS; i++)
  {
    dayOfWeek = output[NUM_HEADERS + i][DATE].getDay();
    output[NUM_HEADERS + i][             DAY] = output[NUM_HEADERS + i][DATE].getDate();
    output[NUM_HEADERS + i][        WEEK_DAY] = dayOfWeek;
    output[NUM_HEADERS + i][NAME_OF_WEEK_DAY] = WEEK_DAYS[dayOfWeek];
  }

  SpreadsheetApp.getActive().getSheetByName('Holidays').getRange(1, 1, NUM_HOLIDAYS + NUM_HEADERS, NUM_COLUMNS).setValues(output);
}

/**
* This function resets the protections of the sheet that were removed at 3 A.M. in order for the employee to
* run a function on a trigger of which they would have otherwise not been able to run. This function will be 
* run at 5 A.M. by the owner's trigger, which is following the employee's trigger ran at 4 A.M.
*
* @author Jarren Ralf
*/
function setProtections()
{
  var spreadsheet  = SpreadsheetApp.getActive();
  var descriptions =                    ['Dates',   'Header # 1', 'Hours',   'Supervisor Signature Range']; // The descriptions of the range protections
  var ranges = spreadsheet.getRangeList(['A11:C26', 'A1:H7',      'G11:H27', 'D27:E28']).getRanges();       // The ranges to protect
  var numRanges = ranges.length;
  var employeeEmail = spreadsheet.getSheetByName('Timesheet').getRange(3, 4).getValue();
  var protections = [];

  for (var i = 0; i < numRanges; i++)
  {
    protections[i] = ranges[i].protect()
    protections[i].setDescription(descriptions[i]);
    
    // Remove only the employee from editing the final range (leaving the supervisor), and remove all of the editors for the rest
    (i == numRanges - 1) ? protections[i].removeEditor(employeeEmail) : protections[i].removeEditors(protections[i].getEditors());
  }
}

/**
* This function is a quick work around to set a yearly trigger. The trigger runs a function every month. 
* That function only executes when the month is January. 
* In this case specifically, it sets all of the dates on the Pay Periods sheet of this spreadsheet.
*/
function shouldDatesReset()
{
  var date = new Date();
  const JANUARY = 0;
  if (date.getMonth() === JANUARY)
  {
    setHolidays();
    Utilities.sleep(10000); // Pause so that the Holidays sheet is updated because it's data is used in the setDates function
    setDates();
  }
}

/**
* This function sends an email to the "recipient" or the payroll manager, with an attached pdf copy of the employees hours
* and a message that states that the supervisor has not approved the hours of the employee.
*
* @param   {String}      htmlDoc   : The email template to be created
* @param {Spreadsheet} spreadsheet : The active spreadsheet
* @author Jarren Ralf
*/
function unapprovedTimesheet(htmlDoc, spreadsheet)
{
  var employeeFullName,  supervisorFullName,  recipientFirstName,  employeeEmail, recipientEmail, supervisorEmail, payPeriod,   timesheet;
     [employeeFullName,, supervisorFullName,, recipientFirstName,, employeeEmail, recipientEmail, supervisorEmail, payPeriod,,, timesheet] = getEmailVariables();
  
  if (!isApproved(timesheet)) // Only send the unnapproved email if the supervisor signature is missing from the timesheet
  {
    // Read in and set the appropriate variables on the html template
    var templateHtml = HtmlService.createTemplateFromFile(htmlDoc);
    templateHtml.employeeName   = employeeFullName;
    templateHtml.supervisorName = supervisorFullName;
    templateHtml.recipientName  = recipientFirstName;
    templateHtml.employeeEmail  = employeeEmail;
    
    var emailSubject = 'Timesheet for pay period ' + payPeriod; // The subject of the email
    var message = templateHtml.evaluate().getContent();         // Get the contents of the html document
    var timeSheetPDF = getTimesheetPDF();
    
    // Fire an email with following chosen parameters
    MailApp.sendEmail({         to: recipientEmail, 
                           replyTo: employeeEmail,
                              name: employeeFullName,
                                cc: supervisorEmail,
                           subject: emailSubject, 
                          htmlBody: message, 
                       attachments: timeSheetPDF});

    spreadsheet.toast('Unapproved Timesheet Attached', 'Email Sent', 10);
  }
  else // The signature of the supervisor must already be on the sheet. Therefore only the supervisor can resend the timesheet by clicking on the Approve Hours button on the Control page
  {
    const ui = SpreadsheetApp.getUi();
    ui.alert('WARNING! Email was NOT Sent', 'You are unable to send a second copy of your timesheet to the payroll manager. ' + 
    'Please advise your supervisor to click on the \'Approve Hours\' button if you would like another copy of your timesheet sent.', 
    ui.ButtonSet.OK);
  }
}

/**
* This function sends the standard unapproved timesheet email. See the unapprovedTimesheet function.
*
* @author Jarren Ralf
*/
function unapprovedTimesheetEmail()
{
  unapprovedTimesheet('UnapprovedTimesheet', SpreadsheetApp.getActive());
}

/**
* This function adds a vacation pay request to the body of the email.
*
* @author Jarren Ralf
*/
function unapprovedTimesheet_withVacationPayEmail()
{
  unapprovedTimesheet('UnapprovedTimesheet_withVacationPay', SpreadsheetApp.getActive());
}

/**
* Change the value of the edited cell to blank or back to the old value.
*
* @param  {Object} oldValue The value of a cell prior to a user edit
* @param  {Range}  range    The range of the user edit
* @author Jarren Ralf
*/
function undoEdit(oldValue, range)
{
  (!hasContent(oldValue)) ? range.setValue('') : range.setValue(oldValue);
}

/**
* This function updates the status message on the Supervisor page.
*
* @param {Boolean} isActive Whether the spreadsheet is active or not
* @author Jarren Ralf
*/
function updateStatusMessage(isActive)
{
  var message, colour;
  var statusRange = SpreadsheetApp.getActive().getSheetByName('Control Sheet').getRange(1, 1, 1, 3);
  (isActive) ? (message = "This spreadsheet is active", colour = "green") : (message = "This spreadsheet is deactivated", colour = "red");
  statusRange.setFontColor(colour);
  statusRange.setValues([[message, null, isActive]]);
}