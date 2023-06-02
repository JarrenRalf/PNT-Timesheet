/**
 * This function handles the onEdit events of the spreadsheet, specifically looking for changes to the timesheet.
 * 
 * @param {Object} e : The event object that is created whenever the spreadsheet is edited.
 * @author Jarren Ralf
 */
function installedOnEdit(e)
{
  const spreadsheet = e.source;
  const sheet = spreadsheet.getActiveSheet(); // The active sheet that the onEdit event is occuring on
  const checkboxRange = e.range;
  const row = checkboxRange.rowStart;
  const col = checkboxRange.columnStart;
  const isSingleCell = (row === checkboxRange.rowEnd && col === checkboxRange.columnEnd);

  try
  {
    if (sheet.getSheetName() === "Timesheet" && isSingleCell && checkboxRange.isChecked() && row === 1)
    {
      if (col === 6)
        getApproval(spreadsheet, sheet, checkboxRange)
      else if (col === 8)
        approveHours(spreadsheet, sheet, checkboxRange)
    }
  } 
  catch (err) 
  {
    var error = err['stack'];
    Logger.log(error)
    Browser.msgBox(error)
    throw new Error(error);
  }
}

/**
* This function places the supervisor's signature on the timesheet, representing approval of the employees hours. 
* After which it sends the hours approval email, then the trigger to send the unnapproved timesheet email is deleted.
*
* @param {Spreadsheet} spreadsheet : The active spreadsheet.
* @param    {Sheet}      sheet     : The active sheet.
* @param    {Range}  checkboxRange : The active range.
* @author Jarren Ralf
*/
function approveHours(spreadsheet, sheet, checkboxRange)
{
  if (sheet.getRange(28, 8).getFormula()) // The employee signature range
  {
    var year, month, startDay, endDay, emailDay;
    [year, month, startDay, endDay,, today] = determinePayPeriod();
    [,, emailDay] = generateDates(year, month, startDay, endDay, spreadsheet);
    sheet.getRange(27, 4).setFormula("=SupervisorSignature").activate(); // Set the formula to import the supervisor signature

    if (today > emailDay)
      sendEmail_HoursApproval_NoAttachment()
    else
    {
      const isVacationPayRequested = sheet.getRange(8, 2).getBackground() === '#f4c7c3';             
      (isVacationPayRequested) ? sendEmail_HoursApproval_withVacationPay() : sendEmail_HoursApproval(); // Send email
    }

    checkboxRange.offset(0, -1, 1, 2).setValues([['', '']]).removeCheckboxes()
    spreadsheet.toast('Timesheet: Approved', 'Email Sent', 20);
  }
  else
  {
    checkboxRange.uncheck()
    spreadsheet.toast('Employee must endorse their own timesheet before you can approve it.', 'Warning', 20);
  }
}

/**
* This function sets up the the reminder email trigger when no arguments are sent to it.
* When the function has received an argument, the unapproved email trigger is initiated.
*
* @author Jarren Ralf
*/
function autoSendEmail()
{
  const spreadsheet = SpreadsheetApp.getActive()
  var year, month, startDay, endDay, emailDay, reminderDay;
  [year, month, startDay, endDay] = determinePayPeriod();
  [,, emailDay, reminderDay] = generateDates(year, month, startDay, endDay, spreadsheet);
  
  if (arguments[0]) // Vacation pay requested **The second argument will be true
    ScriptApp.newTrigger("sendEmail_UnapprovedTimesheet_withVacationPay").timeBased().onMonthDay(emailDay.getDate()).atHour(9).create();
  else if (arguments.length === 1) // Otherwise, NO vacation pay because the second argument is false
    ScriptApp.newTrigger("sendEmail_UnapprovedTimesheet").timeBased().onMonthDay(emailDay.getDate()).atHour(9).create();
  else // If someone runs the function manually, this will turn on the reminder email trigger
    ScriptApp.newTrigger("sendEmail_Reminder").timeBased().onMonthDay(reminderDay.getDate()).atHour(9).create();
}

/**
* Calculates Easter in the Gregorian/Western (Catholic and Protestant) calendar 
* based on the algorithm by Oudin (1940) from http://www.tondering.dk/claus/cal/easter.php
*
* @param {Number} year : The current year
* @returns {Number[]} The month and the day of Good Friday
*/
function calculateGoodFriday(year)
{
	var f = Math.floor,
		// Golden Number - 1
		G = year % 19,
		C = f(year / 100),
		// related to Epact
		H = (C - f(C / 4) - f((8 * C + 13)/25) + 19 * G + 15) % 30,
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
* This function runs three triggers at the start of each pay period
*
* @author Jarren Ralf
*/
function createEmployeeTriggers()
{
  ScriptApp.newTrigger('installedOnEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create()
  
  ScriptApp.newTrigger("deleteReminderEmailTrigger").timeBased().onMonthDay(1).atHour(3).create();
  ScriptApp.newTrigger("deleteReminderEmailTrigger").timeBased().onMonthDay(16).atHour(3).create();

  ScriptApp.newTrigger("resetTimesheet").timeBased().onMonthDay(1).atHour(4).create();
  ScriptApp.newTrigger("resetTimesheet").timeBased().onMonthDay(16).atHour(4).create();
  
  ScriptApp.newTrigger("autoSendEmail").timeBased().onMonthDay(1).atHour(5).create();  // Create Reminder email Trigger
  ScriptApp.newTrigger("autoSendEmail").timeBased().onMonthDay(16).atHour(5).create(); // Create Reminder email Trigger

  resetTimesheet(); // Reset the timesheet with the current pay period
  autoSendEmail();  // Turn on the trigger to send the reminder email
}

/**
* This function runs four triggers at the start of each pay period. The fifth trigger runs at the start of every month.
*
* @author Jarren Ralf
*/
function createSupervisorTriggers()
{
  var protection, sheetName;

  SpreadsheetApp.getActive().getSheets()
    .map(sheet => {
      sheetName = sheet.getSheetName()

      if (sheetName !== 'Timesheet') // Protect every sheet that is not the timesheet
      {
        protection = (sheetName === 'Supervisor' || sheetName === 'Timesheet_EmailCopy') ? // Also hide the Supervisor and Timesheet_EmailCopy sheet
          sheet.hideSheet().protect().addEditor(Session.getEffectiveUser()) : sheet.protect().addEditor(Session.getEffectiveUser())
          
        protection.removeEditors(protection.getEditors());
      }
      else // Leave the Timesheet unprotected except for 6 ranges
      {
        sheet.getRangeList(['H1:H1', 'H10:H10', 'D27:D27', 'G27:G27', 'J1:J']).getRanges()
          .map(range => {
            protection = range.protect().addEditor(Session.getEffectiveUser());
            protection.removeEditors(protection.getEditors())
          })
      }
    })

  ScriptApp.newTrigger("setHolidaysAndPayPeriodsAnnually").timeBased().onMonthDay(1).atHour(2).create();
}

/**
* This function determines which is the current Pay Period, based on the current year, month and day.
* The period is either 1-15 or 16-[end of month]. 
*
* @return {[Number, Number, Number, Number, Number, Date]} An array containing the year, month, startDay, endDay, firstDayOfWeek, and today related to the current Pay Period
* @author Jarren Ralf
*/
function determinePayPeriod()
{
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth();
  var startDay, endDay, firstDayOfWeek;

  //Check which pay period we are in, (either 1-15 or 16-[end of month]) and then set the appropriate values
  (today.getDate() < 16) ? (startDay = 1, endDay = 15) : (startDay = 16, endDay = getDaysInMonth(month, year));
  firstDayOfWeek = new Date(year, month, startDay).getDay();
  
  return [year, month, startDay, endDay, firstDayOfWeek, today];
}

/**
 * Deletes the reminder email trigger.
 *
 * @author Jarren Ralf
 */
function deleteReminderEmailTrigger()
{
  deleteTriggers("sendEmail_Reminder");
}

/**
 * Deletes all of the triggers and removes all of the protections.
 *
 * @author Jarren Ralf
 */
function deleteAllTriggersAndRemoveAllProtections()
{
  ScriptApp.getProjectTriggers().map(trigger => ScriptApp.deleteTrigger(trigger))

  [SpreadsheetApp.ProtectionType.RANGE, SpreadsheetApp.ProtectionType.SHEET].map(protectionType => 
    SpreadsheetApp.getActive().getProtections(protectionType).map(protection => (protection.canEdit()) ? protection.remove() : '')
  )
}

/**
 * This function either deletes all of the triggers associated with a given trigger handles.
 *
 * @param {String[]} triggerHandle : The name of the functions to delete the associated triggers for
 * @author Jarren Ralf
 */
function deleteTriggers(...triggerHandles)
{
  ScriptApp.getProjectTriggers().map(trigger => (triggerHandles.includes(trigger.getHandlerFunction())) ? ScriptApp.deleteTrigger(trigger) : '')
}

/**
 * Deletes the unapproved timesheet email triggers.
 *
 * @author Jarren Ralf
 */
function deleteUnapprovedTimesheetEmailTrigger()
{
  deleteTriggers("sendEmail_UnapprovedTimesheet", "sendEmail_UnapprovedTimesheet_withVacationPay");
}

/**
* This is a function that generates the payPeriod, payDay, timesheet-emailDay, and reminder-emailDay for a 
* given year, month, and pay period start and end days.
*
* @param   {Number}       year     : The chosen year
* @param   {Number}      month     : The chosen month
* @param   {Number}     startDay   : The start day of the period
* @param   {Number}      endDay    : The end day of the period
* @param {Spreadsheet} spreadsheet : The active spreadsheet
* @return {[String, Date, Date, Date]} An array containing the payPeriod, payDay, emailDay, and reminderDay
* @author Jarren Ralf
*/
function generateDates(year, month, startDay, endDay, spreadsheet)
{  
  const SUNDAY = 0, MONDAY = 1, TUESDAY = 2, THURSDAY = 4, FRIDAY = 5, SATURDAY = 6,
        FEBRUARY = 2, MARCH = 2, APRIL =  3, OCTOBER = 10, NOVEMBER = 1;
  const timezone = spreadsheet.getSpreadsheetTimeZone();
  const format = 'MM/dd/yyyy';
  const endOfPayPeriod = new Date(year, month, endDay);
  const payPeriodString = Utilities.formatDate(new Date(year, month, startDay), timezone, format) + ' - ' + 
                          Utilities.formatDate(endOfPayPeriod, timezone, format);
  const dayOfWeek = endOfPayPeriod.getDay(); // Get the day of week for the end of period
  var is_EmailDay_EffectedByHoliday = false, is_ReminderDay_EffectedByHoliday = false;
  
  // Check if it is a holiday that will effect the pay period dates and then make the relevant changes
  if (month == FEBRUARY && startDay == 1 && new Date(year, month, getMonday(3, month, year)).getDate() === 15) // Family Day
    dayOfWeek = 0, endDay -= 1;
  else if (month == OCTOBER && startDay == 1) // Thanksgiving Day
  {
    if (dayOfWeek == THURSDAY)
      is_ReminderDay_EffectedByHoliday = true;
    else if (dayOfWeek > MONDAY && dayOfWeek < THURSDAY)
      is_EmailDay_EffectedByHoliday = true;
  }
  else if (month == NOVEMBER && startDay == 1) // Remembrance Day
  {
    const remembranceDay = new Date(year, month, getHoliday(year, month));
    const dayOfWeek_RemembranceDay  = remembranceDay.getDay();
    const dayOfMonth_RemembranceDay = remembranceDay.getDate();
    
    if (dayOfWeek_RemembranceDay == TUESDAY || (dayOfWeek_RemembranceDay == FRIDAY && dayOfMonth_RemembranceDay == 10) || (dayOfWeek_RemembranceDay == MONDAY && dayOfMonth_RemembranceDay == 12))
      is_ReminderDay_EffectedByHoliday = true;
    else if (dayOfWeek_RemembranceDay > TUESDAY && dayOfWeek_RemembranceDay < SATURDAY)
      is_EmailDay_EffectedByHoliday = true;
  }
  else if ((month == MARCH + 1 && startDay == 16) || (month == APRIL + 1 && startDay == 1)) // Good Friday
  {
    var MMM, DD;
    [MMM, DD] = calculateGoodFriday(year);

    const LAST_DAY_IN_PAY_PERIOD = (MMM == APRIL) ? 15 : 31;
    const  EARLIEST_REMINDER_DAY = LAST_DAY_IN_PAY_PERIOD - 5;
    
    // Check if the month is a match for the holiday and if the days might effect the pay period dates
    if ( month - 1 == MMM && (DD >= EARLIEST_REMINDER_DAY && DD <= LAST_DAY_IN_PAY_PERIOD) )
    {    
      if (DD == EARLIEST_REMINDER_DAY)
        is_ReminderDay_EffectedByHoliday = true;
      else if (DD == LAST_DAY_IN_PAY_PERIOD || dayOfWeek == SUNDAY || dayOfWeek == SATURDAY)
        endDay -= 1;
      else if (DD > EARLIEST_REMINDER_DAY && DD < LAST_DAY_IN_PAY_PERIOD)
        is_EmailDay_EffectedByHoliday = true;
    }
  }

  // If the end of pay period falls on a weekend, the pay day needs to roll back to the previous friday
  if (dayOfWeek == SATURDAY)
    endDay -= 1;
  else if (dayOfWeek == SUNDAY)
    endDay -= 2;
  
  const payDate = new Date(year, month, endDay)
  const payWeekDay = payDate.getDay();
  
  if (is_EmailDay_EffectedByHoliday)
    endDay -= (payWeekDay - 2 >= TUESDAY) ? 1 : 3;
  else if (payWeekDay - 2 <= SUNDAY)
    endDay -= 2;
  
  // Send the timesheet email the chosen number of business days before pay day
  endDay -= 2;
  const emailDate = new Date(year, month, endDay, 10)
  
  // If the emailDay is not effected by a holiday
  if (!is_EmailDay_EffectedByHoliday)
  {   
    if (is_ReminderDay_EffectedByHoliday)
      endDay -= (payWeekDay - 3 >= TUESDAY) ? 1 : 3; // If the reminder day is effected, subtract an extra business day
    else if (payWeekDay - 3 == SUNDAY) // The reminder day falls on a weekend, therefore subtract two business days
      endDay -= 2;
  }
  
  // Send the reminder email one business day before the timesheet needs to be submitted
  endDay -= 1;
  const reminderDate = new Date(year, month, endDay)
  
  return [payPeriodString, payDate, emailDate, reminderDate];
}

/**
* This function sends the supervisor an email asking for approval of hours and turns on the automatic trigger to 
* send the unnapproved timesheet of the employee two business days before the pay day.
*
* @param {Spreadsheet} spreadsheet : The active spreadsheet.
* @param    {Sheet}      sheet     : The active sheet.
* @param    {Range}  checkboxRange : The active range.
* @author Jarren Ralf
*/
function getApproval(spreadsheet, sheet, checkboxRange)
{
  const isVacationPayRequested = sheet.getRange(8, 2).getBackground() === '#f4c7c3'; 
  sheet.getRange(28, 8).setFormula('=EmployeeSignature') // Add the employees signature to the bottom of the sheet
  var year, month, startDay, endDay, today;
  
  [year, month, startDay, endDay,_,today] = determinePayPeriod();
  [,,emailDate] = generateDates(year, month, startDay, endDay, spreadsheet);
  
  if (today > emailDate) // The timesheet is late (the email date has passed)
  {
    (isVacationPayRequested) ? sendEmail_UnapprovedTimesheet_withVacationPay() : sendEmail_UnapprovedTimesheet(); 
    spreadsheet.toast('Sending Timesheet directly to the Payroll Manager', 'Email Sent', 20);
  }
  else
  {
    sendEmail_GetApproval() // Send the supervisor the email asking for approval of hours
    autoSendEmail(isVacationPayRequested) // Turn on the trigger for sending the unnapproved email
    spreadsheet.toast('Requesting approval from your supervisor', 'Email Sent', 20);
  }

  checkboxRange.offset(0, -1, 1, 2).setValues([['', '']]).removeCheckboxes()
}

/**
 * This function converts the given sheet into a pdf BLOB object.
 * 
 * @license MIT
 * 
 * © 2020 xfanatical.com. All Rights Reserved.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet
 * @param    {sheet}       sheet    : The sheet object that will be converted into a pdf file.
 * @param    {String}      name     : The name of the timesheet attachment
 * @return The packing slip sheet as a BLOB object that will eventually get converted to pdf document that will be attached to an email sent to the customer
 * @author Jason Huang
 */
function getTimesheetPDF(spreadsheet, sheet, name)
{
  // A credit to https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
  // these parameters are reverse-engineered (not officially documented by Google)
  // they may break overtime.
  const exportUrl = spreadsheet.getUrl().replace(/\/edit.*$/, '') + '/export?'
      + 'exportFormat=pdf&format=pdf&size=LETTER&portrait=true&fitw=true&top_margin=0.75&bottom_margin=0.75&left_margin=0.25&right_margin=0.25'           
      + '&sheetnames=false&printtitle=false&pagenum=UNDEFINED&gridlines=false&fzr=TRUE&gid=' + sheet.getSheetId();

  var response;

  for (var i = 0; i < 5; i++)
  {
    response = UrlFetchApp.fetch(exportUrl, {
      contentType: 'application/pdf',
      muteHttpExceptions: true,
      headers: { 
        Authorization: 'Bearer ' +  ScriptApp.getOAuthToken(),
      },
    })

    if (response.getResponseCode() === 429)
      Utilities.sleep(3000) // printing too fast, retrying
    else
      break;
  }
  
  if (i === 5)
    throw new Error('Printing failed. Too many sheets to print.');
  
  return response.getBlob().setName(name)
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
* This function calculates the day that New Years Day, Canada Day, Remembrance Day, and Christmas Day, is observed on for the giving year and month. 
*
* @param  {Number}  year The given year
* @param  {Number} month The given month
* @return {Number}   day The day of the Holiday for the particular year and month
* @author Jarren Ralf
*/
function getHoliday(year, month)
{
  const JANUARY = 0, JULY = 6, NOVEMBER = 10, DECEMBER = 11, SUNDAY = 0, SATURDAY = 6;
  
  if (month == JANUARY || month == JULY || month == DECEMBER) // New Years Day or Canada Day or Christmas Day
  {
    const holiday = (month == DECEMBER) ? new Date(year, month, 25) : new Date(year, month);
    const dayOfWeek = holiday.getDay();
    var day = (dayOfWeek == SATURDAY) ? holiday.getDate() + 2 : ( (dayOfWeek == SUNDAY) ? holiday.getDate() + 1 : holiday.getDate() ); // Rolls over to the following Monday
  }
  else if (month == NOVEMBER) // Remembrance Day
  {
    const holiday = new Date(year, month, 11);
    const dayOfWeek = holiday.getDay();
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
  const firstDayOfMonth = new Date(year, month).getDay();
  
  if (isVictoriaDay)
    n = (firstDayOfMonth % 6 < 2) ? 4 : 3; // Corresponds to the Monday preceding May 25th 
  
  return ((8 - firstDayOfMonth) % 7) + 7*n - 6;
}

/**
* This function gets the spreadsheets URL.
*
* @return {Spreadsheet} spreadsheet : The active spreadsheet
* @return    {String}       url     : The url of the current spreadsheet
* @author Jarren Ralf
*/
function getSheetUrl(spreadsheet)
{
  return spreadsheet.getUrl() + '#gid=' + spreadsheet.getSheetByName('Timesheet').getSheetId()
}

/**
* This function checks if the supervisor's signature is found on the timesheet, which signifies approval.
*
* @param {Spreadsheet} spreadsheet : The active spreadsheet
* @return {Boolean} Whether the timesheet is approved or not.
* @author Jarren Ralf
*/
function isUnapproved(spreadsheet)
{
  return spreadsheet.getSheetByName('Timesheet').getRange(27, 4).getFormula() === '';
}

/**
* This function checks if the givenn umber is even or not.
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
* This function checks if the given number represents a day of the week or not.
*
* @param  {Number} num The given number [0, 6] representing the days of the week
* @return {Boolean} Whether the given "day" is a week day or not
* @author Jarren Ralf
*/
function isWeekDay(day)
{
  return day != 0 && day != 6
}

/**
* This function clears the appropriate cells in the timesheet then calculates and sets the correct days for 
* the pay period. It also fills in the default hours of the employee as Mon-Friday 9-5, as well as printing
* the pay period on the sheet.
*
* @param {Spreadsheet} spreadsheet : The active spreadsheet.
* @author Jarren Ralf
*/
function resetTimesheet()
{  
  const spreadsheet = SpreadsheetApp.getActive()
  const WEEK_DAYS = ['Sunday', 'Monday','Tuesday', 'Wednesday','Thursday', 'Friday' ,'Saturday'];
  const sheet = spreadsheet.getSheetByName('Timesheet');
  const formats = new Array(16).fill(['@', '@', 'h:mm am/pm', 'h:mm am/pm', "0", '@', '#.00']);
  const values = [], backgrounds = []
  var year, month, startDay, endDay, firstDayOfWeek, payPeriod, dayOfWeek;
  [year, month, startDay, endDay, firstDayOfWeek] = determinePayPeriod();
  [payPeriod] = generateDates(year, month, startDay, endDay, spreadsheet);

  sheet.getRange(1, 5, 1, 4).setFontColors([['#274e13', '#274e13', '#980000', '#980000']]).setFontSizes([[10, 20, 10, 20]]).setFontWeights([['bold', 'normal', 'bold', 'normal']])
    .setFontFamily('Arial').setHorizontalAlignments([['right', 'left', 'right', 'left']]).setVerticalAlignments([['bottom', 'bottom', 'middle', 'bottom']]).setWrap(true)
    .setValues([['Get Timesheet Approved', '', 'Approve Timesheet', '']])
  sheet.getRangeList(['F1:F1', 'H1:H1']).insertCheckboxes()
  sheet.getRangeList(['E8:E8', 'H8:H8', 'D27:D27', 'H28:H28']).clearContent();

  for (var i = 0; i < 8; i++)
    backgrounds.push(['white', 'white', 'white', 'white', 'white', 'white', 'white'], ['#e5edfc', '#e5edfc', '#e5edfc', '#e5edfc', '#e5edfc', '#e5edfc', '#e5edfc'])
  
  // Set the days, start time and end time for the current pay period in the appropriate columns along with the formulas
  for (var i = 0; i < endDay - startDay + 1; i++, firstDayOfWeek++)
  {
    dayOfWeek = firstDayOfWeek % 7; // The day of the week represented by a number 0-6
    values.push([startDay + i, WEEK_DAYS[dayOfWeek], (isWeekDay(dayOfWeek)) ? '9:00 AM' : '', (isWeekDay(dayOfWeek)) ? '5:00 PM' : '', '', '', ''])
  }

  if (values.length < 16)
    values.push(...new Array(16 - values.length).fill(['', '', '', '', '', '', '']))
  
  sheet.getRange(4, 4).setValue(payPeriod).offset(7, -2, 16, 7).setNumberFormats(formats).setBackgrounds(backgrounds).setValues(values);

  deleteUnapprovedTimesheetEmailTrigger();
}

/**
 * This function send's the appropriate message based on the htmlTemplate that is passed to the function.
 * 
 * @param {HtmlOutput} htmlTemplate : The html file that contains the email template.
 * @author Jarren Ralf
 */
function sendEmail(htmlTemplate)
{
  const spreadsheet = SpreadsheetApp.getActive()
  const templateHtml = HtmlService.createTemplateFromFile(htmlTemplate);
  const timesheet = spreadsheet.getSheetByName('Timesheet_EmailCopy')
  const timesheetValues = timesheet.getSheetValues(2, 4, 5, 1);
  const supervisorValues = spreadsheet.getSheetByName('Supervisor').getSheetValues(1, 2, 2, 1);
  var year, month, startDay, endDay, payPeriod; 
  [year, month, startDay, endDay] = determinePayPeriod();
  [payPeriod] = generateDates(year, month, startDay, endDay, spreadsheet)

  templateHtml.payPeriod      =  payPeriod;
  templateHtml.employeeName   =  timesheetValues[0][0].split(' ', 1)[0]; // Employee's first name
  templateHtml.managerName    =  timesheetValues[3][0].split(' ', 1)[0]; // Manager's first name
  templateHtml.supervisorName = supervisorValues[0][0].split(' ', 1)[0]; // Supervisor's first name

  switch (htmlTemplate)
  {
    case 'GetApproval':
      
      templateHtml.sheetURL = getSheetUrl(spreadsheet);
      
      MailApp.sendEmail({        to: supervisorValues[1][0], // Supervisor's email
                            replyTo: timesheetValues[1][0],  // Employee's email
                               name: timesheetValues[0][0],  // Employee's full name
                            subject: timesheetValues[0][0] + '\'s timesheet approval for pay period ' + payPeriod, 
                           htmlBody: templateHtml.evaluate().getContent()
      });
      break;
    case 'HoursApproval':
    case 'HoursApproval_withVacationPay':

      templateHtml.employeeEmail = timesheetValues[1][0];

      MailApp.sendEmail({          to: timesheetValues[4][0],  // Manager's email
                              replyTo: supervisorValues[1][0], // Supervisor's email
                                 name: supervisorValues[0][0], // Supervisor's full name
                                   cc: supervisorValues[1][0] + ',' + timesheetValues[1][0], // Supervisor & Employee's email
                              subject: 'APPROVED: ' + timesheetValues[0][0] + '\'s timesheet for pay period ' + payPeriod, 
                             htmlBody: templateHtml.evaluate().getContent(),
                          attachments: getTimesheetPDF(spreadsheet, timesheet, 
                            'Timesheet_' + timesheetValues[0][0].toString().replace(/ /g, "_") + '_' + payPeriod.toString().replace(" - ", "_") + '.pdf')
      });
      
      spreadsheet.toast('Approved Timesheet Attached', 'Email Sent', 20);
      break;
    case 'HoursApproval_NoAttachment':
      
      MailApp.sendEmail({       to: timesheetValues[4][0],  // Manager's email
                           replyTo: supervisorValues[1][0], // Supervisor's email
                              name: supervisorValues[0][0], // Supervisor's full name
                                cc: supervisorValues[1][0] + ',' + timesheetValues[1][0], // Supervisor & Employee's email
                           subject: 'APPROVED: ' + timesheetValues[0][0] + '\'s timesheet for pay period ' + payPeriod, 
                          htmlBody: message});
      
      spreadsheet.toast('Since the timesheet is past due, you have just sent a brief timesheet approval email with no attachment.', 'Email Sent', 20);
      break;
    case 'Reminder':

      templateHtml.sheetURL = getSheetUrl(spreadsheet);
      
      MailApp.sendEmail({       to: timesheetValues[1][0], // Employee's email
                              name: "TIMESHEET REMINDER",
                           subject: 'REMINDER: Get your timesheet approved for pay period ' + payPeriod, 
                          htmlBody: templateHtml.evaluate().getContent()});
      break;
    case 'UnapprovedTimesheet':
    case 'UnapprovedTimesheet_withVacationPay':

      if (isUnapproved(spreadsheet)) // Only send the unnapproved email if the supervisor signature is missing from the timesheet
      {
        templateHtml.employeeEmail = timesheetValues[1][0];
        
        MailApp.sendEmail({          to: timesheetValues[4][0],  // Manager's email
                                replyTo: timesheetValues[1][0],  // Employee's email
                                   name: timesheetValues[0][0],  // Employee's full name
                                     cc: supervisorValues[1][0] + ',' + timesheetValues[1][0], // Supervisor & Employee's email
                                subject: 'Timesheet for pay period ' + payPeriod, 
                               htmlBody: templateHtml.evaluate().getContent(), 
                            attachments: getTimesheetPDF(spreadsheet, timesheet, 
                              'Timesheet_' + timesheetValues[0][0].toString().replace(/ /g, "") + '_' + payPeriod.toString().replace(" - ", "-") + '.pdf')
        });

        spreadsheet.toast('Unapproved Timesheet Attached', 'Email Sent', 20);
      }

      break;
  }
}

/**
 * This function sends the GetApproval email from the employee to the supervisor.
 * 
 * @author Jarren Ralf
 */
function sendEmail_GetApproval()
{
  sendEmail('GetApproval')
}

/**
 * This function sends the HoursApproval email from the supervisor to the manager.
 * 
 * @author Jarren Ralf
 */
function sendEmail_HoursApproval()
{
  sendEmail('HoursApproval')
}

/**
 * This function sends the HoursApproval_withVacationPay email from the supervisor to the manager.
 * 
 * @author Jarren Ralf
 */
function sendEmail_HoursApproval_withVacationPay()
{
  sendEmail('HoursApproval_withVacationPay')
}

/**
 * This function sends the HoursApproval_NoAttachment email from the supervisor to the manager but the timesheet is not attached.
 * 
 * @author Jarren Ralf
 */
function sendEmail_HoursApproval_NoAttachment()
{
  sendEmail('HoursApproval_NoAttachment')
}

/**
 * This function sends the Reminder email from the employees google email to employees prefered email.
 * 
 * @author Jarren Ralf
 */
function sendEmail_Reminder()
{
  sendEmail('Reminder')
}

/**
 * This function sends the UnapprovedTimesheet email from the employee to the manager.
 * 
 * @author Jarren Ralf
 */
function sendEmail_UnapprovedTimesheet()
{
  sendEmail('UnapprovedTimesheet')
}

/**
 * This function sends the UnapprovedTimesheet_withVacationPay email from the employee to the manager.
 * 
 * @author Jarren Ralf
 */
function sendEmail_UnapprovedTimesheet_withVacationPay()
{
  sendEmail('UnapprovedTimesheet_withVacationPay')
}

/**
 * This function sets the Holidays, as well as the pay periods and relevant dates on the appropriate sheets.
 * 
 * @author Jarren Ralf
 */
function setHolidaysAndPayPeriods()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const timezone = spreadsheet.getSpreadsheetTimeZone();
  const format = 'MM/dd/yyyy';
  const JAN =  0, FEB =  1, MAY =  4, JUL =  6, AUG =  7, SEP =  8, OCT =  9, NOV = 10, DEC = 11;
  const YEAR = new Date().getFullYear(); // An integer corresponding to today's year
  const WEEK_DAYS = ['Sunday', 'Monday','Tuesday', 'Wednesday','Thursday', 'Friday' ,'Saturday'];
  const payPeriods = [["Pay Period", "Pay Day", "Email Day", "Reminder Day"]];
  var MMM, DD;
  [MMM, DD] = calculateGoodFriday(YEAR);

  const newYearsDay     = new Date(YEAR, JAN, getHoliday(YEAR, JAN));
  const familyDay       = new Date(YEAR, FEB, getMonday(3, FEB, YEAR));
  const goodFriday      = new Date(YEAR, MMM, DD);
  const victoriaDay     = new Date(YEAR, MAY, getMonday(0, MAY, YEAR, 1));
  const canadaDay       = new Date(YEAR, JUL, getHoliday(YEAR, JUL));
  const bcDay           = new Date(YEAR, AUG, getMonday(1, AUG, YEAR));
  const labourDay       = new Date(YEAR, SEP, getMonday(1, SEP, YEAR));
  const thanksgivingDay = new Date(YEAR, OCT, getMonday(2, OCT, YEAR));
  const remembranceDay  = new Date(YEAR, NOV, getHoliday(YEAR, NOV));
  const christmasDay    = new Date(YEAR, DEC, getHoliday(YEAR, DEC));

  const holidays = [["Holidays",                                                                                                                   "", ""],
                    ["https://www2.gov.bc.ca/gov/content/employment-business/employment-standards-advice/employment-standards/statutory-holidays", "", ""],
                    ["Name",          "Date (Observed)", "Day of Week"],
                    ["New Year's Day",   newYearsDay,     WEEK_DAYS[newYearsDay.getDay()]],
                    ["Family Day",       familyDay,       WEEK_DAYS[familyDay.getDay()]],
                    ["Good Friday",      goodFriday,      WEEK_DAYS[goodFriday.getDay()]],
                    ["Victoria Day",     victoriaDay,     WEEK_DAYS[victoriaDay.getDay()]],
                    ["Canada Day",       canadaDay,       WEEK_DAYS[canadaDay.getDay()]],
                    ["BC Day",           bcDay,           WEEK_DAYS[bcDay.getDay()]],
                    ["Labour Day",       labourDay,       WEEK_DAYS[labourDay.getDay()]],
                    ["Thanksgiving Day", thanksgivingDay, WEEK_DAYS[thanksgivingDay.getDay()]],
                    ["Remembrance Day",  remembranceDay,  WEEK_DAYS[remembranceDay.getDay()]],
                    ["Christmas Day",    christmasDay,    WEEK_DAYS[christmasDay.getDay()]]];

  for (var i = 0; i < 24; i++)
    payPeriods.push((isEven(i)) ? 
      generateDates(YEAR, i/2, 1, 15, spreadsheet).map(date => (typeof date !== 'string') ? Utilities.formatDate(date, timezone, format) : date) : 
      generateDates(YEAR, (i - 1)/2, 16, getDaysInMonth((i - 1)/2, YEAR), spreadsheet).map(date => (typeof date !== 'string') ? Utilities.formatDate(date, timezone, format) : date))

  spreadsheet.getSheetByName('Holidays').getRange(1, 1, 13, 3).setValues(holidays);
  spreadsheet.getSheetByName('Pay Periods').getRange(1, 1, 25, 4).setValues(payPeriods);
}

/**
* This function is a quick work around to set a yearly trigger. The trigger runs a function every month. 
* That function only executes when the month is January. 
* In this case specifically, it sets all of the dates on the Pay Periods sheet of this spreadsheet.
*
* @author Jarren Ralf
*/
function setHolidaysAndPayPeriodsAnnually()
{
  if (new Date().getMonth() === 0)
    setHolidaysAndPayPeriods();
}