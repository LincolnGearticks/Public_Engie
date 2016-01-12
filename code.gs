//These two functions allow for someone to correct the data and then redo the entries without sending an e-mail
//Remember to re-star the e-mails with images that you want added
function noEmailYesterday() { //as if it was run yesterday (to correct for a mistake in yesterday's doc)
  noEmailToday(yesterdaysDate_()); //pretend that it is yesterday
}
function noEmailToday(date) { //as if it was run today (to correct for a mistake in today's doc)
  makeEntryToday(true, date);
}
function makeEntryYesterday() {
  makeEntryToday(false, yesterdaysDate_());
}
function makeEntryToday(avoidEmail, date) { //the main routine
  var log = ''; //a record of what happened (will get Logger.log'd and will get e-mailed if not doing a manual run)
  var thrownError = null; //will store an error, if it gets thrown (so it can be caught, handled, and rethrown)
  try {
    const TIMEZONE = /* Enter value here */; //timezone to use for all displayed dates
    //Folders to put entry docs in (using IDs from Google Drive URLs)
    const ENTRY_FOLDER = /* Enter value here */;
    const OUTREACH_FOLDER = /* Enter value here */;
    //The spreadsheet with the entry data
    const RESPONSES_SHEET = /* Enter value here */;
    //Addresses to e-mail documents to
    const EMAIL_TO = /* Enter value here */;
    const LOG_EMAIL_TO = /* Enter value here */; //for script logs
    //Image for header
    const HEADER_ICON_FILE = /* Enter value here */;
    //Header titles
    const NORMAL_HEADER = /* Enter value here */;
    const OUTREACH_HEADER = /* Enter value here */;

    const OUTREACH = 'Outreach'; //name of Outreach task group
    const HEADER_ICON = HEADER_ICON_FILE && DriveApp.getFileById(HEADER_ICON_FILE).getBlob();
    //Spreadsheet constants
    const DATA_START_ROW = 2;
    const DATA_START_COLUMN = 1;
    //Indices of each element of today's data
    const AUTHOR_DATA = 1;
    const START_TIME = 2;
    const END_TIME = 3;
    //Keys for entry element
    const GROUP = 0;
    const TASK = 1;
    const REFL = 2;
    const KEY = 3;
    const DATA = 4;
    const AUTHOR_ENTRY = 5;
    //Keys for image object
    const CAPTION = 0;
    const ATTACHMENT = 1;
    const FULL_SIZE = 2;
    const OUTREACH_PIC = 3;
    //Names for headers
    const KEY_HEADING = 'Key Learning';
    const REFLECTIONS_HEADING = 'Reflections';
    const DATA_HEADING = 'Data';
    //Picture entry constants
    const QPICW = 295; //quarter page picture width
    const QPICH = 320; //quarter page picture height
    const FPICW = 605; //full page picture width
    const FPICH = 760; //full page picture height
    const QCELLW = 235; //quarter page cell width
    const FCELLW = 465;  //full page cell width
    //Header image size
    const ICON_SIZE = 50; //pixels high and wide

    //Set up the styles used in the documents
    const HEADER_STYLE = {}; //used in header and for task group names
    HEADER_STYLE[DocumentApp.Attribute.BOLD] = true;
    HEADER_STYLE[DocumentApp.Attribute.FONT_SIZE] = 14;
    HEADER_STYLE[DocumentApp.Attribute.SPACING_BEFORE] = 0;
    HEADER_STYLE[DocumentApp.Attribute.SPACING_AFTER] = 0;
    const MEDIUM_HEADER_STYLE = clone(HEADER_STYLE); //used for task names
    MEDIUM_HEADER_STYLE[DocumentApp.Attribute.FONT_SIZE] = 12;
    const SMALL_HEADER_STYLE = clone(HEADER_STYLE); //used for 'Reflections', 'Key Learning', and 'Data' headers
    SMALL_HEADER_STYLE[DocumentApp.Attribute.FONT_SIZE] = 10;
    const REGULAR_STYLE = {}; //used for most of the normal text
    REGULAR_STYLE[DocumentApp.Attribute.BOLD] = false;
    REGULAR_STYLE[DocumentApp.Attribute.FONT_SIZE] = 10;
    const TIME_BLOCK_STYLE = {}; //used for the time block
    TIME_BLOCK_STYLE[DocumentApp.Attribute.FONT_SIZE] = 9;
    TIME_BLOCK_STYLE[DocumentApp.Attribute.SPACING_BEFORE] = 0;
    TIME_BLOCK_STYLE[DocumentApp.Attribute.SPACING_AFTER] = 0;
    TIME_BLOCK_STYLE[DocumentApp.Attribute.LINE_SPACING] = 1;
    const PARA_STYLE = {};
    PARA_STYLE[DocumentApp.Attribute.BOLD] = false;
    PARA_STYLE[DocumentApp.Attribute.FONT_SIZE] = 10;
    PARA_STYLE[DocumentApp.Attribute.SPACING_AFTER] = 0;
    PARA_STYLE[DocumentApp.Attribute.SPACING_BEFORE] = 0;
    PARA_STYLE[DocumentApp.Attribute.LINE_SPACING] = 1;
    PARA_STYLE[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
    const SMALLEST_SIZE = {};
    SMALLEST_SIZE[DocumentApp.Attribute.FONT_SIZE] = 6;

    //Collect the data from the form and process it
    const TODAY = date || new Date(); //will default to the current day, but another can be passed
    log += 'Entry search date: ' + Utilities.formatDate(TODAY, TIMEZONE, 'yyyy-MM-dd') + '\n';
    log += 'Collecting data from spreadsheet\n';
    var todaysData = collectData(); //raw data rows from spreadsheet
    var entries = entriesFromData(todaysData); //data processed into entries

    //Create entry doc and put it in the correct folder
    log += 'Creating normal Doc\n';
    var normalDoc = DocumentApp.create(Utilities.formatDate(TODAY, TIMEZONE, 'yyyy-MM-dd') + ' Engie');
    var normalDocFile = DriveApp.getFileById(normalDoc.getId());
    DriveApp.getFolderById(ENTRY_FOLDER).addFile(normalDocFile);
    DriveApp.getRootFolder().removeFile(normalDocFile);

    var outreachDoc; //the outreach doc (if it needs to be created)
    var normalBody = normalDoc.getBody();
    var outreachBody;
    var body; //the document body to add to (in normal or outreach doc)

    log += 'Creating time block\n';
    todaysData.sort(function(a, b) { //organize time data alphabetically by person's name
      if (a[AUTHOR_DATA] < b[AUTHOR_DATA]) return -1;
      else if (b[AUTHOR_DATA] > a[AUTHOR_DATA]) return 1;
      else return 0;
    });
    var normalData = [], outreachData = [];
    for (var i in todaysData) {
      var entriesForDatum = entriesFromData([todaysData[i]]);
      var outreachEntries = filterByText(entriesForDatum, GROUP, OUTREACH);
      if (outreachEntries.length) outreachData.push(todaysData[i]); //if some of the person's entries were outreach
      if (!entriesForDatum.length || outreachEntries.length != entriesForDatum.length) normalData.push(todaysData[i]); //if some of the person's entries were not outreach or the person submitted no entries (probably a coach)
    }
    function makeTimeBlock(filteredData, body) {
      var timeBlock = body.insertTable(body.getNumChildren() - 1);
      var timeStart, timeEnd, entryRow, entryCell; //start time string, end time string, cell to put team member's time in
      for (var i = 0; i < filteredData.length; i++) { //iterate over form responses from today
        if (!(i % 3)) entryRow = timeBlock.appendTableRow(); //make new row every 3 entries
        timeStart = Utilities.formatDate(new Date(filteredData[i][START_TIME]), TIMEZONE, 'h:mm a');
        timeEnd = Utilities.formatDate(new Date(filteredData[i][END_TIME]), TIMEZONE, 'h:mm a');
        entryCell = entryRow.appendTableCell(filteredData[i][AUTHOR_DATA] + ': ' + timeStart + ' - ' + timeEnd);
        entryCell.setWidth(157).setPaddingTop(0).setPaddingBottom(0);
      }
      timeBlock.setAttributes(TIME_BLOCK_STYLE).setBorderWidth(0);
      if (filteredData.length) body.insertHorizontalRule(body.getNumChildren() - 1); //if there are no entries, there is no need for a separator
    }
    makeTimeBlock(normalData, normalBody);

    var taskGroups = findAllValuesCaseInsensitive(entries, GROUP); //the names of the task groups that have entries for them
    var entriesInGroup, keyTrigger, dataTrigger, taskNames, taskName, entriesForTask, taskTable, dataRow;
    var keepDoc = false; //whether or not any normal entries were received
    var j, k;
    log += 'Writing entries\n';
    for (var i in taskGroups) { //iterate over task groupings (Programming, Mechanical Build, etc.)
      taskGroup = taskGroups[i];
      //Open new doc and create Outreach entry for outreach items
      if (taskGroup == OUTREACH) {
        if (!outreachDoc) { //if we need to create an outreach doc
          log += 'Creating outreach Doc\n';
          outreachDoc = DocumentApp.create(Utilities.formatDate(TODAY, TIMEZONE, 'yyyy-MM-dd') + ' Outreach Engie');
          var outreachDocFile = DriveApp.getFileById(outreachDoc.getId());
          DriveApp.getFolderById(OUTREACH_FOLDER).addFile(outreachDocFile);
          DriveApp.getRootFolder().removeFile(outreachDocFile);
          outreachBody = outreachDoc.getBody();
          makeTimeBlock(outreachData, outreachBody);
        }
        body = outreachBody;
      }
      else {
        body = normalBody;
        keepDoc = true;
      }
      //Add a header for the task group and the key learning table
      body.insertParagraph(body.getNumChildren() - 1, taskGroup).setAttributes(HEADER_STYLE);
      var keyTable = body.insertTable(body.getNumChildren() - 1).setBorderWidth(0);
      var keyRow = keyTable.appendTableRow();
      keyRow.appendTableCell(KEY_HEADING).setAttributes(SMALL_HEADER_STYLE);
      keyTrigger = false; //used to delete key learning table if there are no entries

      entriesInGroup = filterByTextCaseInsensitive(entries, GROUP, taskGroup); //get all the entries in this task group
      taskNames = [];
      /*Iterate over each entry in a task group
      Extract names of sub tasks and add to the key learnings*/
      for (j in entriesInGroup) { //get all the tasks in this task group and output the key learnings
        taskName = entriesInGroup[j][TASK];
        if (indexOfCaseInsensitive(taskNames, taskName) == -1) taskNames.push(taskName);
        if (entriesInGroup[j][KEY]) {
          keyRow = keyTable.appendTableRow();
          keyRow.appendTableCell(entriesInGroup[j][KEY] + ' –' + entriesInGroup[j][AUTHOR_ENTRY]).setAttributes(REGULAR_STYLE).setPaddingTop(0).setPaddingBottom(0).editAsText().setItalic(true);
          keyTrigger = true;
        }
      }
      if (!keyTrigger) body.removeChild(keyTable);

      for (j in taskNames) { //iterate over sub tasks
        taskTable = body.insertTable(body.getNumChildren() - 1).setBorderWidth(0);
        taskTable.appendTableRow().appendTableCell(taskNames[j]).setAttributes(MEDIUM_HEADER_STYLE).setPaddingBottom(0).setPaddingTop(0);
        taskTable.appendTableRow().appendTableCell(REFLECTIONS_HEADING).setAttributes(SMALL_HEADER_STYLE).setPaddingBottom(0).setPaddingTop(0);
        entriesForTask = filterByTextCaseInsensitive(entriesInGroup, TASK, taskNames[j]); //get entries for the task
        //Add all the reflections
        for (k in entriesForTask) {
          var entryText = entriesForTask[k][REFL] + ' –' + entriesForTask[k][AUTHOR_ENTRY];
          if (k != entriesForTask.length - 1) entryText += '\n'; //separate entries with blank lines
          taskTable.appendTableRow().appendTableCell(entryText).setAttributes(PARA_STYLE).setPaddingBottom(0).setPaddingTop(0);
        }
        dataRow = taskTable.appendTableRow();
        dataRow.appendTableCell(DATA_HEADING).setAttributes(SMALL_HEADER_STYLE).setPaddingBottom(0).setPaddingTop(0); //insert 'Data' heading
        dataTrigger = false; //use to delete 'Data' heading if there are no entries
        //Add all the data entries for the subtask
        for (k in entriesForTask) {
          if (entriesForTask[k][DATA]) {
            taskTable.appendTableRow().appendTableCell(entriesForTask[k][DATA] + ' –' + entriesForTask[k][AUTHOR_ENTRY]).setAttributes(PARA_STYLE).setPaddingBottom(0).setPaddingTop(0);
            dataTrigger = true;
          }
        }
        if (!dataTrigger) taskTable.removeChild(dataRow);
      }
      body.insertHorizontalRule(body.getNumChildren() - 1);
    }

    function makeFullWordRegex(word) {
      return new RegExp('(^|\\s)' + word + '(\\s|$)', 'i'); //case-insensitive; requires word to be (at start or following whitespace) and (at end or preceding whitespace)
    }
    const FULL_SEARCH = makeFullWordRegex('full');
    const OUTREACH_SEARCH = makeFullWordRegex('outreach');

    //Go through all starred messages, add to engie notebook
    //Only works for attached images, not inline images
    log += 'Looking for images\n';
    var threads = GmailApp.search('in:inbox is:starred has:attachment').reverse(); //only look for images that are starred (unprocessed) and have attachments, order chronologically
    var emailedPics = []; //array to fill with pics from GMail messages
    var messages = []; //array to fill with messages containing the pictures (to unstar later if the pictures were used)
    var thread, messagesInThread, message, attachments, attachment, subject, messageData;
    for (var i = 0, j, k; i < threads.length; i++) {
      thread = threads[i];
      messagesInThread = thread.getMessages();
      for (j = 0; j < messagesInThread.length; j++) {
        message = messagesInThread[j];
        attachments = message.getAttachments();
        for (k = 0; k < attachments.length; k++) {
          attachment = attachments[k];
          if (begins(attachment.getContentType(), 'image/')) { //don't want to try to interpret PDFs, etc. as images
            //Appearance of 'full' or 'outreach' in subject indicates that we need special handling
            //We are also trying to use the subject line as a place to put captions, so remove the special words
            subject = message.getSubject();
            messageData = [];
            messageData[CAPTION] = capitalize(subject.replace(FULL_SEARCH, '').replace(OUTREACH_SEARCH, '').trim().replace(/  /g, ' '));
            log += 'Got image in e-mail (caption: ' + messageData[CAPTION] + ')\n';
            messageData[ATTACHMENT] = attachment;
            subject = subject.toLowerCase(); //no need to keep track of cases when looking for control words
            messageData[FULL_SIZE] = FULL_SEARCH.test(subject);
            messageData[OUTREACH_PIC] = OUTREACH_SEARCH.test(subject);
            emailedPics.push(messageData);
          }
        }
        messages.push(message);
      }
    }
    //Sort array so full-sized pictures are first
    emailedPics.sort(function(a, b) {
      if (a[FULL_SIZE] > b[FULL_SIZE]) return -1;
      else if (b[FULL_SIZE] < a[FULL_SIZE]) return 1;
      else return 0;
    });

    //Format pictures and insert into appropriate notebooks
    log += 'Inserting images\n';
    var regPicCount = 0; //total number of quarter-sized pictures in regular notebook
    var outPicCount = 0; //total number of quarter-sized pictures in outreach notebook
    var regPictures = false; //whether there are any pictures at all in the regular notebook
    var pic; //the current picture
    //Table and table row to put quarter-size images into in each notebook
    var normalTable, normalLine, outreachTable, outreachLine;
    var currentBody, currentTable, currentLine, currentPicCount, treatAsOutreach; //currentPicCount stores the number of quarter-size pictures so we can figure out when to start a new row or page
    for (var i = 0; i < emailedPics.length; i++) {
      pic = emailedPics[i];
      //Get current elements to be manipulating (so we don't have to keep track of whether we are in outreach doc or not)
      treatAsOutreach = pic[OUTREACH_PIC] && outreachDoc; //put outreach pictures on the main doc if there is no outreach doc
      if (treatAsOutreach) {
        currentBody = outreachBody;
        currentTable = outreachTable;
        currentLine = outreachLine;
        currentPicCount = outPicCount;
      }
      else {
        currentBody = normalBody;
        currentTable = normalTable;
        currentLine = normalLine;
        currentPicCount = regPicCount;
      }
      if (pic[FULL_SIZE]) { //full sized picture
        currentBody.appendPageBreak(); //each picture needs its own page
        var tableCell = currentBody.appendTable().setBorderWidth(0).appendTableRow().appendTableCell().setWidth(FCELLW).setAttributes(PARA_STYLE); //add a table cell to contain the image and caption
        var picCell = tableCell.insertImage(0, pic[ATTACHMENT]); //display image
        var newDims = resizeDims(picCell.getWidth(), picCell.getHeight(), FPICW, FPICH);
        picCell.setWidth(newDims.width).setHeight(newDims.height).getParent().setAlignment(DocumentApp.HorizontalAlignment.CENTER); //resize image
        var textOfMessage = tableCell.insertParagraph(tableCell.getNumChildren() - 1, pic[CAPTION]).setAttributes(PARA_STYLE).editAsText(); //add caption
        textOfMessage.setText(textOfMessage.getText().trim()); //remove extra whitespace
        if (!treatAsOutreach) regPictures = true;
      }
      else { //quarter sized picture
        if (currentPicCount % 4 == 0) { //4 photos per page
          if (!(currentPicCount == 0 && !entries.length)) currentBody.appendPageBreak(); //if there were no entries, the first page break is unnecessary
          currentTable = currentBody.appendTable().setBorderWidth(0);
        }
        if (currentPicCount % 2 == 0) currentLine = currentTable.appendTableRow(); //2 photos per row
        currentPicCount++;
        //Set back modified values
        if (treatAsOutreach) {
          outreachTable = currentTable;
          outreachLine = currentLine;
          outPicCount = currentPicCount;
        }
        else {
          normalTable = currentTable;
          normalLine = currentLine;
          regPicCount = currentPicCount;
        }
        var tableCell = currentLine.appendTableCell().setWidth(QCELLW).setAttributes(PARA_STYLE); //add a table cell to contain the image and caption
        var picCell = tableCell.insertImage(0, pic[ATTACHMENT]); //display image
        var newDims = resizeDims(picCell.getWidth(), picCell.getHeight(), QPICW, QPICH);
        picCell.setWidth(newDims.width).setHeight(newDims.height).getParent().setAlignment(DocumentApp.HorizontalAlignment.CENTER); //resize image
        var textOfMessage = tableCell.insertParagraph(tableCell.getNumChildren() - 1, pic[CAPTION]).setAttributes(PARA_STYLE).editAsText(); //add caption
        textOfMessage.setText(textOfMessage.getText().trim()); //remove whitespace
      }
      tableCell.removeChild(tableCell.getChild(tableCell.getNumChildren() - 1));
    }
    if (regPicCount) regPictures = true;

    const FORM_DATE = Utilities.formatDate(TODAY, TIMEZONE, 'M/d/yyyy'); //formatted date for header
    //Creates a header for the normal/outreach doc
    function createHeader(doc, title, filteredData) {
      log += 'Creating header: ' + title + '\n';
      var header = doc.addHeader().setAttributes(HEADER_STYLE);
      var headerTable = header.appendTable().setBorderWidth(0);
      var headerTableRow = headerTable.appendTableRow();
      //titleCell has the title of the notebook and the tick icon
      var titleCell = headerTableRow.appendTableCell(title);
      titleCell.setWidth(290).setVerticalAlignment(DocumentApp.VerticalAlignment.BOTTOM).setPaddingLeft(0).setPaddingBottom(0);
      var titleParagraph = titleCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      if (HEADER_ICON) titleParagraph.insertInlineImage(0, HEADER_ICON).setWidth(ICON_SIZE).setHeight(ICON_SIZE);
      const START_TIME = minTime(filteredData);
      const END_TIME = maxTime(filteredData);
      var dateCell = headerTableRow.appendTableCell(FORM_DATE + ' ' + START_TIME + ' - ' + END_TIME); //dateCell has the date of the meeting
      dateCell.setWidth(180).setVerticalAlignment(DocumentApp.VerticalAlignment.BOTTOM).setPaddingLeft(0).setPaddingBottom(0);
      dateCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setAttributes(REGULAR_STYLE);
      header.insertHorizontalRule(header.getNumChildren() - 1); //to separate the header from the body when it is printed
      header.getChild(header.getNumChildren() - 1).setAttributes(SMALLEST_SIZE); //minimize the size of extra paragraph element to shrink header
    }
    //Create header for each page of regular entry
    createHeader(normalDoc, NORMAL_HEADER, normalData);
    normalDoc.saveAndClose(); //we are done with the normal doc

    var docsToEmail = [];
    if (keepDoc || regPictures) docsToEmail.push(normalDoc); //if there were entries or pictures, e-mail people
    else { //otherwise get rid of the files
      log += 'Removing unused normal file\n';
      DriveApp.getFolderById(ENTRY_FOLDER).removeFile(normalDocFile);
    }
    //Create header and timeBlock for Outreach Notebook page if necessary
    if (outreachDoc) {
      createHeader(outreachDoc, OUTREACH_HEADER, outreachData);
      outreachDoc.saveAndClose();
      docsToEmail.push(outreachDoc);
    }
    if (!avoidEmail && docsToEmail.length && EMAIL_TO) { //don't e-mail unless there are things to e-mail and people to e-mail them to
      log += 'E-mailing\n';
      emailDocs(docsToEmail);
    }
    if (docsToEmail.length) { //if there were entries (the docs weren't both deleted)
      log += 'Unstarring messages\n';
      for (var message in messages) messages[message].unstar(); //so the images don't get inserted into every subsequent entry
    }
  }
  catch (e) {
    log += e.message + '\n';
    log += 'on line ' + e.lineNumber;
    thrownError = e;
  }
  log = log.trim();
  Logger.log(log);
  if (LOG_EMAIL_TO && !avoidEmail) MailApp.sendEmail(LOG_EMAIL_TO, 'Entry Log', log); //send logs
  if (thrownError) throw thrownError; //so person running the script can see the error

  //Sends out an e-mail with the docs attached as PDFs
  function emailDocs(docs) {
    const NAME = docs[0].getName();
    for (var doc in docs) docs[doc] = docs[doc].getAs(MimeType.PDF);
    MailApp.sendEmail(EMAIL_TO, NAME, '', {
      attachments: docs
    });
  }

  //Resize picture dimensions
  function resizeDims(picWidth, picHeight, newWidth, newHeight) {
    const DIM_RATIO = picWidth / picHeight;
    var newW = newWidth;
    var newH = Math.round(newWidth / DIM_RATIO);
    if (newH > newHeight) { //too tall when scaled to have the correct width -> needs to be restrained by height instead
      newH = newHeight;
      newW = Math.round(newHeight * DIM_RATIO);
    }
    return {
      width: newW,
      height: newH
    };
  }

  //Function to collect data from form's spreadsheet
  // returns array in the following columns:
  //timestamp(0), author(1), starttime(2), endtime(3), task1(4), group1(5), refl1(6),
  //key1(7), data1(8), y/n(9), task2(10), group2(11), refl2(12), key2(13), data2(14), task3(15), group3(16),
  //refl3(17), key3(18), data4(19), task4(20), group4(21), refl4(22), key4(23), data4(24)
  function collectData() {
    if (!RESPONSES_SHEET) return []; //if using the image portion only, there should be no entries
    var ss = SpreadsheetApp.openById(RESPONSES_SHEET);
    var sheet = ss.getSheets()[0]; //data from form
    var timestamps = sheet.getRange(DATA_START_ROW, DATA_START_COLUMN, sheet.getLastRow() - DATA_START_ROW + 1, 1).getValues(); //get the values of the submission timestamps
    var todaysData = [];
    const LAST_COLUMN = sheet.getLastColumn();
    var spreadsheetValues = sheet.getRange(DATA_START_ROW, DATA_START_COLUMN, sheet.getLastRow() - DATA_START_ROW + 1, LAST_COLUMN - DATA_START_COLUMN + 1).getValues();
    var todayIndex = 0; //index in todaysData to insert the data
    for (var i = 0, j; i < timestamps.length; i++) {
      var aDate = new Date(timestamps[i]);  //create a date object from the form's submitted date
      if (Utilities.formatDate(aDate, TIMEZONE, 'yyyyMMdd') == Utilities.formatDate(TODAY, TIMEZONE, 'yyyyMMdd')) { //if the date in the cell is today's date
        todaysData[todayIndex] = [];
        for (j = 0; j < LAST_COLUMN; j++) todaysData[todayIndex][j] = spreadsheetValues[i][j]; //insert the remaining values of the entry
        todayIndex++;
      }
    }
    return todaysData;
  }

  //Create an array from todaysData that puts all entries etc. in same columns:
  //entries columns: task group, task name, reflections, key learning, data, author
  function entriesFromData(todaysData) {
    const AUTHOR = 1;
    const ENTRY_OFFSET = 5; //except between 1 and 2 because of the y/n column
    var ENTRY_STARTS = [4]; //columns containing the task name of an entry (the first column of that entry's data)
    ENTRY_STARTS.push(ENTRY_STARTS[ENTRY_STARTS.length - 1] + ENTRY_OFFSET + 1);
    ENTRY_STARTS.push(ENTRY_STARTS[ENTRY_STARTS.length - 1] + ENTRY_OFFSET);
    ENTRY_STARTS.push(ENTRY_STARTS[ENTRY_STARTS.length - 1] + ENTRY_OFFSET);
    const TASK_OFFSET = 0; //offsets relative to the start of the entry data
    const GROUP_OFFSET = 1;
    const REFL_OFFSET = 2;
    const KEY_OFFSET = 3;
    const DATA_OFFSET = 4;
    var entries = [];
    var subEntry;
    for (var i in todaysData) {
      for (subEntry in ENTRY_STARTS) {
        if (todaysData[i][ENTRY_STARTS[subEntry] + REFL_OFFSET]) { //require a task name and a reflection
          entries.push([
            todaysData[i][ENTRY_STARTS[subEntry] + GROUP_OFFSET],
            capitalize(todaysData[i][ENTRY_STARTS[subEntry] + TASK_OFFSET].trim()),
            todaysData[i][ENTRY_STARTS[subEntry] + REFL_OFFSET].trim().replace(/\n{2,}/g, '\n'), //by removing double newlines, entries don't look split-up
            todaysData[i][ENTRY_STARTS[subEntry] + KEY_OFFSET].trim(),
            todaysData[i][ENTRY_STARTS[subEntry] + DATA_OFFSET].trim(),
            todaysData[i][AUTHOR]
          ]);
        }
      }
    }
    return entries;
  }

  //Functions to generate start time and end times of meetings.
  //takes the array of today's data and returns the meeting times formatted to our time zone
  //minTime finds start of first entry, maxTime finds end of last entry
  function minTime(todaysData) {
    const START_TIME_COLUMN = 2;
    const DEFAULT_DATE = new Date(0);
    var mintime = DEFAULT_DATE; //make sure it is always a valid date
    var startTime;
    for (var i in todaysData) {
      startTime = new Date(todaysData[i][START_TIME_COLUMN]);
      if (mintime == DEFAULT_DATE || startTime < mintime) mintime = startTime; //if we haven't yet found a minimum time or this time is earlier
    }
    return Utilities.formatDate(mintime, TIMEZONE, 'hh:mm a');
  }
  function maxTime(todaysData) {
    const END_TIME_COLUMN = 3;
    const DEFAULT_DATE = new Date(0);
    var maxtime = DEFAULT_DATE;
    var endTime;
    for (var i in todaysData) {
      endTime = new Date(todaysData[i][END_TIME_COLUMN]);
      if (maxtime == DEFAULT_DATE || endTime > maxtime) maxtime = endTime;
    }
    return Utilities.formatDate(maxtime, TIMEZONE, 'hh:mm a');
  }

  //Copies over all keys of an object
  function clone(obj) {
    var target = {};
    for (var i in obj) {
      if (obj.hasOwnProperty(i)) target[i] = obj[i];
    }
    return target;
  }

  //Checks whether a string starts with another string
  function begins(full, sub) {
     return full.substring(0, sub.length) == sub;
  }

  //Capitalizes a string
  function capitalize(str) {
    if (str.length) return str[0].toUpperCase() + str.substring(1);
    else return '';
  }

  //Returns whether two strings are equal, ignoring case
  function equalsCaseInsensitive(thing1, thing2) {
    return thing1.toLowerCase() == thing2.toLowerCase();
  }
  //Selects only the matching values from an array
  function filterByText(elements, matchKey, matchString) {
    var result = [];
    for (var element in elements) {
      element = elements[element];
      if (element[matchKey] == matchString) result.push(element);
    }
    return result;
  }
  //Selects only the matching values from an array, ignoring case
  function filterByTextCaseInsensitive(elements, matchKey, matchString) {
    var result = [];
    for (var element in elements) {
      element = elements[element];
      if (equalsCaseInsensitive(element[matchKey], matchString)) result.push(element);
    }
    return result;
  }
  //Returns the index of an element in an array, ignoring case
  function indexOfCaseInsensitive(arr, element) {
    for (var i = 0; i < arr.length; i++) {
      if (equalsCaseInsensitive(arr[i], element)) return i;
    }
    return -1;
  }
  //Gets all the possible values for a certain key in the elements of an array
  function findAllValuesCaseInsensitive(elements, matchKey) {
    var values = [];
    for (var i in elements) {
      var value = elements[i][matchKey];
      if (indexOfCaseInsensitive(values, value) == -1) values.push(value);
    }
    return values;
  }
}
function yesterdaysDate_() {
  return new Date(new Date().getTime() - 86400000);
}