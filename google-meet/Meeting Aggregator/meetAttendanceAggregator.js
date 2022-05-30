/* 
 * 
 * Description: Takes all google meet attendance files and then aggregates them into one file.
 *
 * Prerequisists: Need all sheets to be in one main folder - this is the Pending Folder
 *                Need a folder for processed filed - this is the processed Folder
 *                Need a sheet where all the data will be stored - this is your sheet URL amd SHEET_TAB_NAME
 */

/*
 * Global Variables that define where to pull and place the data for the Google Sheets
 *
 * @description
 * 
 *  SHEET_URL:
 *      Output Google Sheet where aggregated data will be stored
 *
 *  SHEET_TAB_NAME:
 *      Output Google Sheet Tab where aggregated data will be stored
 *
 *  GOOGLE_MEET_ATTENDANCE_FOLDER_ID:
 *      Folder ID {That is the last part of the URL of the google drive folder https://drive.google.com/drive/u/0/folders/{{Folder_ID}} }
 *      that hosts all the pending CSVs that still need to be processed into the aggregatd data.
 *
 *  PROCESSED_FOLDER_ID:
 *      Folder ID where the processed files will be placed.  USED IF TRASH_FILES == false
 */

const SHEET_URL = "";
const SHEET_TAB_NAME = "";
const GOOGLE_MEET_ATTENDANCE_FOLDER_ID = "";
const PROCESSED_FOLDER_ID = "";

var pending_folder_list = [];
var checked_folder_list = [];
var file_list =[];

/*
 * Function: set_sheet_headers
 *
 * @description
 *      Sets the Headers in the output google sheets.  Headers are:  
 *      File_Name | File_Id | Meeting_Name | Meeting Date | Name | Email | Duration | Time_Joined | Time Exited | Date_Time_Joined| Date_Time_Exited | Duration_Minutes
 */

function set_sheet_headers() {
  var sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(SHEET_TAB_NAME);
  sheet.appendRow(
    ["Original File Name", "File Id", "Meeting Code", "Meeting date", "Name", "Email", "Duration", "Time Joined", "Time Exited", "Date Time Joined", "Date Time Exited", "Duration_Minutes"]);
}

function process_all_folders_to_get_files(id){
  
  var folder = DriveApp.getFolderById(id);
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var childfolders = folder.getFolders();

  var j = file_list.length;
  while(files.hasNext()){
    file = files.next();
    file_list[j] =file.getId();
    j++;
  }

  var i = pending_folder_list.length;
  while(childfolders.hasNext()){
   
    var folder_id = childfolders.next().getId();
    pending_folder_list[i] = folder_id;
    i++;
  }

  if (pending_folder_list.length > 0){

    var item_to_be_processed = pending_folder_list.shift();
    checked_folder_list[checked_folder_list.length] =  item_to_be_processed;
    process_all_folders_to_get_files(item_to_be_processed);

  }else{

    while(file_list.length >0){
      
      var file_to_read = file_list.shift(); 
      importGoogleSheetByFileId(file_to_read);
      var processed_folder = DriveApp.getFolderById(PROCESSED_FOLDER_ID);
      var move_file = DriveApp.getFileById(file_to_read);
      move_file.moveTo(processed_folder);
    }  
  }
}

function start_processing_pending_folder(){
  process_all_folders_to_get_files(GOOGLE_MEET_ATTENDANCE_FOLDER_ID)
}



/*
 * Function: importGoogleSheetByFileId
 * 
 * @ params {string} file_id - string of Google_sheet Filed_ID of format: 
 *                              https://docs.google.com/spreadsheets/d/{{file_id}
 * 
 * @ description - Takes in a file_id and opens that google_sheet. reads and copies the values into the output spreadsheet. 
 * 
 */
function importGoogleSheetByFileId(file_id) {

  var output_ss = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(SHEET_TAB_NAME);

  var input_ss = SpreadsheetApp.openById(file_id);
  var file = DriveApp.getFileById(file_id);

  var sheet = input_ss.getSheets()[0];
  var data = sheet.getSheetValues(1,1,sheet.getLastRow() ,sheet.getLastColumn());

  
// first row is headers so we skip that. Loop through the rest of the rows and append to output google sheet.
  for (var i = 1; i < data.length; i++) {

    var attendee_name =  "\'"+ data[i][0] + " " + data[i][1];
    var attendee_email = data[i][2];
    var attendee_duration= data[i][3];
    var attendee_time_joined = data[i][4];
    var attendee_time_exited = data[i][5];

    var after_space = file.getName().substring(file.getName().indexOf(' ') + 1);
    console.log(after_space)
    var meeting_name = after_space.substring(after_space.indexOf(' ') + 1).replace("Attendance Report", "");
    console.log(meeting_name)
  
    var meeting_date = file.getName().substring(0, 10);

    var joined = "=concatenate(text(\"" + meeting_date + "\",\"mm/dd/yyyy\")&\" \"&text(\"" + data[i][4] + "\",\"hh:mm:ss\"))";
    var exited = "=concatenate(text(\"" + meeting_date + "\",\"mm/dd/yyyy\")&\" \"&text(\"" + data[i][5] + "\",\"hh:mm:ss\"))";


    // explode the time column value so we can tell if it's a string of hours, mins, or secs (or combination)
    var time = attendee_duration.split(' '); // split string on comma space
    if (time[1] == "hr") {
      var time_duration_sec = parseInt(time[0]) * 60 * 60;
    }if (time[1] == "hr" && time[3] == "min"){
      var time_duration_sec = parseInt(time[0]) * 60 * 60 + parseInt(time[2]) * 60;
    }if (time[1] == "min"){ 
      var time_duration_sec = parseInt(attendee_duration.substring(0, attendee_duration.indexOf(' '))) * 60;
    }if (time[1] == "sec"){
       var time_duration_sec = parseInt(attendee_duration);
    }
    var time_duration_mins = time_duration_sec/60
    // append the data to the sheet
    output_ss.appendRow([file.getName(), file_id, meeting_name, meeting_date, attendee_name, attendee_email, attendee_duration, attendee_time_joined, attendee_time_exited, joined, exited, time_duration_mins]);
  }
}