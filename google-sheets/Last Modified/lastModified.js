/**
 * Purpose:
 *    Monitors specific columns for changes and updates a cell with the last modified date when a
 *    change occurs within those cells.
 * 
 * KNOWN LIMITATIONS: 
 *    When dragging over multiple monitored cells. The function may time out.  Tested to be aroudn 85 rows
 */

 const MONITORED_RANGE = [2,3,4, 6]; //  type: array[int] -The range of columns that are to be monitored for changes.
 const DATE_MODIFIED_COLUMN = 11; //  type[int] - The column in which the date last modified will be stored.
 const OLD_VALUE_COLUMN = 12;
 const TARGET_SHEET_TAB_NAME = ""; // type:string - Sheet to be monitored
 
 const CELL_BACKGROUND_COLOUR = "#ea9999";
 const CELL_BORDER_COLOUR = "#cc0000";
 
 
 /**
  * A simple trigger event that runs anytime a change on the sheet occurs.
  *
  * @param {event} event     The event object from Google Sheets.
  * @description   Checks to see if a the event was triggered within the correct column and rows
  *                if so it will call the `trackChanges function` otherwise exits.
  * @customfunction
  */
 function onEdit(event) {
 
   var range = event.range;
 
   var firstRow = range.getRow();
   var lastRow = range.getLastRow();
 
   var firstColumn = range.getColumn();
   var lastColumn = range.getLastColumn();
 
   Logger.log("Column: "+ firstColumn)
   if (range.getSheet().getName() != TARGET_SHEET_TAB_NAME) return;
   if (range.getRow() == 1) return;
 
   Logger.log("RANGE MODIFIED: ("+ firstRow+","+firstColumn+ " : "+ lastRow + " : "+ lastColumn +")")
   for( var activeColumn = firstColumn; activeColumn <= lastColumn; activeColumn=activeColumn+1){
     for (var i = 0; i < MONITORED_RANGE.length; i = i + 1) {
       
       if (MONITORED_RANGE[i] == activeColumn) {
          trackChanges(event);
       }
     }
   }
 }
 
 
 /**
  * Function to track cell changes and update the last modified column.
  * 
  * @param {event} event  The triggered event object
  */
 function trackChanges(event) {
 
   var range = event.range;
   var sheet = range.getSheet();
 
   var firstRow = range.getRow();
   var lastRow = range.getLastRow();
 
   var firstColumn = range.getColumn();
   var lastColumn = range.getLastColumn();
   var oldValue = event.oldValue;
   
   if (oldValue == undefined) {
     oldValue = "Unable to store previous values";
   }
 
   for( var activeColumn = firstColumn; activeColumn <= lastColumn; activeColumn=activeColumn+1){
     for (var activeRow = firstRow; activeRow <= lastRow; activeRow = activeRow + 1) {
       for (var i = 0; i <= MONITORED_RANGE.length; i = i + 1) {
 
         if (MONITORED_RANGE[i] == activeColumn) {
      
           sheet.getRange(activeRow, activeColumn).setBackground(CELL_BACKGROUND_COLOUR);
           sheet.getRange(activeRow, activeColumn).setBorder(
             true, true, true, true, false, false, CELL_BORDER_COLOUR, SpreadsheetApp.BorderStyle.DASHED);
           var temp = OLD_VALUE_COLUMN + i -1
           sheet.getRange(activeRow,temp).setValue(oldValue);
         }
       }
     }
   }
 
   for (var modifiedRow = firstRow; modifiedRow <= lastRow; modifiedRow = modifiedRow + 1) {
     updateLastModifiedCell(sheet,modifiedRow,firstRow,lastRow,firstColumn,lastColumn);
     sheet.getRange(modifiedRow,OLD_VALUE_COLUMN).setValue(oldValue);
   }
 }
 
 
 function updateLastModifiedCell(sheet,modifiedRow,firstRow,lastRow,firstColumn,lastColumn){
  // Set Tracking Note
  sheet.getRange(modifiedRow, DATE_MODIFIED_COLUMN).setNote(
     "Cells Modified: \n \t" + sheet.getRange(firstRow, firstColumn).getA1Notation() +
     ":" + sheet.getRange(lastRow, lastColumn).getA1Notation() +
     "\nCell Modified on: \n \t" + new Date() +
     "\nLast Modified by: \n \t" + Session.getActiveUser() +
     "\nPrevious Date: \n \t" + sheet.getRange(modifiedRow, DATE_MODIFIED_COLUMN).getValue()
   );
 
   // Set Value
   sheet.getRange(modifiedRow, DATE_MODIFIED_COLUMN).setValue(new Date());
 
   // Set Formatting
   sheet.getRange(modifiedRow, DATE_MODIFIED_COLUMN).setBackground(CELL_BACKGROUND_COLOUR);
   sheet.getRange(modifiedRow, DATE_MODIFIED_COLUMN).setBorder(
           true, true, true, true, false, false, CELL_BORDER_COLOUR, SpreadsheetApp.BorderStyle.DASHED);
 
 }
 
 // Code for adding comments to a modified cell.
 /**
  * var previousValue = e.oldValue;
 
   if (previousValue == undefined) {
     previousValue = "Unable to store previous values";
   }
 
   
   // Sets modifed cell note
   range.setNote('Cell Modified on: \n \t' +
                  new Date() +
                 '\nLast Modified by: \n \t' + 
                 Session.getActiveUser() + 
                 '\nPrevious Value: \n \t' 
                 + previousValue);
  */