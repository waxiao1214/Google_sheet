// Add Custom Item In Menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation')
    .addItem('Es Vs Actual Daily Report', 'dailyReport')
    .addItem('Es Vs Actual Monthly Report', 'monthlyReport')
    .addItem('Daily Hours Logged Report', 'logReport')
    .addItem('Daily Hours New Month Creation', 'newMonthCreate')
    .addToUi();
}
function modifySheet(sheet) {
   //delete columns except 'Project','Name','Estimated time'
   var headings = sheet.getDataRange().offset(0, 0, 1).getValues()[0];
                    
   sheet.deleteColumns(1, headings.indexOf('Project'))
   sheet.deleteColumns(2, headings.indexOf('Name')-headings.indexOf('Project')-1);
   sheet.deleteColumns(3, headings.indexOf('Estimated Time')-headings.indexOf('Name')-1);
   sheet.deleteColumns(4, headings.indexOf('Tracked Time')-headings.indexOf('Estimated Time')-1);
   
   //Set the style 
   sheet.setColumnWidth(1, 300);
   sheet.setColumnWidth(2, 600);
   sheet.setColumnWidth(3, 100);
   
   //Rearrange Columns Order
    var columnSpec = sheet.getRange("A1:A");
    sheet.moveColumns(columnSpec, 4);
    var columnSpec = sheet.getRange("B1:B");
    sheet.moveColumns(columnSpec, 4); 
}
function deleteRows(sheet){
   
    var RANGE = sheet.getDataRange();
    var DELETE_VAL = ['Project Management','PM Activities','CH Internal-Angela Harper'];
    
     // The column to search for the DELETE_VAL (Zero is first)
    var COL_TO_SEARCH = 0;
    var rangeVals = RANGE.getValues();
    var newRangeVals = [];
    
   
    for(var i = 0; i < DELETE_VAL.length; i++){
      for(var n = rangeVals.length-1 ; n >=0  ; n--){
          if(rangeVals[n][COL_TO_SEARCH].toLowerCase() === DELETE_VAL[i].toLowerCase()){   
            sheet.deleteRow(n+1);
          };
      } 
    };
 
}
function duplicateProcess(sheet){
    // Sort the sheet without sorting header
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort(1);
    var column = 1;
    var lastRow = sheet.getLastRow();
    var columnRange = sheet.getRange(1,column,lastRow);
    var rangeArray = columnRange.getValues();
    //Convert to one dimensional array
    rangeArray = [].concat.apply([],rangeArray);
   
    //sort the data and find duplicates
   
    
    var duplicates = [];
    var indexes = [];
    for(var i=0;i<rangeArray.length-1;i++){
        if(rangeArray[i].toLowerCase()==rangeArray[i+1].toLowerCase()){
            duplicates.push(rangeArray[i]);
            indexes.push(i);
        }
    }
    
    var  count = {};
    duplicates.forEach(function(i) { count[i] = (count[i]||0) + 1;});
        
    //Highlight all the duplicates
    for(var i=0;i< indexes.length;i++){
      sheet.getRange(indexes[i]+1, column).setBackground("yellow");
    }
    // Get the duplicated str and its number of repeating
    var dup = [];
    for (var key in count){
       if(count.hasOwnProperty(key)){
           dup.push(key)
       }
    }
    //Add the hypen and number to duplicated cell
    for(var i=0;i<dup.length;i++){
      var firstIndex = rangeArray.indexOf(dup[i]);
      for(var j=0;j<=count[dup[i]];j++){
          var val = sheet.getRange(firstIndex+1+j, column).getValue()+"-"+j;
          sheet.getRange(firstIndex+1+j, column).setValue(val);
      }
    }  
}


function copyFromOtherSpreadSheet(){
  var ss = SpreadsheetApp.openByUrl(
    'https://docs.google.com/spreadsheets/d/1n7vXCx7LN6xkBLrru1WA-FKx2DCoG49-uR8J8gdQgaI/edit');
  //Get the current month and current month task sheet
  //var cur_month = Utilities.formatDate(new Date(), 'PST', 'MMMM');
  var cur_month = "August";
  var sheet_name = cur_month.concat(" Tasks");
  var task_sheet = ss.getSheetByName(sheet_name);
  var range = task_sheet.getRange(1, 1, task_sheet.getLastRow()-1,task_sheet.getLastColumn()-1).getValues();
  //Copy the task_sheet to the active spreadsheet
  var active_ss = SpreadsheetApp.getActiveSpreadsheet()
  active_ss.insertSheet('Sheet1');
  var sheet1 = active_ss.getSheetByName('Sheet1');
  sheet1.getRange(1, 1, task_sheet.getLastRow()-1,task_sheet.getLastColumn()-1).setValues(range);
  //task_sheet.deleteColumn(4);
  sheet1.getRange(2, 1, task_sheet.getLastRow() - 1, task_sheet.getLastColumn()).sort(1);
}
function processSheet1(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  var range = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();

 
  //delet rows whose time columns are 0 and empty
  var range1=sheet.getRange("A1:A").getValues();
  var filtered_r = range1.filter(String).length;
  sheet.deleteRows(filtered_r+1, sheet.getLastRow()-filtered_r-1);
  var delrow = [];
  for(var i=0;i<sheet.getLastRow();i++){
    if(range[i][2]==""&&(range[i][3]==0||range[i][3]=="")){
      delrow.push(i+1);
    }
  }
  for(var i=delrow.length-1;i>=0;i--){
      sheet.deleteRow(delrow[i]);
  }
  
  
}
function compareTwoSheets(){
  //Generate new sheet for result
  var ss =  SpreadsheetApp.getActiveSpreadsheet();
  
  
  //Reason 1: Finding new tasks
  var worksheet = ss.getSheetByName("Worksheet");
  var columnrange = worksheet.getRange(2, 1, worksheet.getLastRow());
  //Convert this into one dimensional array
  taskarray_works = columnrange.getValues();
  var taskarray_works = [].concat.apply([],taskarray_works);
 
  var comparesheet = ss.getSheetByName("Sheet1");
  
  var task_on_cmpsheet = comparesheet.getRange(2, 1, comparesheet.getLastRow()).getValues();
  task_on_cmpsheet = [].concat.apply([],task_on_cmpsheet);
  var cmplt_tasks = getDifferences(task_on_cmpsheet,taskarray_works);
  //Find the new tasks which are on Worksheet but not on compare sheet
  var new_tasks = getDifferences(taskarray_works,task_on_cmpsheet);
 
 
  
    //Reason 2: Finding completed tasks
  //Find the completed task from compare sheet which are on compare sheet but not on Worksheet
  
  //Copy new task from Worksheet to compare sheet
  copyRows(new_tasks,worksheet,comparesheet);
  
 
  //Copy complete tasks from compare to worksheet
 
  copyRows(cmplt_tasks,comparesheet,worksheet);
  comparesheet.getRange(2, 1, comparesheet.getLastRow() - 1, comparesheet.getLastColumn()).sort(1);
  worksheet.getRange(2, 1, worksheet.getLastRow() - 1, worksheet.getLastColumn()).sort(1);

  //reason 3
  var col_range = worksheet.getRange("B1:B");
  col_range.copyTo(comparesheet.getRange("B1:B"));
  
  //reason 4
  var estimate = worksheet.getRange("C1:C");
  estimate.copyTo(comparesheet.getRange("C1:C"));
  
  var res_sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("result");
  //Copy the first row from compare sheet to result sheet
  var first_row = ss.getSheetByName('Sheet1').getRange(1,1,1,res_sheet.getMaxRows()).getValues();
  res_sheet.getRange(1,1,1,res_sheet.getMaxRows()).setValues(first_row);
 
  //Insert the formula into the A2 of the res_sheet
  res_sheet.getRange("A2").setFormula("=Worksheet!A2=Sheet1!A2")
  var filldown_range = res_sheet.getRange(2, 1,500, 3);
  res_sheet.getRange("A2").copyTo(filldown_range);
}
function getDifferences(a1,a2){
   var result = [];
   for (var i = 0; i < a1.length; i++) {
      if (a2.indexOf(a1[i]) === -1) {
        result.push(a1[i]);
      }
    }
    return result;
}
function copyRows(tasks,source,dest){
 
  //Get the row number of specific task in sheet1
    var range = source.getRange(2,1,source.getLastRow()).getValues();
    range = [].concat.apply([],range); //making it into an one dimensional array
    var rows = [];
    
    for(var i=0;i<tasks.length;i++){
      var task = tasks[i];
      var pos = range.indexOf(task)+2;
      var row = source.getRange(pos,1,1,source.getLastColumn()).getValues();
      rows.push(row);
    }
    dest.getRange(dest.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
}
function removeDuplicates(sheet) {

  var data = sheet.getDataRange().getValues();
  var newData = [];
  for (var i in data) {
    var row = data[i];
    var duplicate = false;
    for (var j in newData) {
      if (row.join() == newData[j].join()) {
        duplicate = true;
      }
    }
    if (!duplicate) {
      newData.push(row);
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}
function deleteAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for(var i=0;i<sheets.length;i++){
     ss.deleteSheet(sheets[i]);
  }
 
}

function deleteTasks(){
  var ss = SpreadsheetApp.openByUrl(
    'https://docs.google.com/spreadsheets/d/1n7vXCx7LN6xkBLrru1WA-FKx2DCoG49-uR8J8gdQgaI/edit');
  //Get the current month and current month task sheet
  //var cur_month = Utilities.formatDate(new Date(), 'PST', 'MMMM');
  var cur_month = "August";
  var sheet_name = cur_month.concat(" Tasks");
  var task_sheet = ss.getSheetByName(sheet_name);
  var range = task_sheet.getRange(2,1,task_sheet.getLastRow()-1,3)
  range.clear();
  var worksheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Worksheet');
  var range_work = worksheet.getRange(2,1,task_sheet.getLastRow()-1,3).getValues();
  task_sheet.getRange(2,1,task_sheet.getLastRow()-1,3).setValues(range_work);
  
  //Step 5
  var hour_sheet = ss.getSheetByName("Est vs Actual Hours August");
  var hour_range = hour_sheet.getRange(2,1,hour_sheet.getLastRow()-1,3)
  hour_range.clear();
 
  var hour_work = task_sheet.getRange(2,1,task_sheet.getLastRow()-1,3);
  hour_work.copyTo((hour_range),{contentsOnly:true});
  //hour_sheet.getRange(2,1,hour_sheet.getLastRow()-1,3).setValues(hour_work);
}
function dailyReport() {
  var sheet = SpreadsheetApp.getActiveSheet();
  modifySheet(sheet);
  deleteRows(sheet)
  duplicateProcess(sheet);
  copyFromOtherSpreadSheet();
  processSheet1()
  compareTwoSheets();
  deleteTasks()
}