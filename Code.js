// @ts-nocheck
var scriptProperties = PropertiesService.getScriptProperties();

function Otterize() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var location = 'trsta'
    var menuEntries = [];

    menuEntries.push({name: "Shelf List From User", functionName: "shelflistFromUser"});
    menuEntries.push({name: "Shelf List From Inventory", functionName: "shelflistFromInventory"});
    menuEntries.push({name: "Get info for ItemIds", functionName: "getInfo"});
    menuEntries.push({name: "Change Location Code", functionName: "changeCode"});
    menuEntries.push({name: "Produce Reshelve Sheet", functionName: "runReshelve"});
    menuEntries.push({name: "Should Be There But Aren't", functionName: "shouldBeThere"});
    menuEntries.push({name: "There But Should Not Be", functionName: "shouldNotBeThere"});
    spreadsheet.addMenu("Inventory", menuEntries);
    url = "https://librarycatalog2.ccc.edu/iii/sierra-api/v5/token";

    var options = {
        "method" : "POST",
        "headers" : {
        "Authorization" : "Basic " + "\"" + cred + "\"",
        }
    };

    var response = UrlFetchApp.fetch(url,options);
    var json_data = JSON.parse(response.getContentText());
    var accesstoken = json_data.access_token;
    //spreadsheet.getRange('J2').setValue(accesstoken);

    scriptProperties.setProperty('accesstoken', accesstoken);
    scriptProperties.setProperty('spreadsheet', spreadsheet);
    scriptProperties.setProperty('location', location);

}//end Otterize

function onOddit(e) {
  var sheet = SpreadsheetApp.getActiveSheet();
  // cell must have a value, be only one row, and be from the first sheet
  if( e.range.getValue() && (e.range.getNumRows() == 1) && sheet.getIndex() == 1 && e.range.getRow != 1) {


  var accesstoken = scriptProperties.getProperty('accesstoken');
  var value = e.range.getValue();
  //populate column B w 'buhbhingo'
  //e.range.offset(0,1).setValue(accesstoken);

  var url = 'https://librarycatalog2.ccc.edu:443/iii/sierra-api/v6/items/query?offset=0&limit=1';

  var options = {
   "method" : "POST",
   "headers" : {
       "Authorization" : "Bearer " + accesstoken
     },
   "contentType" : "raw",
   "payload" : '{"target":{"record":{"type":"item"},"field":{"tag":"b"}},"expr":{"op":"equals","operands":["' + value + '",""]}}'
  };

  var result = UrlFetchApp.fetch(url, options);
  var json_data = JSON.parse(result.getContentText());
//    e.range.offset(0,1).setValue(json_data)
  //make sure we have data back ...
  if(json_data) {
    var entries = JSON.stringify(json_data.entries);

    var itemID = entries.split('/')[7].split("\"")[0];
    e.range.offset(0,9).setValue(itemID);

    var url = 'https://librarycatalog2.ccc.edu/iii/sierra-api/v5/items/' + itemID;

    var options = {
   "method" : "GET",
   "headers" : {
       "Authorization" : "Bearer " + accesstoken
     }
    };
    var result = UrlFetchApp.fetch(url, options);
    var json_data = JSON.parse(result.getContentText());
    var in_cn = json_data.callNumber;
    var loc = new locCallClass;
    var out_cn = loc.returnNormLcCall(in_cn);
    //e.range.offset(0,1).setValue(result);
    //e.range.offset(0,1)setValue(JSON.parse(result.getContextText()));
    if(json_data) {
      e.range.offset(0,3).setValue(in_cn);
      e.range.offset(0,4).setValue(out_cn)
      e.range.offset(0,5).setValue(json_data.status.display);
      e.range.offset(0,6).setValue(json_data.status.duedate);
      e.range.offset(0,7).setValue(json_data.location.code);
      e.range.offset(0,8).setValue('=\"' + Utilities.formatDate(new Date(), "GMT-4:00", "yyyy-MM-dd' 'HH:mm:ss") + '\"')
      //e.range.offset(0,6).setValue(json_data.bibIds[0])

      var bibId = json_data.bibIds[0];
      var url = 'https://librarycatalog2.ccc.edu/iii/sierra-api/v5/bibs/' + bibId;
      var result = UrlFetchApp.fetch(url, options);
      var json_data = JSON.parse(result.getContentText());
      if(json_data) {
        e.range.offset(0,1).setValue(json_data.title);
        e.range.offset(0,2).setValue(json_data.author);
      }//end3rdif
    }//end2ndif
  }//end1stif
  else{
    e.range.offset(0,1).setValue('buhbingo');
    }
}//end onOddit
}//endrow if

function writeItemIds() { // puts itemID's in a range into column H
  var accesstoken = scriptProperties.getProperty('accesstoken');
  var location = scriptProperties.getProperty('location');
  var starting  = scriptProperties.getProperty('starting');
  var ending = scriptProperties.getProperty('ending');
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if(!(spreadsheet.getActiveSheet().getName()==='shelflist')) {
    SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName('shelflist'))
  }


  var url = 'https://librarycatalog2.ccc.edu:443/iii/sierra-api/v6/items/query?offset=0&limit=3000';
  var options = {
   "method" : "POST",
   "headers" : {
       "Authorization" : "Bearer " + accesstoken
     },
   "contentType" : "raw",
   "payload" : '{"queries":[{"target":{"record":{"type":"item"},"id":79},"expr":{"op":"equals","operands":["'+location+'",""]}},"and",{"target":{"record":{"type":"bib"},"field":{"tag":"c"}},"expr":{"op":"between","operands":["'+starting+'","'+ending+'"]}}]}'
  };
    
  let row = 2;
  var result = UrlFetchApp.fetch(url, options);
  var json_data = JSON.parse(result.getContentText());
  for (let i = json_data.entries.length -1; i >= 0; i--){
    let itemID = json_data.entries[i].link.split('/')[7].split("\"")[0];
    spreadsheet.getRange('J' + row).setValue(itemID);
    row++
  }//endforloop
}//end writeItemIds

function getInfo() {
  var accesstoken = scriptProperties.getProperty('accesstoken');

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if(!(spreadsheet.getActiveSheet().getName()==='shelflist')) {
    SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName('shelflist'))
  }

  var lr=spreadsheet.getLastRow()

  for (i=2; i<=lr; i++){
    if (( spreadsheet.getRange('A'+i).isBlank() ) && (i != 2)){
      var firstBlank = i-1;
      console.log('firstblank: ' + firstBlank)
      break;
    }//end if
  }// end for loop

  for (i=firstBlank; i<lr+1;i++) {
    if (!spreadsheet.getRange('J' + i).isBlank()){

      var itemID = spreadsheet.getRange('J'+i).getValue();
      var url = 'https://librarycatalog2.ccc.edu/iii/sierra-api/v5/items/' + itemID;

      var options = {
        "method" : "GET",
        "headers" : {
        "Authorization" : "Bearer " + accesstoken
      }
      };

      var result = UrlFetchApp.fetch(url, options);
      var anotherjson_data = JSON.parse(result.getContentText());
      console.log('anotherjson data: ' + anotherjson_data)
      var in_cn = anotherjson_data.callNumber;
      console.log('in_cn: ' + in_cn)
      var loc = new locCallClass;
      var out_cn = loc.returnNormLcCall(in_cn);
      if(anotherjson_data) {
        spreadsheet.getRange('A' + i).setValue(anotherjson_data.barcode);
        spreadsheet.getRange('D' + i).setValue(in_cn);
        spreadsheet.getRange('E' + i).setValue(out_cn);
        spreadsheet.getRange('F' + i).setValue(anotherjson_data.status.display);
        spreadsheet.getRange('G' + i).setValue(anotherjson_data.status.duedate);
        spreadsheet.getRange('H' + i).setValue(anotherjson_data.location.code);
        spreadsheet.getRange('I' + i).setValue('=\"' + Utilities.formatDate(new Date(), "GMT-6:00", "yyyy-MM-dd' 'HH:mm:ss") + '\"');

        var bibId = anotherjson_data.bibIds[0];
        var url = 'https://librarycatalog2.ccc.edu/iii/sierra-api/v5/bibs/' + bibId;
        var result = UrlFetchApp.fetch(url, options);
        var yetanotherjson_data = JSON.parse(result.getContentText());
        spreadsheet.getRange('B' + i).setValue(yetanotherjson_data.title);
        spreadsheet.getRange('C' + i).setValue(yetanotherjson_data.author);
      }//end second if
    }//end first if
  }//endfor
    var range = spreadsheet.getDataRange();
    spreadsheet.sort(5);
}//end getInfo

function changeCode() {

  var location = Browser.inputBox("Input location code:");
  scriptProperties.setProperty('location', location);

}

function shelflistFromInventory() {

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheets()[0];
  if(!(spreadsheet.getActiveSheet().getName()==='shelflist')) {
    SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName('shelflist'))
  }

  var lr=sheet.getLastRow()
  console.log(lr);

  var starting = sheet.getRange(1, 5).getValue();
  var ending = sheet.getRange(lr, 5).getValue();

  scriptProperties.setProperty('starting', starting);
  scriptProperties.setProperty('ending', ending);
  writeItemIds();

}

function shelflistFromUser() {
  var starting = Browser.inputBox("Starting Call Number");
  var ending = Browser.inputBox("Ending Call Number");

  scriptProperties.setProperty('starting', starting);
  scriptProperties.setProperty('ending', ending);
  writeItemIds();
}

function joinShelfListToInventory() {
  /*
  This function, joinShelfListToInventory, will do what equates to a LEFT OUTER JOIN on two sheets -
  the shelflist (what is in the system, and what we expect to be on the shelf) and
  the inventory (what physical items we scanned and found to be on the shelf).
  This will tell us several things
  1) What items are missing that we expect should be there
  2) What order the items should go in, and how they should be placed on the shelf.
  */

  // check for inventory sheet and the shelflist sheet
  var inventory_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("inventory"),
      shelflist_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("shelflist");
      //reshelve_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("reshelve");

  // if both those exist ...
  if( inventory_sheet && shelflist_sheet ) {
    // create reshelve sheet
    var reshelve_sheet = 1;
    // create an array for the reshelving sheet, shelflist, and inventory
    var reshelve = [],
        shelflist = [],
        inventory = [],
        shelflist_barcodes_range = shelflist_sheet.getRange('A:A').getValues(), //barcodes from shelflist
        inventory_barcodes_range = inventory_sheet.getRange('A:A').getValues(); //barcodes from inventory

    //Logger.log(shelflist_barcodes_range.length + ' ' + inventory_barcodes_range.length);

    //fill the arrays ...
    for (var i=0; i<shelflist_barcodes_range.length; i++) {
      shelflist.push(shelflist_barcodes_range[i][0]);
    }
    for (var i=0; i<inventory_barcodes_range.length; i++) {
      inventory.push(inventory_barcodes_range[i][0]);
    }

    //loop through the shelflist, and find the index of the bar code from inventory
    for (var i=0; i<shelflist.length; i++) {
      var match = inventory.indexOf(shelflist[i]);
        reshelve.push(match);
    } //end for

    //remove the reshelve sheet (if it's there), and place a new one into the spreadsheet
    var reshelve_sheet = SpreadsheetApp.getActive().getSheetByName('reshelve');
    if(reshelve_sheet){
      SpreadsheetApp.getActive().deleteSheet(reshelve_sheet);
    }

    //create the spreadsheet 'reshelve', and put it at the end of the other sheets.
    // old sheet ...
    //var reshelve_sheet = SpreadsheetApp.getActive().insertSheet('reshelve', SpreadsheetApp.getActive().getSheets().length);

    // insert a new sheet, add it to the last index, and make sure the new sheet has the proper amount of rows ... using shelflist_sheet as our template
    var reshelve_sheet = SpreadsheetApp.getActive().insertSheet('reshelve', SpreadsheetApp.getActive().getSheets().length, {template: shelflist_sheet});

    /*reshelve_sheet.insertRows( Math.floor(reshelve_sheet.getMaxRows()),
                              shelflist.length - Math.floor(reshelve_sheet.getMaxRows()) );
    */

    //start filling the reshelve sheet
    var shelflist_range = shelflist_sheet.getRange( 'A1:D' + Math.floor(shelflist_sheet.getMaxRows()) ),
        reshelve_range = reshelve_sheet.getRange( 'A1:E' + Math.floor(reshelve_sheet.getMaxRows()) ),
        shelflist_range_values = shelflist_range.getValues();

    if (shelflist_range_values.length != reshelve_range.getValues().length) {
      Logger.log('length of reshelve sheet does not match the length of the shelflist');
      return(0); // we can't / shouldn't go on if these don't match up
    }

    //we need to add the index value to the shelflist_range_values array ... do it here
    for (var i=0; i<shelflist_range_values.length; i++) {
      var position = i+1;

      //item not found in shelflist, mark the row accordingly
      if (reshelve[i] == -1){
        //hopefully this shouldn't happen often, as the getRange is an expensive operation
        reshelve_sheet.getRange('A' + position + ':E' + position).setBackground('LightCoral');
        shelflist_range_values[i][4] = null;
      }

      else{
        shelflist_range_values[i][4] = reshelve[i] + 1;
      }

    } //end for

    //finally set values in the reshelve sheet
    reshelve_range.setValues(shelflist_range_values);
    reshelve_sheet.autoResizeColumn(2);

  } //end if
} //end function joinShelfListToInventory()

function runReshelve() {
  SpreadsheetApp.getUi()
     .alert('Running Script to Produce "reshelve" sheet. \n\nClick OK to Continue');
  joinShelfListToInventory();
} //end function runReshelve


function shouldBeThere() {
//paste in inventory
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheets()[0];
  // check whether there's a sheet shouldBeThere
  // make it active
  if ( spreadsheet.getSheetByName('shouldBeThere') == null){
    var shouldBeThere = SpreadsheetApp.getActive().insertSheet('shouldBeThere', SpreadsheetApp.getActive().getSheets().length);
  } else {
    if(!(spreadsheet.getActiveSheet().getName()==='shouldBeThere')) {
      SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName('shouldBeThere'));
      }
  }

  copySheet = spreadsheet.getSheetByName('inventory');
  pasteSheet = spreadsheet.getSheetByName('shouldBeThere');

  pasteSheet.setFrozenRows(0);
  pasteSheet.getRange(1,1,sheet.getMaxRows(), sheet.getMaxColumns()).clearContent();

  var source = copySheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  var destination = pasteSheet.getRange(pasteSheet.getLastRow()+1,1);
  source.copyTo(destination);
  pasteSheet.setFrozenRows(1);
//paste in shelf list

  copySheet = spreadsheet.getSheetByName('shelflist');
  pasteSheet = spreadsheet.getSheetByName('shouldBeThere');


  var source = copySheet.getRange(2, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  var destination = pasteSheet.getRange(pasteSheet.getLastRow()+1,1);
  source.copyTo(destination);
//sort by A
  pasteSheet.sort(1);
//Dedupe
  // Go down the pastsheet spreadsheet one by one, comparing barcodes
  var lr=pasteSheet.getLastRow();
  for (i=2; i<=lr; i++){
    var minusone = i - 1;
    if ( spreadsheet.getRange('A'+i).getValue() == spreadsheet.getRange('A'+ minusone).getValue()){
  // if two barcodes match, delete them both
      //pasteSheet.getRange(i, 1 ,1, pasteSheet.getMaxColumns()).setBackground('LightCoral');
      //pasteSheet.getRange(minusone ,1, 1, pasteSheet.getMaxColumns()).setBackground('LightCoral');
      pasteSheet.getRange(i, 1 ,1, pasteSheet.getMaxColumns()).clearContent();
      pasteSheet.getRange(minusone ,1, 1, pasteSheet.getMaxColumns()).clearContent();
      //console.log(i);
    }//endif
  }//endforloop
  pasteSheet.sort(1)
//delete anything w/ a status other than available
  var lr=pasteSheet.getLastRow();
  for (i=2; i<=lr; i++){
    if (!(pasteSheet.getRange('F'+i).getValue() == 'Available') || !(pasteSheet.getRange('G'+i).isBlank())){
      pasteSheet.getRange(i, 1 ,1, pasteSheet.getMaxColumns()).clearContent();
    }//endif
  }//endfor
  pasteSheet.sort(1)
//delete anything w a due date in the future
//mark missing or at least get ready to export into a format that makes that easy for a bulk update in Sierra
}//end function shouldBeThere

//checks for inventory items with status other than available, a due date,
//or the wrong location code. Copies these items from the inventory sheet
//to a new sheet named shouldNotBeThere
function shouldNotBeThere() {
  var location = scriptProperties.getProperty('location');

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheets()[0];
  // check whether there's a sheet shouldNotBeThere
  // make it active
  if ( spreadsheet.getSheetByName('shouldNotBeThere') == null){
    var shouldBeThere = SpreadsheetApp.getActive().insertSheet('shouldNotBeThere', SpreadsheetApp.getActive().getSheets().length);
  } else {
    if(!(spreadsheet.getActiveSheet().getName()==='shouldNotBeThere')) {
      SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName('shouldNotBeThere'));
      }
  }

  copySheet = spreadsheet.getSheetByName('inventory');
  pasteSheet = spreadsheet.getSheetByName('shouldNotBeThere');

  pasteSheet.setFrozenRows(0);
  pasteSheet.getRange(1,1,sheet.getMaxRows(), sheet.getMaxColumns()).clearContent();
  var source = copySheet.getRange(1, 1, 1, pasteSheet.getMaxColumns());
  var destination = pasteSheet.getRange(pasteSheet.getLastRow()+1,1);
  source.copyTo(destination);
  pasteSheet.setFrozenRows(1);

  var lr=copySheet.getLastRow();
  for (i=2; i<=lr; i++){
    if (!(copySheet.getRange('F'+i).getValue() == 'Available') || !(copySheet.getRange('G'+i).isBlank()) || !(copySheet.getRange('H'+i).getValue() == location)){
      var source = copySheet.getRange(i, 1 ,1, pasteSheet.getMaxColumns());
      var destination = pasteSheet.getRange(pasteSheet.getLastRow()+1,1);
      source.copyTo(destination);
      console.log(location)
    }//endif
  }//endfor

  }//end shouldNotBeThere
