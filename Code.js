// @ts-nocheck
var scriptProperties = PropertiesService.getScriptProperties();

function Otterize() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var location = 'trsta'
    var menuEntries = [];

    menuEntries.push({name: "Shelf List From User", functionName: "shelflistFromUser"});
    menuEntries.push({name: "Shelf List From Inventory", functionName: "shelflistFromInventory"});
    menuEntries.push({name: "Get info for ItemIds", functionName: "getInfo"});
    menuEntries.push({name: "Try Barcodes again", functionName: "tryAgain"});
    menuEntries.push({name: "Change Location Code", functionName: "changeCode"});
    menuEntries.push({name: "Hi-Lite Misshelvings", functionName: "hiliteMisshelvings"});
    menuEntries.push({name: "ID Missing Items", functionName: "shouldBeThere"});
    menuEntries.push({name: "ID Items w wrong status or location", functionName: "shouldNotBeThere"});
    menuEntries.push({name: "Write Stats", functionName: "writeStats"});
    menuEntries.push({name: "Copy Sheet", functionName: "copySheet"});

  
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

        //e.range.offset(0,11).setValue(json_data.title.length);
        if (json_data.title.length > 0) {
          e.range.offset(0,1).setValue(json_data.title);
        }
        else {
          e.range.offset(0,1).setValue('none');
        }//endtitleelse
        //e.range.offset(0,11).setValue(json_data.author.length);
        if (json_data.author.length > 0) {
          e.range.offset(0,2).setValue(json_data.author);
        }
        else {
          e.range.offset(0,2).setValue('none');
        }//endauthorelse
      }//end3rdif
    }//end2ndif
  }//end1stif
 // else{
//    e.range.offset(0,1).setValue('buhbingo');
    //}
}//end onOddit
}//endrow if


function writeItemIds() { // puts itemID's in a range into column J
  var accesstoken = scriptProperties.getProperty('accesstoken');
  var location = scriptProperties.getProperty('location');
  var starting  = scriptProperties.getProperty('starting');
  var ending = scriptProperties.getProperty('ending');
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if(!(spreadsheet.getActiveSheet().getName()==='shelflist')) {
    SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName('shelflist'))
  }

  sheet = spreadsheet.getSheetByName('shelflist');

  sheet.getRange(2,1,sheet.getLastRow(), sheet.getLastColumn()).clearContent();

  //just debugging here...
  spreadsheet.getRange('K2').setValue(starting);
  spreadsheet.getRange('K3').setValue(ending);
  spreadsheet.getRange('L2').setValue(location);


   // spreadsheet.getRange('M2').setValue('buhbingo');

    var url = 'https://librarycatalog2.ccc.edu:443/iii/sierra-api/v6/items/query?offset=0&limit=3000';
    //doing it first with bib callnos
    var options = {
    "method" : "POST",
    "headers" : {
        "Authorization" : "Bearer " + accesstoken
      },
    "contentType" : "raw",
      "payload" : '{"queries":[[[{"target":{"record":{"type":"item"},"field":{"tag":"c"}},"expr":{"op":"greater_than_or_equal","operands":["'+starting+'",""]}},"and",{"target":{"record":{"type":"item"},"field":{"tag":"c"}},"expr":{"op":"less_than_or_equal","operands":["'+ending+'",""]}}],"or",[{"target":{"record":{"type":"bib"},"field":{"tag":"c"}},"expr":{"op":"greater_than_or_equal","operands":["'+starting+'",""]}},"and",{"target":{"record":{"type":"bib"},"field":{"tag":"c"}},"expr":{"op":"less_than_or_equal","operands":["'+ending+'",""]}}]],"and",{"target":{"record":{"type":"item"},"id":79},"expr":{"op":"equals","operands":["'+location+'",""]}},"and",{"target":{"record":{"type":"item"},"id":65},"expr":{"op":"not_exists","operands":["      ",""]}},"and",{"target":{"record":{"type":"item"},"id":88},"expr":{"op":"equals","operands":["-",""]}}]}'
    };

    var row = 2;
    var result = UrlFetchApp.fetch(url, options);
    var json_data = JSON.parse(result.getContentText());
    for (let i = json_data.entries.length -1; i >= 0; i--){
      let itemID = json_data.entries[i].link.split('/')[7].split("\"")[0];
      spreadsheet.getRange('J' + row).setValue(itemID);
      row++
    }//endforloop
}//end writeItemIds

function tryAgain() {

  var accesstoken = scriptProperties.getProperty('accesstoken');
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheets()[0];
  if(!(spreadsheet.getActiveSheet().getName()==='inventory')) {
    SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName('inventory'))
  }

  var lr=sheet.getLastRow()
  for (i=2; i<=lr; i++){
    if (!(spreadsheet.getRange('A'+i).isBlank()) && (spreadsheet.getRange('B'+i).isBlank())){
      var value = spreadsheet.getRange('A'+i).getValue();
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
      //console.log(json_data);
      //console.log('data totals: ' + json_data.total);
      //console.log('as int: ' + parseInt(json_data.total));
      //
      //make sure we have data back ...
      if(parseInt(json_data.total) > 0 ) {
        var entries = JSON.stringify(json_data.entries);
        console.log(json_data);

        var itemID = entries.split('/')[7].split("\"")[0];
        spreadsheet.getRange('J' + i).setValue(itemID);

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
          spreadsheet.getRange('D' + i).setValue(in_cn);
          spreadsheet.getRange('E' + i).setValue(out_cn);
          spreadsheet.getRange('F' + i).setValue(json_data.status.display);
          spreadsheet.getRange('G' + i).setValue(json_data.status.duedate);
          spreadsheet.getRange('H' + i).setValue(json_data.location.code);
          spreadsheet.getRange('I' + i).setValue('=\"' + Utilities.formatDate(new Date(), "GMT-6:00", "yyyy-MM-dd' 'HH:mm:ss") + '\"');

          var bibId = json_data.bibIds[0];
          var url = 'https://librarycatalog2.ccc.edu/iii/sierra-api/v5/bibs/' + bibId;
          var result = UrlFetchApp.fetch(url, options);
          var yetanotherjson_data = JSON.parse(result.getContentText());
          if (yetanotherjson_data.title.length > 0) {
            spreadsheet.getRange('B' + i).setValue(yetanotherjson_data.title);
          }
          else{
            spreadsheet.getRange('B' + i).setValue('none');
          }
          if (yetanotherjson_data.author.length > 0) {
            spreadsheet.getRange('C' + i).setValue(yetanotherjson_data.author);
          }
          else{
            spreadsheet.getRange('C' + i).setValue('none');
          }
          }//end3rdif
        }     //end2ndif
      else {
        spreadsheet.getRange('B' + i).setValue('UNATTACHED BARCODE');
      }

      }//end1stif
  }// end for loop
}//end tryAgain

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

  var starting = sheet.getRange(2, 5).getValue();
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
  pasteSheet.sort(4)
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
      //console.log(location)
    }//endif
  }//endfor

  }//end shouldNotBeThere

function writeStats() {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var yourNewSheet = activeSpreadsheet.getSheetByName("Stats");

    if (yourNewSheet != null) {
        activeSpreadsheet.deleteSheet(yourNewSheet);
    }

    yourNewSheet = activeSpreadsheet.insertSheet();
    yourNewSheet.setName("Stats");

   activeSpreadsheet.getRange('A1').setValue('Location');
   activeSpreadsheet.getRange('A2').setValue('Books Scanned');
   activeSpreadsheet.getRange('A3').setValue('Duration');
   activeSpreadsheet.getRange('A4').setValue('Books/Minute');
   activeSpreadsheet.getRange('A5').setValue('Call Number Range');
   activeSpreadsheet.getRange('A6').setValue('Missing');
   activeSpreadsheet.getRange('A7').setValue('Cataloging Corrections');
   activeSpreadsheet.getRange('A8').setValue('Unattached Barcodes');
   activeSpreadsheet.getRange('A9').setValue('Found Items');

  var inventory_sheet = activeSpreadsheet.getSheetByName("inventory");
  var location = inventory_sheet.getRange('H2').getValue();
  activeSpreadsheet.getRange('B1').setValue(location);

  var booksscanned = inventory_sheet.getLastRow() - 1;
  activeSpreadsheet.getRange('B2').setValue(booksscanned);

  var time1 = inventory_sheet.getRange('I2').getValue().slice(11);

  var seconds1 = time1.slice(6);

  var minutes1 = time1.slice(3,5);

  var hours1 = time1.slice(0,2);

  var totalseconds1 = parseInt(seconds1) + (parseInt(minutes1)*60) + (parseInt(hours1)*60*60)
  console.log('total seconds 1: ' + totalseconds1);

  var time2 = inventory_sheet.getRange('I' + (booksscanned +1)).getValue().slice(11);
  //console.log(time2);

  var seconds2 = time2.slice(6);

  //console.log(seconds2);
  var minutes2 = time2.slice(3,5);

  //console.log(minutes2);
  var hours2 = time2.slice(0,2);
  //console.log(hours2);

  var totalseconds2 = parseInt(seconds2) + (parseInt(minutes2)*60) + (parseInt(hours2)*60*60)
  console.log('total seconds 2: ' + totalseconds2);

  var secondsdiff = totalseconds2 - totalseconds1;
  var minutesdiff  = Math.floor(secondsdiff/60);
  activeSpreadsheet.getRange('B3').setValue(minutesdiff + ' minutes');

  var booksamin = (booksscanned/minutesdiff);
  activeSpreadsheet.getRange('B4').setValue(booksamin);

  var starting = inventory_sheet.getRange('D2').getValue();
  var ending = inventory_sheet.getRange('D' + inventory_sheet.getLastRow()).getValue();
  activeSpreadsheet.getRange('B5').setValue(starting);
  activeSpreadsheet.getRange('C5').setValue(ending);


  var missing_sheet = activeSpreadsheet.getSheetByName("shouldBeThere");
  var missing = missing_sheet.getLastRow()-1
  activeSpreadsheet.getRange('B6').setValue(missing);

  var miscataloged_sheet = activeSpreadsheet.getSheetByName("miscataloged");
  var missing = miscataloged_sheet.getLastRow()-1
  activeSpreadsheet.getRange('B7').setValue(missing);

  var unattached_sheet = activeSpreadsheet.getSheetByName("unattached barcodes");
  var unattached = unattached_sheet.getLastRow()-1
  activeSpreadsheet.getRange('B8').setValue(unattached);

  var unattached_sheet = activeSpreadsheet.getSheetByName("found!");
  var unattached = unattached_sheet.getLastRow()-1
  activeSpreadsheet.getRange('B9').setValue(unattached);

}

function copySheet(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var date = new Date();
  var date = date.toString().slice(0,15);
  var copy=ss.copy(date + ' Scanning Session');
  var copyId=copy.getId()
  var sheetNumber=ss.getSheets().length;
  for(var i=0; i<sheetNumber;i++)  {
    var values=ss.getSheets()[i].getDataRange().getValues();
    SpreadsheetApp.openById(copyId).getSheets()[i].getDataRange().setValues(values);
  }
}

function hiliteMisshelvings(){


  //initializing stuff
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('inventory');
  // check whether there's a sheet shouldBeThere
  // make it active
  if ( spreadsheet.getSheetByName('shelflist') == null){
    SpreadsheetApp.getUi()
      .alert('Run a Shelflist Function first, please.');
    return;
  }

  var inventory_sheet = spreadsheet.getSheetByName("inventory"),
      shelflist_sheet = spreadsheet.getSheetByName("shelflist");

  inventory_sheet.getRange(1,1,inventory_sheet.getLastRow(),inventory_sheet.getLastColumn()).setBackground(null);
  inventory_sheet.getRange(2,11,inventory_sheet.getLastRow()).clear();



    // create an array for the reshelving sheet, shelflist, and inventory
  var shelflist = [],
      inventory = [],
      shelflist_barcodes_range = shelflist_sheet.getRange('A:A').getValues(), //barcodes from shelflist
      inventory_barcodes_range = inventory_sheet.getRange('A:A').getValues(); //barcodes from inventory

    //fill the arrays ...
    for (var i=0; i<shelflist_barcodes_range.length; i++) {
      shelflist.push(shelflist_barcodes_range[i][0]);
    }
    for (var i=0; i<inventory_barcodes_range.length; i++) {
      inventory.push(inventory_barcodes_range[i][0]);
    }

  var lr=inventory_sheet.getLastRow();
  for (i=2; i<=lr; i++){
    var minusone = i-1;

    var inventory_barcode = inventory_sheet.getRange(i,1,1,1).getValue();
    var shelflist_index = shelflist.indexOf(inventory_barcode);

    if (shelflist_index == -1) {
      var shelflist_index = 'no match';
    }
    inventory_sheet.getRange('K' + i).setValue(shelflist_index);
    var current_value = parseInt(inventory_sheet.getRange('K' + i).getValue());
    if (shelflist_index !== 'no match'){
    var previous_value = parseInt(inventory_sheet.getRange('K' + minusone).getValue());
    }
    //console.log('current value: ' + current_value );
    //console.log('previous value: ' + previous_value );
    var current_title = inventory_sheet.getRange('B' + i).getValue();
    var previous_title = inventory_sheet.getRange('B' + minusone).getValue();

    if ( (current_value < previous_value) && current_title !== previous_title ) {
      console.log("bingo misshelvo");
      inventory_sheet.getRange('A' + i + ':K' + i).setBackground('Yellow');
    }//endif
  }//end for
}//end findMisshelvings

