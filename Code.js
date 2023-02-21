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
    //spreadsheet.getRange('I2').setValue(accesstoken);

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
    e.range.offset(0,8).setValue(itemID);

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
      e.range.offset(0,6).setValue(json_data.location.code);
      e.range.offset(0,7).setValue('=\"' + Utilities.formatDate(new Date(), "GMT-4:00", "yyyy-MM-dd' 'HH:mm:ss") + '\"')
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
    spreadsheet.getRange('I' + row).setValue(itemID);
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
      console.log(firstBlank)
      break;
    }//end if
  }// end for loop

  for (i=firstBlank; i<lr+1;i++) {
    if (!spreadsheet.getRange('I' + i).isBlank()){

      var itemID = spreadsheet.getRange('I'+i).getValue();
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
        spreadsheet.getRange('G' + i).setValue(anotherjson_data.location.code);
        spreadsheet.getRange('H' + i).setValue('=\"' + Utilities.formatDate(new Date(), "GMT-6:00", "yyyy-MM-dd' 'HH:mm:ss") + '\"');

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
    range.sort(4);
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
