// @ts-nocheck
var scriptProperties = PropertiesService.getScriptProperties();

function Otterize() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [];

    menuEntries.push({name: "Get ItemIds in Range", functionName: "Cauterize"});
    menuEntries.push({name: "Get info for ItemIds", functionName: "getInfo"});
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

    scriptProperties.setProperty('accesstoken', accesstoken)
    scriptProperties.setProperty('spreadsheet', spreadsheet)

}//end Otterize

function Cauterize() { // puts itemID's in a range into column H
  var accesstoken = scriptProperties.getProperty('accesstoken');
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var starting = Browser.inputBox("Starting Call Number");
  var ending = Browser.inputBox("Ending Call Number");

  var url = 'https://librarycatalog2.ccc.edu:443/iii/sierra-api/v6/items/query?offset=0&limit=3000';
  var options = {
   "method" : "POST",
   "headers" : {
       "Authorization" : "Bearer " + accesstoken
     },
   "contentType" : "raw",
   "payload" : '{"queries":[{"target":{"record":{"type":"item"},"id":79},"expr":{"op":"equals","operands":["trsta",""]}},"and",{"target":{"record":{"type":"bib"},"field":{"tag":"c"}},"expr":{"op":"between","operands":["'+starting+'","'+ending+'"]}}]}'
  };
    
  let row = 2;
  var result = UrlFetchApp.fetch(url, options);
  var json_data = JSON.parse(result.getContentText());
  for (let i = json_data.entries.length -1; i >= 0; i--){
    let itemID = json_data.entries[i].link.split('/')[7].split("\"")[0];
    spreadsheet.getRange('H' + row).setValue(itemID);
    row++
  }//endforloop
}//end Cauterize

function getInfo() {
  var accesstoken = scriptProperties.getProperty('accesstoken');
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var lr=spreadsheet.getLastRow()

  for (i=2; i<lr; i++){
    if (( spreadsheet.getRange('A'+i).isBlank() ) && (i != 2)){
      var firstBlank = i-1;
      break;
    }//end if
  }// end for loop

  for (i=firstBlank; i<lr+1;i++) {
    if (!spreadsheet.getRange('H' + i).isBlank()){

      var itemID = spreadsheet.getRange('H'+i).getValue();
      var url = 'https://librarycatalog2.ccc.edu/iii/sierra-api/v5/items/' + itemID;

      var options = {
        "method" : "GET",
        "headers" : {
        "Authorization" : "Bearer " + accesstoken
      }
      };

      var result = UrlFetchApp.fetch(url, options);
      var anotherjson_data = JSON.parse(result.getContentText());
      var in_cn = anotherjson_data.callNumber;
      var loc = new locCallClass;
      var out_cn = loc.returnNormLcCall(in_cn);
      if(anotherjson_data) {
        spreadsheet.getRange('C' + i).setValue(in_cn);
        spreadsheet.getRange('D' + i).setValue(out_cn)
        spreadsheet.getRange('E' + i).setValue(anotherjson_data.status.display);
        spreadsheet.getRange('F' + i).setValue(anotherjson_data.location.code);
        spreadsheet.getRange('G' + i).setValue('=\"' + Utilities.formatDate(new Date(), "GMT-6:00", "yyyy-MM-dd' 'HH:mm:ss") + '\"')

        var bibId = anotherjson_data.bibIds[0];
        var url = 'https://librarycatalog2.ccc.edu/iii/sierra-api/v5/bibs/' + bibId;
        var result = UrlFetchApp.fetch(url, options);
        var yetanotherjson_data = JSON.parse(result.getContentText());
        spreadsheet.getRange('A' + i).setValue(yetanotherjson_data.title);
        spreadsheet.getRange('B' + i).setValue(yetanotherjson_data.author);
      }//end second if
    }//end first if
  }//endfor
    var range = spreadsheet.getDataRange();
    range.sort(4);
}//end getInfo
