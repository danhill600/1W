// @ts-nocheck

var scriptProperties = PropertiesService.getScriptProperties();

function Otterize() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [];

    menuEntries.push({name: "Produce Shelflist", functionName: "Cauterize"});
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
    spreadsheet.getRange('I2').setValue(accesstoken);

    scriptProperties.setProperty('accesstoken', accesstoken)
    scriptProperties.setProperty('spreadsheet', spreadsheet)

}//end Otterize

function Cauterize() {
  var accesstoken = scriptProperties.getProperty('accesstoken');
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 // var spreadsheet = scriptProperties.getProperty('spreadsheet');

  var url = 'https://librarycatalog2.ccc.edu:443/iii/sierra-api/v6/items/query?offset=0&limit=3000';
  var starting = 'ac 8.a59'
  var ending = 'bl2755.3h58'
  var options = {
   "method" : "POST",
   "headers" : {
       "Authorization" : "Bearer " + accesstoken
     },
   "contentType" : "raw",
   "payload" : '{"queries":[{"target":{"record":{"type":"item"},"id":79},"expr":{"op":"equals","operands":["trsta",""]}},"and",{"target":{"record":{"type":"bib"},"field":{"tag":"c"}},"expr":{"op":"between","operands":["ac 8.a59","bl2755.3h58"]}}]}'
  };
    
  let row = 2;
  var result = UrlFetchApp.fetch(url, options);
  var json_data = JSON.parse(result.getContentText());
  for (let i = json_data.entries.length -1; i >= 0; i--){
    let itemID = json_data.entries[i].link.split('/')[7].split("\"")[0];
    console.log(itemID);
    
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
    if(json_data) {
      spreadsheet.getRange('C' + row).setValue(in_cn);
      spreadsheet.getRange('D' + row).setValue(out_cn)
      spreadsheet.getRange('E' + row).setValue(anotherjson_data.status.display);
      spreadsheet.getRange('F' + row).setValue(anotherjson_data.location.code);
      spreadsheet.getRange('G' + row).setValue('=\"' + Utilities.formatDate(new Date(), "GMT-4:00", "yyyy-MM-dd' 'HH:mm:ss") + '\"')
      
      var bibId = anotherjson_data.bibIds[0];
      var url = 'https://librarycatalog2.ccc.edu/iii/sierra-api/v5/bibs/' + bibId;
      var result = UrlFetchApp.fetch(url, options);
      var yetanotherjson_data = JSON.parse(result.getContentText());
      spreadsheet.getRange('A' + row).setValue(yetanotherjson_data.title);
      spreadsheet.getRange('B' + row).setValue(yetanotherjson_data.author);
    }//endif
    
   row++
  
  }//endforloop
    var range = spreadsheet.getDataRange();
    range.sort(4);
}//end Cauterize
