/////////////////////////// Triggers configuration ///////////////////////////////

function CreateTimeDrivenTriggers(func) {

  // Trigger every Monday at 09:00.
  var TRIG_ID = "TRIG_ID" + func
  var scriptProperties = PropertiesService.getScriptProperties();
  var triggerId = scriptProperties.getProperty(TRIG_ID);
  if (triggerId == null || triggerId == 0){

      var trigerid = ScriptApp.newTrigger(func)
          .timeBased()
          .onWeekDay(ScriptApp.WeekDay.MONDAY)
          .atHour(9)
          .create();
      var trig_id = trigerid.getUniqueId();

      scriptProperties.setProperty(TRIG_ID, trig_id);

  }else{
  
    var ui = SpreadsheetApp.getUi(); // Same variations.
    ui.alert('You have already create a trigger. Delete the trigger and try again');
    
  }
    
}
function DeleteTrigger(func) {

  var TRIG_ID = "TRIG_ID" + func
  var scriptProperties = PropertiesService.getScriptProperties();
  var triggerId = scriptProperties.getProperty(TRIG_ID);
  if (triggerId != null){

    var allTriggers = ScriptApp.getProjectTriggers();

    for (var i = 0; i < allTriggers.length; i++) {
      // If the current trigger is the correct one, delete it.
      if (allTriggers[i].getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(allTriggers[i]);
      }       
      break;
    }
    scriptProperties.deleteProperty(TRIG_ID);

  }

}
function SP_DeleteTrigger(){

  DeleteTrigger(SP_main)

}
function FB_DeleteTrigger(){

  DeleteTrigger(FB_main)

}
function ST_DeleteTrigger(){

  DeleteTrigger(ST_main)

}
function SP_SetTrigger(){

  CreateTimeDrivenTriggers(SP_main)

}

function FB_SetTrigger(){

  CreateTimeDrivenTriggers(FB_main)

}

function ST_SetTrigger(){

  CreateTimeDrivenTriggers(FB_main)

}

/////////////////////////// UI ///////////////////////////////
//Shopify:
function SP_config(){
  //Function description: Prepare the active spreedsheet for the addon: creete all necesary sheets and templates
  //Example:
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     'We are going to empty this SpreadSheet',
     'Do you want to continue?',
      ui.ButtonSet.YES_NO);
  

  if (result == ui.Button.YES){

    deleteSheets()
    SP_template_gen();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var url = ss.getUrl();

    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('URL_SP', url);

    sheet =ss.getSheetByName("Help");
    var protection = sheet.protect().setDescription('Protected sheet');
    protection.setWarningOnly(true)

    var res = ['Spreadsheet URL:',url];
    sheet.appendRow(res);

    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('URL', url);

    var headerRange = sheet.getRange(5, 1, 1,2 );
 
  // Apply each format to the top row: bold white text,
  // blue-green background, and a solid black border
  // around the cells.
    headerRange
      .setFontWeight('bold')
      .setFontColor('#ffffff')
      .setBackground('#007272')
      .setBorder(
        true, true, true, true, null, null,
        null,
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM);


  }
  var sheet_del = ss.getSheetByName("4656");
  ss.deleteSheet(sheet_del);
}

function SP_template_gen(){
  //Function description: generate SP template
  //Example:

  var SS = SpreadsheetApp.getActiveSpreadsheet();
  Data = SS.insertSheet();
  Help = SS.insertSheet();

  Help.setName("Help");
  Data.setName("Data");

  var path = "id,updated_at,name,total_price,fulfillment_status,financial_status,status,total_discounts,source_name,referring_site";																																																							
  var template = path.split(",");



  var sheet = SS.getSheetByName("Data");
  sheet.appendRow(template);
  var protection = sheet.protect().setDescription('Protected sheet');
  protection.setWarningOnly(true)

  var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());

  // Apply each format to the top row: bold white text,
  // blue-green background, and a solid black border
  // around the cells.  activeSheet.setFrozenRows(1);

  headerRange
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#007272')
    .setBorder(
      true, true, true, true, null, null,
      null,
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.setFrozenRows(1);



  var cont = [["API key:"],
             ["API password:"],
             ["Store url:"]];
  sheet = SS.getSheetByName("Help");
  var ran = sheet.getRange(2,1,3,1)
  ran.setValues(cont);
  ran
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#007272');

  ran = sheet.getRange(2,2,3,1)

    .setBorder(
      true, true, true, true, null, null,
      null,
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM) 
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#1155cc');

}

//Facebook:
function FB_config(){
  //Function description: Prepare the active spreedsheet for the addon: creete all necesary sheets and templates
  //Example:
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     'We are going to empty this SpreadSheet',
     'Do you want to continue?',
      ui.ButtonSet.YES_NO);
  

  if (result == ui.Button.YES){

    deleteSheets()
    FB_template_gen();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var url = ss.getUrl();


    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('URL_FB', url);

    sheet =ss.getSheetByName("Help");
    var protection = sheet.protect().setDescription('Protected sheet');
    protection.setWarningOnly(true)

    var res = ['Spreadsheet URL:',url];
    sheet.appendRow([0]);
    sheet.appendRow(res);

    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('URL', url);

    var headerRange = sheet.getRange(5, 1, 1,2 );
 
  // Apply each format to the top row: bold white text,
  // blue-green background, and a solid black border
  // around the cells.
  headerRange
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#007272')
    .setBorder(
      true, true, true, true, null, null,
      null,
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM);


  }
  var sheet_del = ss.getSheetByName("4656");
  ss.deleteSheet(sheet_del);
}

function FB_template_gen(){
  //Function description: generate Fb template
  //Example:

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  Ad_level = ss.insertSheet();
  AdSet_level = ss.insertSheet();
  Campaign_level = ss.insertSheet();
  Help = ss.insertSheet()

  Help.setName("Help");
  Ad_level.setName("Ad Level");
  AdSet_level.setName("AdSet Level");
  Campaign_level.setName("Campaign Level");
																																																								
  var sheets = ["Campaign Level","AdSet Level","Ad Level"];
  var template = [["date","campaign name","objective","impressions","clicks","spend","conversions","conversion values","cpc","cpm","ctr","frequency"],
  ["date","campaign name","adset name","clicks","impressions","cpm","ctr"],
  ["date","campaign name","adset name","ad name","clicks","impressions","cpm","ctr"]] ;

  for(var i=0;i<sheets.length;i++){

    var sheet = ss.getSheetByName(sheets[i]);
    sheet.appendRow(template[i]);
    var protection = sheet.protect().setDescription('Protected sheet');
    protection.setWarningOnly(true)

    var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
 
  // Apply each format to the top row: bold white text,
  // blue-green background, and a solid black border
  // around the cells.  activeSheet.setFrozenRows(1);

    headerRange
      .setFontWeight('bold')
      .setFontColor('#ffffff')
      .setBackground('#007272')
      .setBorder(
        true, true, true, true, null, null,
        null,
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      sheet.setFrozenRows(1);


  }

  var cont = [["Facebook Identification:"],
             ["Access Token:"],
             ["ID:"]];
  sheet = ss.getSheetByName("Help");
  var ran = sheet.getRange(1,1,3,1)
  ran.setValues(cont);
  ran
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#007272');

  ran = sheet.getRange(2,2,2,1)

    .setBorder(
      true, true, true, true, null, null,
      null,
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM) 
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#1155cc');

}

//Stripe:
function ST_config(){
  //Function description: Prepare the active spreedsheet for the addon: creete all necesary sheets and templates
  //Example:
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     'We are going to empty this SpreadSheet',
     'Do you want to continue?',
      ui.ButtonSet.YES_NO);
  

  if (result == ui.Button.YES){

    deleteSheets()
    ST_template_gen();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var url = ss.getUrl();


    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('URL_ST', url);

    sheet =ss.getSheetByName("Help");
    var protection = sheet.protect().setDescription('Protected sheet');
    protection.setWarningOnly(true)

    var res = ['Spreadsheet URL:',url];
    sheet.appendRow(res);

    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('URL', url);

  }
  var sheet_del = ss.getSheetByName("4656");
  ss.deleteSheet(sheet_del);
}

function ST_template_gen(){
  //Function description: generate Fb template
  //Example:

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  Data = ss.insertSheet();
  Help = ss.insertSheet()

  Help.setName("Help");
  Data.setName("Data");

																																																								
  var sheets = ["Data"];
  var template = [["Created",	"Available_on",	"id",	"Currency",	"Exchange rate",	"Fee",	"Net revenue",	"Status"]] ;

  for(var i=0;i<sheets.length;i++){

    var sheet = ss.getSheetByName(sheets[i]);
    sheet.appendRow(template[i]);
    var protection = sheet.protect().setDescription('Protected sheet');
    protection.setWarningOnly(true)

    var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
 
  // Apply each format to the top row: bold white text,
  // blue-green background, and a solid black border
  // around the cells.  activeSheet.setFrozenRows(1);

    headerRange
      .setFontWeight('bold')
      .setFontColor('#ffffff')
      .setBackground('#007272')
      .setBorder(
        true, true, true, true, null, null,
        null,
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      sheet.setFrozenRows(1);


  }

  var cont = [["Api Key"]];
  sheet = ss.getSheetByName("Help");
  var ran = sheet.getRange(2,1,1,1)
  ran.setValues(cont);

  ran = sheet.getRange(2,2,2,1)
  ran
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#007272');

  ran = sheet.getRange(2,1,2,1)

    .setBorder(
      true, true, true, true, null, null,
      null,
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM) 
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#1155cc');

}

//General:
function onInstall(e) {
  onOpen(e);
  // Perform additional setup as needed.
}
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.

   var ui = SpreadsheetApp.getUi();
   var menu = ui.createMenu("Data Extrator");

  menu.addSubMenu(ui.createMenu('Shopify')
        .addItem("Look for more data","SP_main")
        .addItem("Configuration","SP_config")
        .addItem("Set trigger","SP_SetTrigger")
        .addItem("Delete trigger","SP_DeleteTrigger"))

  menu.addSubMenu(ui.createMenu('Facebook')
      .addItem("Look for more data","FB_main")
      .addItem("Configuration","FB_config")
      .addItem("Set trigger","FB_SetTrigger")
      .addItem("Delete trigger","FB_DeleteTrigger"))

  menu.addSubMenu(ui.createMenu('Stripe')
      .addItem("Look for more data","ST_main")
      .addItem("Configuration","ST_config")
      .addItem("Set trigger","ST_SetTrigger")
      .addItem("Delete trigger","ST_DeleteTrigger"))

  menu.addToUi()

}

/////////////////////////// AUX ///////////////////////////////

function setCharAt(str,index,chr) {
  //Function description: replace the letter located at position index by chr
  //Example: setChartAt("hello",1,'o')-->>hollo

    if(index > str.length-1) return str;
    return str.substring(0,index) + chr + str.substring(index+1);
}
function deleteSheets(){
  //Function description: delete all sheets in the active spreadsheet, it only left one sheet named "4656"
  //Example:

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();


  Help = ss.insertSheet()
  var name = "4656";
  Help.setName(name);

  for(i in sheets){

    if(sheets[i].getName() != name)
    ss.deleteSheet(sheets[i]);
    

  }
  
}

/////////////////////////// Data Extraction Facebook Ads ///////////////////////////////

function FB_main(){


  for(var i=0;i<3;i++){ 
    
      var url = FB_url(i,0) 
      FB_api_call(url,0)
      url = FB_url(i,2) 
      var data = FB_api_call(url,1) 
      Logger.log(data)
      FB_writes(i,data['data'])
      url = FB_pagination(data['paging'],url)

      while(url != "end"){

        var data = FB_api_call(url,1) 
        FB_writes(i,data['data'])
        url = FB_pagination(data['paging'],url)

      }
  }



}
function FB_dates(i){
  //Function description:
  //Return the last date stored or maximun.
  //Considerations: cell A1 must be call date
  //  Input: i: the index of the sheet level: campaign, adset, ad
  //Date format: [2020, 11, 28, 2021, 05, 18] (first date is last stored date + 1day and last date is todays date)

  var userProperties = PropertiesService.getDocumentProperties();
  var URL = userProperties.getProperty('URL_FB');
  var SS = SpreadsheetApp.openByUrl(URL) 
  var sheets = ["Campaign Level","AdSet Level","Ad Level"];

  var sheet = SS.getSheetByName(sheets[i]);
  var lastRow = sheet.getLastRow();
  var last_date = sheet.getRange(lastRow, 1).getValue();
  var timeZone_sheet = SS.getSpreadsheetTimeZone(); // See Apps Script documentation

  

  if(last_date != "date"){ //check if we have data in the sheet

    last_date.setDate(last_date.getDate()+1);//add a date to the last stored data
    var last_date = Utilities.formatDate(last_date, timeZone_sheet, "yyy-MM-dd");
    var tdy_date = Utilities.formatDate(new Date(), timeZone_sheet, "yyy-MM-dd");//'2021-03-15'
    var last_date_s = last_date.split("-");// the first element is empty
    var tdy_date_s = tdy_date.split("-");


    if(last_date_s[0] < tdy_date_s[0]){ //if the last date year is less than todays year

      last_date_s.push(tdy_date_s[0]); 
      last_date_s.push(tdy_date_s[1]); 
      last_date_s.push(tdy_date_s[2]); 
      return last_date_s;

    }else if (last_date_s[1] < tdy_date_s[1]){ //if the last date year is todays year and last date month is less than todays month

      last_date_s.push(tdy_date_s[0]); 
      last_date_s.push(tdy_date_s[1]); 
      last_date_s.push(tdy_date_s[2]); 
      return last_date_s;

    }else if (last_date_s[2] < tdy_date_s[2]){ //if we have same month and same year and last date is less than todays day number

      last_date_s.push(tdy_date_s[0]);
      last_date_s.push(tdy_date_s[1]); 
      last_date_s.push(tdy_date_s[2]);  
      return last_date_s;

    }

  }else{// get lifespam data

    var data_range = "maximum";

    return data_range;

  }
}
function FB_url(i,sel){
  //Function description:
  //    Return: url.
  //     Input: i: the index of the sheet level: campaign, adset, ad(only matters for sel = 0)
  //          sel: the type of url: 0->>first call
  //                                1->>job completed      
  //                                2->>data call
  //

  var userProperties = PropertiesService.getDocumentProperties();
  var URL = userProperties.getProperty('URL_FB');
  var SS = SpreadsheetApp.openByUrl(URL) 

  var SHEET = SS.getSheetByName("Help");
  var TOKEN = SHEET.getRange(2 , 2).getValue();
  var ENTITY_ID = SHEET.getRange(3 , 2).getValue();

  switch (sel)  {
    case 0:
        var level = ["campaign","adset","ad"];
        var field = ["date_start,campaign_name,objective,impressions,clicks,spend,conversions,conversion_values,cpc,cpm,ctr,frequency","date_start,campaign_name,adset_name,clicks,impressions,cpm,ctr","date_start,campaign_name,adset_name,ad_name,clicks,impressions,cpm,ctr"];
    
        var fields = field[i];
        var date_range = FB_dates(i);

        if ( date_range.length == 6){

          var data_range_call = '&time_range={"since":"'
              + date_range[0] 
              + '-'
              + date_range[1]
              + '-'
              + date_range[2] 
              + '","until":"'
              + date_range[3] 
              + '-'
              + date_range[4]
              + '-'
              + date_range[5] 
              + '"}';

        }else{

          var data_range_call = '&date_preset=maximum';

        }
        var url = `https://graph.facebook.com/v10.0/act_${ENTITY_ID}/insights?level=${level[i]}&fields=${fields}${data_range_call}&time_increment=1&access_token=${TOKEN}&limit=50`;//

        break;

    case 1:
        var cache = CacheService.getScriptCache();
        var reportId = cache.get('campaign-report-id');

        var url = `https://graph.facebook.com/v10.0/${reportId}/?access_token=${TOKEN}&fields=async_status`;

        break;

    case 2:
        var cache = CacheService.getScriptCache();
        var reportId = cache.get('campaign-report-id');

        url = `https://graph.facebook.com/v10.0/${reportId}/insights?access_token=${TOKEN}`;
        break;

    default:

        break;
  }
  
  

  Logger.log("Output: FB_url():  " + url)
  return url

}
function FB_api_call(url,sel){
  //Function description:
  //    Return: sel = 0->> data.
  //            sel = 1->> nothing.
  //     Input: i: the index of the sheet level: campaign, adset, ad
  //          sel: the type of url: 0->>first call
  //                                1->>data call
  //

  var f = 0;//aux variable for validation the correct call
  switch (sel)  {
        case 0:
          while(f <= 5 ){

            var options = {
              'method' : 'post',
              'muteHttpExceptions' : true

            };
            
            // Fetches & parses the URL 
            url = encodeURI(url);
            var fetchRequest = UrlFetchApp.fetch(url, options);
            var code = JSON.parse(fetchRequest.getResponseCode());
            if (code == 200){
              f = 50;
            }else{
              f = f + 1;
            }
          }
          if (f == 50){

            var results = JSON.parse(fetchRequest.getContentText());
            // Caches the report run ID
            Logger.log(results)
            var reportId = results.report_run_id;
            var cache = CacheService.getScriptCache();
            var cached = cache.get('campaign-report-id');
            
            if (cached != null) {
              cache.put('campaign-report-id', [], 1);
              Utilities.sleep(1001);
              cache.put('campaign-report-id', reportId, 21600);
            } else {
              cache.put('campaign-report-id', reportId, 21600); 
            };
            
            Logger.log(cache.get('campaign-report-id'));
            
          }
            break;

        case 1:
            //First we need to verify that facebook has the data ready
            var g = 0;
            url_ver = FB_url(0,1);//First argument doesnt matter
            while(g <= 20 ){

                var options = {
                'muteHttpExceptions' : true
                };
                
                var data = UrlFetchApp.fetch(url_ver,options);
                var code = JSON.parse(data.getContentText());
                if (code['async_status'] == 'Job Completed'){

                  g = 50;

                }else{

                  Utilities.sleep(1000);
                  Logger.log("waiting")

                }
            }

            while(f <= 5 ){

                var options = {
                'muteHttpExceptions' : true
                };
                var data = UrlFetchApp.fetch(url,options);
                var code = JSON.parse(data.getResponseCode());

                if (code == 200){
                  f = 50;
                }else{
                  f = f + 1;
                }
            }
            if (f == 50){
              data = JSON.parse(data.getContentText()); 
              return data;
            }
            break;
        default:
            break;
    }




  
}
function FB_pagination(pagination,url){
  //Function description:
  //Append the data store in the input: data
  //Inputs: DATA:  element ['paging'] from: JSON.parse(data.getContentText()); ({cursors={after=MTUZD, before=MAZDZD}})
  //         url: last url that did the call
  //           i: index of the sheet: Campaign, adset, ad
  //Output: url   -> next url to call
  //      : "end" -> end of pagination


  if (pagination != undefined){

    var next = pagination;

  }
  if (typeof(pagination.next) == 'undefined') {

    Logger.log("end of pagination")
    return("end")

  }
  else {
    
    var comp = url.substring(url.length-6,url.length-12);

    if (comp == '&after'){

      url = url.substring(url.length-12,0) // quita la pagination
    }

    if(next!= null){

      next = pagination;
      next = next['cursors'];
      next = next['after'];
      url = url + '&after=' + next; //añade la pagination
      Logger.log('FB_pagination():' + url)
      return url;
    }
  }





}
function FB_writes(i,data){
  //Function description:
  //Append the data store in the input: data
  //Inputs: DATA:  element ['data'] from: JSON.parse(data.getContentText());
  //           i: index of the sheet: Campaign, adset, ad


  var userProperties = PropertiesService.getDocumentProperties();
  var URL = userProperties.getProperty('URL_FB');
  var SS = SpreadsheetApp.openByUrl(URL) 
  var sheets = ["Campaign Level","AdSet Level","Ad Level"];
  var SHEET = SS.getSheetByName(sheets[i]);

  var fields = ["date_start,campaign_name,objective,impressions,clicks,spend,conversions,conversion_values,cpc,cpm,ctr,frequency","date_start,campaign_name,adset_name,clicks,impressions,cpm,ctr","date_start,campaign_name,adset_name,ad_name,clicks,impressions,cpm,ctr"];

  var patharray = fields[i].split(",");
  //var patharraytwo = fields_second.split(",");
  

  //Data structure: array[campaign0 , campaignn1,...]
  //                campaigns is an object or an array


  var row = [0]; // each row that is going to be send to the sheet
  var cell = [0,0];


  for(var o=0;o<data.length;o++){ //extracting data campaign by capaign. 
  
      row = []; //delate all content from the row   
      var element = data[o]; // gets each campaign/adset or ad data   

      for(var j=0;j<patharray.length;j++){ //Bucle por cada elemento de las rows
      
        cell = element[patharray[j]];
      
        if (Number.isNaN(Number(cell)) == true){ // conver data (string format) to number format
          row.push(cell)                
        }else{
          row.push(Number(cell));
        }
      }
      SHEET.appendRow(row);
        
  }
}

/////////////////////////// Data Extraction Shopify ///////////////////////////////

function SP_dates(){
  //Function description: return the last date stored + 1 second.
  //Date format: 2021-02-12T22:23:45%2B01:00

  // Global variables:
  var userProperties = PropertiesService.getDocumentProperties();
  var URL = userProperties.getProperty('URL_SP');
  var SS = SpreadsheetApp.openByUrl(URL) 
  var HELP_S = SS.getSheetByName('Help');
  var DATA_S = SS.getSheetByName('Data');

  var timeZone = SS.getSpreadsheetTimeZone(); // See Apps Script documentation
  var lastRow = DATA_S.getLastRow();
  var last_date = DATA_S.getRange(lastRow, 2).getValue();

  last_date = new Date(last_date);
  var last_date = Utilities.formatDate(last_date,timeZone,"yyyy-MM-dd'T'HH:mm:ssXXX");
  //Adding 1 second:
  var seconds = [Number(last_date[17]),Number(last_date[18])]
  Logger.log(seconds)
  if (seconds[1] == 9){
    seconds[1] = 0
    seconds[0] = seconds[0] + 1
  }else{
    seconds[1] = seconds[1] + 1
  }
  last_date = setCharAt(last_date,17,seconds[0]);
  last_date = setCharAt(last_date,18,seconds[1]);

  last_date = last_date.split("+")
  var date = last_date[0] + "%2B" + last_date[1]
  return date

}
function SP_url(){
  //Function description: created the Shopify url fields
  //Return: url
  // Global variables:
  var userProperties = PropertiesService.getDocumentProperties();
  var URL = userProperties.getProperty('URL_SP');
  var SS = SpreadsheetApp.openByUrl(URL) 
  var HELP_S = SS.getSheetByName('Help');


  var xpath = "id,updated_at,name,total_price,fulfillment_status,financial_status,status,total_discounts,source_name,referring_site";
  var data_range = SP_dates();
  var SHOP_ID = HELP_S.getRange(4, 2).getValue() 

  var url = "https://" + SHOP_ID + '/admin/api/2020-10/orders.json?limit=20&order=updated_at asc&updated_at_min=' + data_range + '&fields=' + xpath;


  return url;
}
function SP_api_call(url){
  
  // Global variables:
  var userProperties = PropertiesService.getDocumentProperties();
  var URL = userProperties.getProperty('URL_SP');
  var SS = SpreadsheetApp.openByUrl(URL) 
  var HELP_S = SS.getSheetByName('Help');

  Logger.log(url)
  var API_KEY = HELP_S.getRange(2 , 2).getValue() 
  var API_PASSWORD = HELP_S.getRange(3, 2).getValue() 
  var f = 0

  while( f <= 5 ){

    var encoded = Utilities.base64Encode(API_KEY + ':' + API_PASSWORD);

    var headers = {
        "Content-Type" : "application/json",
        "Authorization": "Basic " + encoded
      };

    var options = {
      "contentType" : "application/json",
      'method' : 'GET',
      'headers' : headers, 
      'followRedirects' : false,
    };

  
    var response = UrlFetchApp.fetch(url,options);

    var RESPONSE_CODE = response.getResponseCode();
    var RESPONSE_HEADERS = response.getHeaders();

    Logger.log(RESPONSE_CODE)

    if (RESPONSE_CODE == 200){

      f = 40; //get our of the verification process

      var CONTENT_JSON = JSON.parse(response.getContentText());
      Logger.log(CONTENT_JSON)
      SP_write(CONTENT_JSON)
      return(RESPONSE_HEADERS);

    }else{

      f = f + a
    }

    if (RESPONSE_CODE == 401){

      var ui = SpreadsheetApp.getUi(); // Same variations.
      ui.alert('Did you enter the correct API keys and Passwords?');

    }

  }

}
function SP_write(data){

  // Global variables:
  var userProperties = PropertiesService.getDocumentProperties();
  var URL = userProperties.getProperty('URL_SP');
  var SS = SpreadsheetApp.openByUrl(URL) 
  var DATA_S = SS.getSheetByName('Data');
  var xpath = "id,updated_at,name,total_price,fulfillment_status,financial_status,status,total_discounts,source_name,referring_site";

  var patharray = xpath.split(",");
  data = data['orders'];

  //Variable declaration:
  var row = [];
  for(var j=0;j<patharray.length;j++){

      row = row.concat([0]);
  
  }
  
  for(var i=0;i<data.length;i++){

    var con = data[i]
  
    for(var j=0;j<patharray.length;j++){

     
        if (Number.isNaN(Number(con[patharray[j]])) == true){ // conver data (string format) to number format

          row[j] = con[patharray[j]];             

        }else{

          row[j] = Number(con[patharray[j]]);
        
        }

    }
    DATA_S.appendRow(row);
    row = [];
      
  }

  
}
function SP_pagination(headers){


  var links = headers["Link"].split(",")
  var a = [,]

  for (var j = 0; j< links.length; j++){
    a[j] = links[j].split(";")
  } 

  if (a[0][1] == ' rel="next"'){// check if the next link is in the first row

    var link = a[0][0];
    var l = 1;
    var final_link = []
    while(link[l] != ">"){

      final_link.push(link[l]) 
      l = l + 1;
    }
    final_link = final_link.join("")  
    return final_link;

  }else if(a.length == 2){// check if the next link is in the second row
      if( a[1][1] == ' rel="next"'){

        var link = a[1][0];
        var l = 1;
        var final_link = []
        while(link[l] != ">"){

          final_link.push(link[l]) 
          l = l + 1;
        }
        final_link = final_link.join("")  
        return final_link;

      }
  }else{// if rell = next is not found->return 0

      return 0;

  }


}
function SP_main(){

  var page = 0;
  
  var url = SP_url();

  while( page == 0){

    headers = SP_api_call(url);

    if (headers["Link"] == null){

     url = 0;

    }else{

      url = SP_pagination(headers);

    }

    if (url == 0){

      page = 1;
    }


  }

}

/////////////////////////// Data Extraction Stripe ///////////////////////////////

function ST_dates(){
  //Function description: return the newest data stored + 1 second.
  //Date format: Unix time format

  // Global variables:
  var userProperties = PropertiesService.getDocumentProperties();
  var URL = userProperties.getProperty('URL_ST');
  var SS = SpreadsheetApp.openByUrl(URL) 
  var HELP_S = SS.getSheetByName('Help');
  var DATA_S = SS.getSheetByName('Data');

  var newst_date = DATA_S.getRange(2, 1).getValue();

  if (newst_date == ''){

    return ''

  }else{
    
    var newst_date_abs = 0;
    for(i=2;i<=DATA_S.getLastRow();i++){

      newst_date = DATA_S.getRange(i, 1).getValue();
      Logger.log(newst_date)

      if(newst_date >= newst_date_abs){

          newst_date_abs = newst_date 

      }


    }

    newst_date_abs = newst_date_abs + 1;//Adding 1 second:
    var date = '?created%5Bgte%5D=' + newst_date_abs;

    return date
  }



}

function ST_url(){
  //Function description: created the Shopify url fields
  //Return: url

  var data_range = ST_dates();

  var url = "https://api.stripe.com/v1/balance_transactions" + data_range;
  return url;
}

function ST_api_call(url){
  
  // Global variables:
  var userProperties = PropertiesService.getDocumentProperties();
  var URL = userProperties.getProperty('URL_ST');
  var SS = SpreadsheetApp.openByUrl(URL) 
  var HELP_S = SS.getSheetByName('Help');
  var API_KEY = HELP_S.getRange(2 , 2).getValue() 

  var f = 0

  while( f <= 5 ){

    var encoded = Utilities.base64Encode(API_KEY);

    var headers = {
        "Content-Type" : "application/json",
        "Authorization": "Basic " + encoded
      };

    var options = {
      "contentType" : "application/json",
      'method' : 'GET',
      'headers' : headers, 
      'followRedirects' : false,
       muteHttpExceptions: true,
    };

  
    var response = UrlFetchApp.fetch(url,options);

    var RESPONSE_CODE = response.getResponseCode();

    if (RESPONSE_CODE == 200){

      f = 40; //get our of the verification process
      Logger.log(response.getContentText())
      var CONTENT_JSON = JSON.parse(response.getContentText());
      return CONTENT_JSON;

    }else{

      f = f + 1
    }

    if (RESPONSE_CODE == 401){

      var ui = SpreadsheetApp.getUi(); // Same variations.
      ui.alert('Did you enter the correct API keys and Passwords?');

    }

  }

}

function ST_write(data){

  // Global variables:
  var userProperties = PropertiesService.getDocumentProperties();
  var URL = userProperties.getProperty('URL_ST');
  var SS = SpreadsheetApp.openByUrl(URL) 
  var DATA_S = SS.getSheetByName('Data');
  var xpath = "created,available_on,id,currency,exchange_rate,fee,net,status";

  var patharray = xpath.split(",");
  var has_more = data['has_more'];
  Logger.log(has_more)
  data = data['data'];

  //Variable declaration:
  var row = [];
  for(var j=0;j<patharray.length;j++){

      row = row.concat([0]);
  
  }

  for(var i=0;i<data.length;i++){

    var con = data[i]
  
    for(var j=0;j<patharray.length;j++){

     
        if (Number.isNaN(Number(con[patharray[j]])) == true){ // conver data (string format) to number format

          row[j] = con[patharray[j]];             

        }else{

          row[j] = Number(con[patharray[j]]);
        
        }

    }
    row[6] = row[6]/100
    row[5] = row[5]/100

    DATA_S.appendRow(row);
    row = [];
      
  }
  if( has_more == true){

    return con['id']

  }else{

    return 0
  }

  
}

function ST_pagination(id,url){



  if(url[70] == null){

    var url_pag = url + '?starting_after=' + id;
    return url_pag

  }else{

    var url_pag = url + '&starting_after=' + id;
    return url_pag
  }


}

function ST_main(){

  
  var url = ST_url();
  var url_p = url
  var f = 0;
  var id = 1;

  while(f==0){

    data = ST_api_call(url_p);
    id = ST_write(data);

    if(id == 0){//end of pagination

      f = 6;

    }else{

      url_p = ST_pagination(id,url)

    }

  }


}


