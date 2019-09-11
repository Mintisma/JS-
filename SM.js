function onOpen(e) {
   SpreadsheetApp.getUi()
       .createMenu('广告上传快捷方式')
       .addItem('自动填充', 'autofill')
       .addToUi();
 }
 
 
 function autofill() {
     // 初始定义
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  var input = Browser.inputBox('bid & match_type');
  var base_bid = input.split(" ")[0];
  var match_type = input.split(" ")[1];
  
      // bid adjustment
  if (match_type=='broad'){
    var bid = base_bid * 0.9;
  }
  else if(match_type=='exact'){
    var bid = base_bid;
  }
  else {
    ui.alert('Please Enter exact or *broad*');
  }
 
     // 长度变量定义
  var Dvals = ss.getRange("D1:D").getValues();
  var Dlast = Dvals.filter(String).length + 1;
  var Fvals = ss.getRange('F1:F').getValues();
  var Flast = Fvals.filter(String).length + 1;
  var rowCount = sheet.getLastRow();
  
    // 范围定义
  var range_campaign = 'A3:A' + rowCount;
  var range_adGroup = 'B3:B' + rowCount;
  var range_matchType = 'F' + Flast + ':F' + rowCount;
  var range_enable = 'E3:E' + rowCount;
  var range_bid = 'D' + Dlast + ':D' + rowCount;
  var range_bid_type = 'G2:G' + rowCount;
  
    // 自动填充
  sheet.getRange(range_campaign).setValue(sheet.getRange('A2').getValue());
  sheet.getRange(range_adGroup).setValue(sheet.getRange('B2').getValue());
  sheet.getRange(range_matchType).setValue(match_type);
  sheet.getRange(range_enable).setValue('enabled');
  sheet.getRange(range_bid).setValue(bid);
  sheet.getRange(range_bid_type).setValue('manual');
 }
