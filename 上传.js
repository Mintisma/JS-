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

  // 长度变量定义
  var Ivals = ss.getRange("I1:I").getValues();
  var Ilast = Ivals.filter(String).length + 1;
  var Kvals = ss.getRange('K1:K').getValues();
  var Klast = Kvals.filter(String).length + 1;
  var rowCount = sheet.getLastRow();

  // 范围定义
  var range_recordType = 'B4:B' + rowCount;
  var range_campaign = 'C4:C' + rowCount;
  var range_adGroup = 'H4:H' + rowCount;
  var range_matchType = 'K' + Klast + ':K' + rowCount;
  var range_enable = 'M4:O' + rowCount;
  var range_bid = 'I' + Ilast + ':I' + rowCount;

  // match type & bid
  var matchType = Browser.inputBox('match type');
  var groupBid = sheet.getRange('I3').getValue();

  // bid adjustment
  if (matchType=='Broad'){
    var bid = groupBid * 0.9;
  }
  else if(matchType=='Exact'){
    var bid = groupBid;
  }
  else {
    ui.alert('Please Enter Exact or *Broad*');
  }
  // 自动填充
  sheet.getRange(range_recordType).setValue('Keyword');
  sheet.getRange(range_campaign).setValue(sheet.getRange('C3').getValue());
  sheet.getRange(range_adGroup).setValue(sheet.getRange('H3').getValue());
  sheet.getRange(range_matchType).setValue(matchType);
  sheet.getRange(range_enable).setValue('enabled');
  sheet.getRange(range_bid).setValue(bid);
}