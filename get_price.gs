function insertPrice(){
  const SHEET_ID = "yourSheetId";//yourSheetId
  const SHEET_NAME="price_table";
  const START_ROW = 4;
  const START_COLUMN = 2;
  const NUM_COLUMNS = 5;

  // "price_table"シートと最終行を取得
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var numRows = sheet.getDataRange().getLastRow() - START_ROW;

  // セルからデータ取得
  var values = sheet.getRange(START_ROW, START_COLUMN, numRows, NUM_COLUMNS).getValues();
  // データ用配列
  var itemArray = [];

  // getPriceして、配列に挿入
  for(var i = 0; i < values.length; i++){
    itemArray.push(getPrice(i + 1, values[i]));
  }  
  // セルにデータ挿入
  sheet.getRange(START_ROW, START_COLUMN, numRows, NUM_COLUMNS).setValues(itemArray);
}


function getPrice(row, arry){
  var url = arry[0];

  if(url.indexOf("https://kakaku.com") != -1){
    var response = UrlFetchApp.fetch(url);
    
    if(response.getResponseCode() == 200){
      var html = response.getContentText('shift_jis');
      
      // 商品名を取得 ex)<h2 itemprop="name">DeskMini A300/B/BB/BOX/JP</h2>S
      var regexItemName = /<h2 itemprop="name">([\s\S]*?)<\/h2>/;
      var itemName = html.match(regexItemName)[1];

      // カテゴリ取得 ex)ctgname: 'SSD',
      var regexCategory = /ctgname\: \'([\s\S]*?)\'\,/;
      var category = html.match(regexCategory)[1];

      // 最安値を取得 ex)prdlprc: 18997,
      var regexPrice = /prdlprc\: ([0-9]+)\,/;
      var price = html.match(regexPrice)[1];
      
      // 配列作成して、戻り値用配列に挿入(urlが[0]なので1から)
      var item = [itemName, category, arry[4], price];      
      for(var i =1; i < arry.length; i++){
        arry[i] = item[i - 1];
      }

      return arry;
    }else{
      Browser.msgBox("URLにアクセスできませんでした。");
    }
  }else if(url != ""){
    Browser.msgBox(String(row) + "行目は価格.comのURLが入力されていません。");
  }
  return arry;
}
