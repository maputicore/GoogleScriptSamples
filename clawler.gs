function getRangeByPosition(x,y) {
  var alphabets = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];
  
  if (!us._isNumber(x)){
    throw new Error('x is not number');
  }
  
  if (!us._isNumber(y)) {
    throw new Error('y is not number');
  }
  
  if (x > 26) {
    alert("x is out of range["+x+"]");
    return;
  }
  
  cell = alphabets[(x-1)] + y
  
  var objSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var objSheet = objSpreadsheet.getActiveSheet();
  return objSheet.getRange(cell);
  
}

function getValueForCell(x,y) {
  if (x == null) x = 1;
  if (y == null) y = 1;
 
  objRange = getRangeByPosition(x, y);
  return objRange.getValue();
}
function setValueForCell(x, y, value) {
  if (x == null) x = 1;
  if (y == null) y = 1;
 
  var objRange = getRangeByPosition(x, y);
  objRange.setValue(value);
}

function checkLink(urls) {
  str = urls.join("|");
  return (str.indexOf("nofollow") > -1);
}

function updateCell(x, y, options) {
  var text = (us._isString(options["text"]))?options["text"]: null,
      backgroundColor = (us._isString(options["background-color"]))? options["background-color"] : null;
  
  var objRange = getRangeByPosition(x, y);
  
  if (text !== null) {
    objRange.setValue(text);
  }

  if (backgroundColor != null) {
    objRange.setBackground(backgroundColor);
  }
  
}

/**
 *    Class HTTPResponse use
 *     
 */
function main() {
  //=== 初期化 ===
  var target_site_col = 5,   //E4から下が対象
      target_site_row = 4,
      target_site_max = 200, //対象サイトは200サイトまで
      expected_url_col = 6,  //F3から右が対象
      expected_url_row = 3,
  　　expected_url_max = 5,  //URLは5つまで
      list_link_col = 13,    //リンクのリストアップはM4から
      list_link_row = 4,
      error_col = 24,        // エラーはZ1
      error_row = 4;         
  var patterns = [],        // URLパターン用RegExインスタンス
      x, y, y2, y_error,      //行＆列イテレーション用パラメータ
      regexNoindex;
 
  Logger.log("start main");
  
  //Set notice of running status
  updateCell(1, 1, {
    'text': "実行中",
    'background-color': "yellow",
  });
  
 
  regexNoindex = new RegExp("/<meta[^>]*noindex.*?meta>/");
  // 調査URLを取得
  for (var i=0;i<expected_url_max;i++) {
    url = getValueForCell(expected_url_col+i, expected_url_row);
    if (url == "") {
      break; 
    } else {
      patterns[i] = new RegExp("<a [^>]*"+url+"[^>]*>.*?</a>", "g");
    }
  }
  
  // エラーの行
  y_error = error_row;
  
  // A列を起点にループ
  y2 = list_link_row;
  for (var i=0;i<target_site_max;i++) {
    setValueForCell(2,1,"checking No." + (i+1));
    try {
      // 対象サイトを取得    
      url = getValueForCell(target_site_col, target_site_row+i);
      media = getValueForCell(target_site_col - 2, target_site_row+i);
      title = getValueForCell(target_site_col - 1, target_site_row+i);
      date = getValueForCell(target_site_col - 3, target_site_row+i);
      
      url = url.replace(/ /, "");
      media = media.replace(/ /, "");
      title = title.replace(/ /, "");
      
      // URLがなければ終了
      if (url == "") break;   
      
      //座標をずらす
      y = parseInt(i) + target_site_row;
      
      // HTMLを取得
      var response = UrlFetchApp.fetch(url);
      body = response.getContentText();
      
      //404ならスキップ
      if (response.getResponseCode() != 200) continue;
      
      // metaタグ内のnoindexを調査
      noindex = "";
      metas = body.match(/<meta[^>]+>/g);
      if (metas.length != 0) {
        result = metas.join("").match("noindex");
        noindex = (result != null)? "◯" : "";  
      }
      
      setValueForCell(target_site_col + 2,y, noindex);
      
      //domainを調べる
      domain = url.match(/:\/\/([^\/\?]+)/);
      if (us._isArray(domain)) {
        domain = domain.pop()  
      } else {
        domain = url
      }

      // パターン毎にリンクタグをチェック
      for (j in patterns) {
        var pat = patterns[j],
            href = body.match(pat),
            active_links = [];
        
        if (href != null){
          Logger.log(href);
          x = parseInt(j) + expected_url_col;
          
          //nofollowを調べる
          if (checkLink(href)) {
            result = "◎";
            nofollow = "◯";
          } else {
            result = "◯";
            nofollow = "";
          }
       
          //見つかったリンクの書き出し(L-S列)
          for (k in href) {
            link_block = href[k];
            link_url = link_block.match(/(?!.+(pdf|zip|png)['"])href=['"]([^'"]+?)['"]/);
            if (!us._isArray(link_url)) { Logger.log(link_url); continue; }
            if (link_url.length == 0) continue;
            link_url = link_url.pop();
            link_query = '=HYPERLINK("' + link_url + '", "リンクを開く")';
            url_striped = url.replace(/^http.+?\/\//, "");
            
            active_links.push(url_striped);
            
            [(i+1),        // M列 No
             date,         // N列 日付
             media,        // O列 媒体名
             title,        // P列 記事タイトル
             url_striped,  // Q列 URL
             link_query,   // R列 リンクボタン
             domain,       // S列 ドメイン
             link_url,     // T列 リンク部分
             nofollow,     // U列　nofollow
             noindex,      // V列　noindex 
             link_block    // W列　デバッグブロック
             ].forEach (function(val, index) {
              setValueForCell(list_link_col + index, y2, val)              
            });
            y2 += 1;
          }
        }
        
        // リンクの有無とnofollowの確認
        // リンクあり (nofollowもあり)　=> ◎
        // リンクのみあり => ◯
        if (active_links.length == 0) continue;     
        Logger.log("x:"+x+" y:"+y+" is " + result + " [" + domain +"]");
        setValueForCell(x, y, result);
      }
    } catch(e) {
      error_msg = "Line." + i + " / " + e
      Logger.log(error_msg);
      updateCell(1, 1, {
        'text': "ER",
        'background-color': "red",
      });
      
      //setValueForCell(error_col, y_error, error_msg);
      y_error += 1;
    }
  }
  
  //Set notice of running status
  updateCell(1, 1, {
    'text': "終了",
    'background-color': "green",
  }); 
}