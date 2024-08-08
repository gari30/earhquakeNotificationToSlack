/**
 * 気象庁の地震情報(高頻度フィード)を取得してslackに投げる
 * http://xml.kishou.go.jp/xmlpull.html
 */
function getEarthquakeAndVolcanoFeed() {
  const now_date = Utilities.formatDate(new Date(),"JST","yyyy-MM-dd HH:mm:ss");
  Logger.log("スクリプト開始時刻: " + now_date);

  // 前回スクリプト実行時刻の取得(スプレッドシートから取得)
  const sheet = SpreadsheetApp.getActiveSheet();
  if(sheet.getRange(2, 1).getValue() == ''){
    // 空の場合は現在時刻を設定
    var pre_date_text = Utilities.formatDate(new Date(), "JST", "yyyy-MM-dd HH:mm:ss");
  }else{
    var pre_date_text = Utilities.formatDate(sheet.getRange(2, 1).getValue(), "JST", "yyyy-MM-dd HH:mm:ss");
    Logger.log("前回スクリプト実行時刻: " + pre_date_text);
  }
  const pre_date = new Date(pre_date_text);
  pre_date.setMinutes(pre_date.getMinutes() - 2);   // 前回のスクリプト実施時間と、地震発生時間が同じだった場合にもフォローできるよう2分引く

  const atom = XmlService.getNamespace('http://www.w3.org/2005/Atom');
  const feed_url = 'https://www.data.jma.go.jp/developer/xml/feed/eqvol_l.xml';  // feedのURL

  Logger.log("気象情報xmlの取得開始");
  const feed_res = UrlFetchApp.fetch(feed_url).getContentText();
  Logger.log("気象情報xmlの取得完了");
  
  const feed_doc = XmlService.parse(feed_res);
  const feed_xml = feed_doc.getRootElement();
  const feed_locs = getElementsByTagName(feed_xml, 'entry'); // xmlに含まれるentry要素を配列で取得する

  // 気象庁のxmlデータ(feed)より、entry要素を繰り返し探索
  var quake_info = "";  //地震情報
  var quake_data_repote_time;  // 地震発生時刻
  var quake_area = "";  // 地震地域
  var quake_max_int = "";   // 震度
  var one_previous_quake_data_reporte_time = 0;   // 1つ前に通知した地震発生時刻

  // xmlが取得できてから、スクリプト実行時刻の更新をする(スプレッドシートを更新)
  sheet.getRange(2, 1).setValue(now_date);
  Logger.log("スクリプト実行時刻: " + now_date);

  feed_locs.forEach(function(value, i) {
    
    const title_text = value.getChild('title', atom).getText();  // title
    const link_text = value.getChild('link', atom).getAttribute('href').getValue();  // link
    Logger.log("title " + title_text);

    // updatedを取得
    const updated_text = value.getChild('updated', atom).getText(); // updated
    const updated = Utilities.formatDate(new Date(updated_text), "JST", "yyyy-MM-dd HH:mm:ss");
    Logger.log("updated " + updated);

    // titleが震度速報の場合(震度5以上)
    if('震度速報' == title_text) {
      Logger.log("link: " + link_text);
      
      //気象庁のxmlデータ(data)の情報を取得
      const data_url = link_text;
      const data_res = UrlFetchApp.fetch(data_url).getContentText();
      const data_doc = XmlService.parse(data_res);
      const data_xml = data_doc.getRootElement();
      const data_locs = getElementsByTagName(data_xml, 'Item'); // xmlに含まれるItem要素を配列で取得する
      var quake_data_repote_time_text = getElementsByTagName(data_xml, 'TargetDateTime')[0].getValue(); // 地震発生時刻
      quake_data_repote_time_text = quake_data_repote_time_text.replace('T', ' ').replace('+09:00', '');
      quake_data_repote_time = new Date(quake_data_repote_time_text);
      quake_area = getElementsByTagName(data_xml, 'Area')[0].getValue();
      quake_max_int = getElementsByTagName(data_xml, 'Kind')[0].getValue();

      //気象庁のxmlデータ(data)より、Item要素を繰り返し探索
      data_locs.forEach(function(value, i) {
        
        //Item要素の文字列を取得(震度と地域)
        const str_item = value.getValue();
        Logger.log("item: " + str_item);
        Logger.log("地震発生時刻: " + quake_data_repote_time_text);
        
        //「地震発生時刻 > 前回スクリプト実行時刻」かつ「震度が５～９」の場合
        if (quake_data_repote_time >= pre_date) {
          Logger.log("「地震発生時刻」が「前回スクリプト実行時刻」より新しい");

          if (quake_data_repote_time > one_previous_quake_data_reporte_time) {
            Logger.log("この実行内で一つ前に通知した地震とは違う地震");

            if (str_item.match(/震度[５-９]/)){
              quake_info = quake_info + str_item;  //地震情報
              Logger.log("地震情報: " + quake_info);

              Logger.log("Googleフォーム作成");
              const google_form_url = createGoogleForm(quake_data_repote_time_text);
              const payload = {
                "token" : PropertiesService.getScriptProperties().getProperty("SLACK_TOKEN"),
                "channel" : PropertiesService.getScriptProperties().getProperty("SLACK_POST_CHANNEL_ID"),
                "text" : "<!channel> \n【安否確認】\n震度5以上の強い揺れを検知しました。\n" + 
                  "安否確認のため、全く問題ない人は :woman-gesturing-ok: 、問題がある人・会社への連絡事項がある人は :woman-gesturing-no: のスタンプを押して<" + google_form_url + "|こちら>からGoogleFormに記入してください。\n\n" + 
                  "【地震情報】\n" +
                  "地震発生時刻: " + quake_data_repote_time_text + "\n" +
                  "震度: " + quake_max_int + "\n" +
                  "地域: " + quake_area + "\n",
              };
              Logger.log("Slack通知");
              sendToSlack(payload);
              one_previous_quake_data_reporte_time = quake_data_repote_time;  // 通知済みの地震発生時刻を保存
            } else {
              Logger.log("震度がマッチしなかった");
            }
          } else {
            Logger.log("この実行内で一つ前に通知した地震より古い、または同じ地震");
          }
        } else {
          Logger.log("「地震発生時刻」が「前回スクリプト実行時刻」より古かった");
        }
      });
    }
  });
  Logger.log("スクリプト終了時刻: " + Utilities.formatDate(new Date(),"JST","yyyy-MM-dd HH:mm:ss"));
}

/**
 * @param {string} element 検索要素
 * @param {string} tagName タグ
 * @return {string} data 要素
 */
function getElementsByTagName(element, tagName) {
  var data = [], descendants = element.getDescendants();
  for(var i in descendants) {
    var elem = descendants[i].asElement();
    if ( elem != null && elem.getName() == tagName) data.push(elem);
  }
  return data;
}

/**
 * slackにメッセージを送信する
 */
function sendToSlack(payload) {
  const slack_post_url = PropertiesService.getScriptProperties().getProperty("SLACK_MESSAGE_POST_URL");

  const params = {
    "method" : "post",
    "payload" : payload
  };
  Logger.log("slack payload: " + payload);

  // Slackに投稿する
  UrlFetchApp.fetch(slack_post_url, params);
  Logger.log("送信完了");
}

/**
 * Googleフォームを作成してURLを返す
 */
function createGoogleForm(quake_date) {
  const title = "安否確認フォーム_" + quake_date;
  var form = FormApp.create(title);

  // formファイルの保存場所設定
  const current_file_id = SpreadsheetApp.getActiveSpreadsheet().getId();
  const parent_folder = DriveApp.getFileById(current_file_id).getParents().next();
  const form_id = DriveApp.getFileById(form.getId());
  form_id.moveTo(parent_folder);  // カレントディレクトリにformファイルを移動

  // フォームの質問を追加
  // #1 氏名
  form.addTextItem().setTitle('氏名').setRequired(true);

  // #2 あなたの状態
  form.addMultipleChoiceItem().setTitle('あなたの状態')
                              .setChoiceValues([
                                '無事',
                                '重症 (骨折など動けない状態)',
                              ])
                              .showOtherOption(true)
                              .setRequired(true);
  // #3 いまどこにいますか
  form.addMultipleChoiceItem().setTitle('今どこにいますか')
                              .setChoiceValues([
                                '自宅',
                                'オフィス',
                                '通勤途中',
                              ])
                              .showOtherOption(true)
                              .setRequired(true);
  // #4 ご家族の状態
  form.addMultipleChoiceItem().setTitle('ご家族の状態')
                              .setChoiceValues([
                                '無事',
                                '重症者がいる (骨折など動けない状態)',
                                '重体者がいる (命に係わる状態)',
                                'わからない・確認中',
                              ])
                              .showOtherOption(true)
                              .setRequired(true);
  // #5 自宅の状況
  form.addMultipleChoiceItem().setTitle('自宅の状況')
                              .setChoiceValues([
                                '無事、もしくは軽微な損壊、自宅で仕事ができる',
                                '半壊、全壊で住めない、自宅で仕事はできない',
                                'わからない・確認中',
                              ])
                              .showOtherOption(true)
                              .setRequired(true);
  // #6 就業可否
  form.addMultipleChoiceItem().setTitle('就業可否')
                              .setChoiceValues([
                                '問題ない',
                                '相談したい →次の項目に詳細を記入してください',
                              ])
                              .setRequired(true);
  // #7 就業について相談
  form.addParagraphTextItem().setTitle('就業について相談がある方は詳細を記入してください(状況・環境など)');

  // #8 連絡先
  form.addTextItem().setTitle('連絡先(電話番号) SmartHRに登録の番号と違う場合は記入してください');

  // #9 その他連絡事項
  form.addParagraphTextItem().setTitle('その他連絡事項があれば記入してください');
  
  // フォームリンクを返却
  return form.getPublishedUrl();
}
