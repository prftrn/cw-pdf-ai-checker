// --- 1. スプレッドシートに専用メニューを作る処理 ---
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 AI判定メニュー')
    .addItem('フォルダ内のPDFを判定する', 'processPdfInFolder')
    .addToUi();
}

// --- 2. フォルダ内のPDFを1枚ずつ読み込んで判定するメイン処理 ---
function processPdfInFolder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName("検索設定");
  var listSheet = ss.getSheetByName("案件リスト");

  var inputFolderUrl = "あなたの判定待ちフォルダURL";
  var outputFolderUrl = "あなたの判定済みフォルダURL";

  if (!inputFolderUrl || !outputFolderUrl) {
    Browser.msgBox("エラー", "フォルダのURLが設定シートに入力されていません。", Browser.Buttons.OK);
    return;
  }

  try {
    var inputFolderId = extractIdFromUrl(inputFolderUrl);
    var outputFolderId = extractIdFromUrl(outputFolderUrl);

    if (inputFolderId.indexOf("http") === 0) {
      Browser.msgBox("⚠️ URL形式エラー", "URLからIDを読み取れませんでした。ダブルクリックでフォルダの中に入った状態のURLをコピーしてください。", Browser.Buttons.OK);
      return;
    }

    var inputFolder = DriveApp.getFolderById(inputFolderId);
    var outputFolder = DriveApp.getFolderById(outputFolderId);
    var files = inputFolder.getFilesByType(MimeType.PDF); 

    var count = 0; 

    while (files.hasNext()) {
      var file = files.next();
      var fileName = file.getName();

      ss.toast(fileName + " の文字を抽出中...", "システム処理中");

      var pdfText = extractTextFromPdf(file);
      if (pdfText.indexOf("エラー発生") !== -1) {
         Browser.msgBox("🛑 文字抽出エラー", fileName + " の読み込みに失敗しました。\\n詳細:\\n" + pdfText, Browser.Buttons.OK);
         continue; 
      }

      ss.toast(fileName + " をAIに判定させています...", "システム処理中");
      
      var aiResult = callGeminiAPI(fileName, pdfText);

      if (aiResult) {
        listSheet.insertRowBefore(2);
        var now = new Date();
        var timeString = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");

        listSheet.getRange(2, 1, 1, 6).setValues([[
          timeString, aiResult["判定"], fileName, aiResult["業務内容"], aiResult["報酬額"], aiResult["判定理由"]
        ]]);

        count++;
        file.moveTo(outputFolder);
        Utilities.sleep(1500); 
      }
    }

    Browser.msgBox("実行完了", count + " 件のPDFのAI自動判定が完了しました！", Browser.Buttons.OK);

  } catch (e) {
    Browser.msgBox("重大なエラー", "処理中にエラーが発生しました。詳細:\\n" + e.toString(), Browser.Buttons.OK);
  }
}

// --- 3. URLからIDを自動抽出 ---
function extractIdFromUrl(url) {
  var urlStr = url.toString().trim();
  var match1 = urlStr.match(/\/folders\/([a-zA-Z0-9-_]+)/);
  if (match1 && match1[1]) return match1[1];
  var match2 = urlStr.match(/id=([a-zA-Z0-9-_]+)/);
  if (match2 && match2[1]) return match2[1];
  return urlStr;
}

// --- 4. PDF文字抽出 ---
function extractTextFromPdf(file) {
  try {
    var blob = file.getBlob();
    var tempFile;
    
    if (typeof Drive.Files.create === "function") {
      var resourceV3 = { name: "temp_ocr_" + file.getName(), mimeType: MimeType.GOOGLE_DOCS };
      tempFile = Drive.Files.create(resourceV3, blob);
    } else if (typeof Drive.Files.insert === "function") {
      var resourceV2 = { title: "temp_ocr_" + file.getName(), mimeType: MimeType.PDF };
      tempFile = Drive.Files.insert(resourceV2, blob, { ocr: true, ocrLanguage: "ja" });
    } else {
      throw new Error("Drive APIが有効ではありません。");
    }
    
    var doc = DocumentApp.openById(tempFile.id);
    var rawApiKey = "YOUR_API_KEY_HERE";
    
    DriveApp.getFileById(tempFile.id).setTrashed(true);
    return text;
  } catch (e) {
    return "エラー発生: " + e.toString();
  }
}

// --- 5. Gemini AI判定（原因となったモデル名を元の正しい形に修正） ---
function callGeminiAPI(fileName, pdfText) {
  var rawApiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  
  if (!rawApiKey) {
    Browser.msgBox("🛑 APIキー未取得", "プロパティに「GEMINI_API_KEY」がありません。", Browser.Buttons.OK);
    return null;
  }

  var apiKey = rawApiKey.trim();

  if (apiKey.indexOf("AIza") !== 0) {
    Browser.msgBox("🛑 APIキー形式エラー", "設定されたAPIキーが「AIza」から始まっていません。再確認してください。", Browser.Buttons.OK);
    return null;
  }

  // 【ここが唯一の修正点】 -latest を削除し、絶対に繋がる標準モデルに戻しました
  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;

  var systemPrompt = "あなたはクラウドワークスの案件精査のプロです。以下の【絶対防衛ルール】に従い判定してください。結果は必ず、改行を含まない1行のJSONデータのみを出力してください。\n\n" +
                     "【絶対防衛ルール】\n" +
                     "・【Sランク：即応募推奨】: 一般的な日本語でゴールが明確、固定報酬（3千〜1万円）、外部チャット誘導なし。\n" +
                     "・【Aランク：応募候補】: 条件は良いが一部手動確認を推奨。\n" +
                     "・【Bランク：要検討】: 「継続的」「随時追加」など副業としては負担が大きい。\n" +
                     "・【Cランク：絶対スルー】: 「スクレイピング」「API」等ITプロ向け用語あり、または外部チャット誘導ありの危険案件。\n\n" +
                     "【出力フォーマット】\n" +
                     "{\"判定\": \"S〜Cランク\", \"業務内容\": \"30文字以内\", \"報酬額\": \"固定〇〇円など\", \"判定理由\": \"100文字以内の理由\"}";

  var payload = {
    "contents": [{ "parts": [{ "text": systemPrompt + "\n\n【PDF名】" + fileName + "\n\n【本文】\n" + pdfText.substring(0, 5000) }] }],
    "generationConfig": { "responseMimeType": "application/json" }
  };

  var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
       Browser.msgBox("🛑 通信エラー (コード: " + responseCode + ")", "Gemini APIへの通信に失敗しました。\\n\\n【詳細な理由】\\n" + response.getContentText(), Browser.Buttons.OK);
       return null;
    }

    var json = JSON.parse(response.getContentText());
    if (json.candidates && json.candidates.length > 0) {
      var rawText = json.candidates[0].content.parts[0].text;
      
      try {
        var bt = String.fromCharCode(96, 96, 96);
        var cleanText = rawText.replace(new RegExp("^" + bt + "(?:json)?\\n?", "i"), "").replace(new RegExp("\\n?" + bt + "$", "i"), "").trim();
        return JSON.parse(cleanText);
      } catch (parseError) {
        Browser.msgBox("🛑 AI回答形式エラー", "AIの回答形式エラー:\\n" + rawText, Browser.Buttons.OK);
        return null;
      }
    }
  } catch (e) {
    Browser.msgBox("🛑 予期せぬエラー", "callGeminiAPI内でエラー:\\n" + e.toString(), Browser.Buttons.OK);
  }
  return null;
}
