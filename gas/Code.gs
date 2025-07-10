/**
 * Google Apps Script版 1on1ドキュメント作成ツール
 * 
 * 使い方:
 * 1. このスクリプトをGoogle Apps Scriptにコピー
 * 2. createOneOnOneDoc() を実行
 * 3. 初回は認証が必要
 */

/**
 * メイン関数 - 1on1ドキュメントを作成
 */
function createOneOnOneDoc() {
  // UIを取得
  const ui = SpreadsheetApp.getUi();
  
  // Configシートから名前を取得、なければ入力を求める
  let configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!configSheet) {
    // Configシートが存在しない場合は作成
    configSheet = createConfigSheet();
  }
  let yourName = configSheet.getRange('B1').getValue();
  
  if (!yourName || yourName.toString().trim() === '') {
    const yourNameResponse = ui.prompt(
      '1on1ドキュメント作成',
      'あなたの名前を入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (yourNameResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    yourName = yourNameResponse.getResponseText();
  }
  
  // メールアドレスを検索（デフォルト）
  const editorEmail = selectEmailWithSearch();
  if (!editorEmail) {
    return;
  }
  
  // メールアドレスの検証
  if (!isValidEmail(editorEmail)) {
    ui.alert('エラー', '有効なメールアドレスを入力してください。', ui.ButtonSet.OK);
    return;
  }
  
  try {
    // ドキュメントの作成
    const result = createDocument(yourName, editorEmail);
    
    // 成功メッセージ
    const message = `ドキュメントが作成されました！\n\nURL: ${result.url}\n\n${editorEmail} を編集者として招待しました。`;
    ui.alert('成功', message, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('エラー', 'ドキュメントの作成中にエラーが発生しました: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * ドキュメントを作成して共有
 */
function createDocument(yourName, editorEmail) {
  // パートナー名を取得（IDまたはメールアドレスから）
  const partnerName = getPartnerName(editorEmail);
  
  // ドキュメントのタイトル
  const title = `1on1 ${yourName} / ${partnerName}`;
  
  // ドキュメントを作成
  const doc = DocumentApp.create(title);
  const body = doc.getBody();
  
  // 初期コンテンツを設定
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  
  // タイトル（日付）
  body.appendParagraph(date).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('');
  
  // 各人のセクション（リストアイテムとして作成）
  const list1 = body.appendListItem(partnerName);
  list1.setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const subList1 = body.appendListItem('何かあれば');
  subList1.setNestingLevel(1);
  subList1.setGlyphType(DocumentApp.GlyphType.BULLET);
  subList1.setListId(list1);
  
  const list2 = body.appendListItem(yourName);
  list2.setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const subList2 = body.appendListItem('あとで書きます！');
  subList2.setNestingLevel(1);
  subList2.setGlyphType(DocumentApp.GlyphType.BULLET);
  subList2.setListId(list2);
  
  // ドキュメントを保存
  doc.saveAndClose();
  
  // ドキュメントのIDを取得
  const docId = doc.getId();
  const docFile = DriveApp.getFileById(docId);
  
  // 編集者として招待
  docFile.addEditor(editorEmail);
  
  return {
    id: docId,
    url: doc.getUrl()
  };
}

/**
 * メールアドレスの検証
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * メニューを追加（オプション）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('1on1 Tool')
    .addItem('Create New Document', 'createOneOnOneDoc')
    .addItem('Create with Interactive Search', 'showSidebar')
    .addToUi();
}

/**
 * メールアドレスリストを取得
 */
function searchUsers(query) {
  try {
    const users = [];
    const uniqueEmails = new Set(); // メールアドレスの重複チェック用
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 「member」シートから読み込み
    let memberSheet;
    try {
      memberSheet = sheet.getSheetByName('member');
    } catch (e) {
      // シートが存在しない場合は作成
      memberSheet = createMemberListSheet();
    }
    
    if (!memberSheet) {
      return [];
    }
    
    // データ範囲を取得（A列、J列、M列）
    const lastRow = memberSheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }
    
    // A列（名前）、J列（Googleメールアドレス）、M列（ID）を取得
    const nameRange = memberSheet.getRange(2, 1, lastRow - 1, 1);  // A列
    const emailRange = memberSheet.getRange(2, 10, lastRow - 1, 1); // J列
    const idRange = memberSheet.getRange(2, 13, lastRow - 1, 1);   // M列
    const names = nameRange.getValues();
    const emails = emailRange.getValues();
    const ids = idRange.getValues();
    
    // 検索実行
    for (let i = 0; i < names.length; i++) {
      const name = names[i][0];
      const email = emails[i][0];
      const id = ids[i][0];
      
      if (name && email && email.endsWith('@quipper.com')) {
        // 既に同じメールアドレスが追加されている場合はスキップ
        if (uniqueEmails.has(email)) {
          continue;
        }
        
        // クエリでフィルタリング（名前、メールアドレス、IDのいずれかで検索）
        if (!query || 
            name.toLowerCase().includes(query.toLowerCase()) || 
            email.toLowerCase().includes(query.toLowerCase()) ||
            (id && id.toLowerCase().includes(query.toLowerCase()))) {
          users.push({
            name: name + (id ? ` (${id})` : ''), // IDがあれば名前の後ろに表示
            email: email
          });
          uniqueEmails.add(email); // 追加したメールアドレスを記録
        }
      }
    }
    
    // 最大20件に制限
    return users.slice(0, 20);
    
  } catch (error) {
    console.error('ユーザー検索エラー:', error);
    return [];
  }
}

/**
 * Configシートを作成
 */
function createConfigSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.insertSheet('Config');
    
    // ヘッダーを設定
    sheet.getRange('A1').setValue('Your Name:');
    sheet.getRange('B1').setValue(''); // ユーザーが名前を入力する場所
    
    // ヘッダーの書式設定
    sheet.getRange('A1').setFontWeight('bold');
    sheet.getRange('A1').setBackground('#f3f3f3');
    
    // 列幅を調整
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 250);
    
    // 使い方の説明を追加
    sheet.getRange('A3').setValue('Instructions:');
    sheet.getRange('A4').setValue('Enter your name in cell B1');
    sheet.getRange('A5').setValue('This name will be used when creating 1on1 documents');
    
    return sheet;
  } catch (error) {
    console.error('Config sheet creation error:', error);
    return null;
  }
}

/**
 * メンバーリストシートを作成
 */
function createMemberListSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.insertSheet('member');
    
    // ヘッダーを設定
    sheet.getRange('A1').setValue('Name');
    sheet.getRange('B1').setValue('Email Address');
    
    // ヘッダーの書式設定
    const headerRange = sheet.getRange('A1:B1');
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');
    
    // サンプルデータを追加（A列に名前、J列にメールアドレス）
    const sampleNames = [
      ['Taro Yamada'],
      ['Hanako Suzuki'],
      ['Ichiro Tanaka']
    ];
    
    const sampleEmails = [
      ['yamada@example.com'],
      ['suzuki@example.com'],
      ['tanaka@example.com']
    ];
    
    sheet.getRange(2, 1, sampleNames.length, 1).setValues(sampleNames);  // A列
    sheet.getRange(2, 10, sampleEmails.length, 1).setValues(sampleEmails); // J列
    
    // 列幅を調整
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 250);
    
    // 使い方の説明を追加（J列の説明を追加）
    sheet.getRange('D1').setValue('How to use:');
    sheet.getRange('D2').setValue('1. Enter names in column A');
    sheet.getRange('D3').setValue('2. Enter Google email addresses in column J');
    sheet.getRange('D4').setValue('3. Run "1on1 Tool" → "Create New Document" from the menu');
    sheet.getRange('D5').setValue('4. Names/emails registered here will be searchable');
    
    // J列のヘッダーを追加
    sheet.getRange('J1').setValue('Google Email').setFontWeight('bold').setBackground('#f3f3f3');
    
    return sheet;
  } catch (error) {
    console.error('Members sheet creation error:', error);
    return null;
  }
}

/**
 * サイドバーを表示
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('1on1 Document Creator')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * サイドバーから呼ばれる：自分の名前を取得
 */
function getYourName() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet;
  
  try {
    configSheet = sheet.getSheetByName('Config');
  } catch (e) {
    configSheet = createConfigSheet();
  }
  
  if (!configSheet) {
    return '';
  }
  
  return configSheet.getRange('B1').getValue();
}

/**
 * サイドバーから呼ばれる：ドキュメント作成
 */
function createDocumentFromSidebar(yourName, editorEmail) {
  // メールアドレスの検証
  if (!isValidEmail(editorEmail)) {
    throw new Error('Invalid email address');
  }
  
  // ドキュメントの作成
  const result = createDocument(yourName, editorEmail);
  return result;
}

/**
 * パートナー名を取得（memberシートから検索）
 */
function getPartnerName(editorEmail) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const memberSheet = sheet.getSheetByName('member');
    
    if (!memberSheet) {
      // memberシートがない場合はメールアドレスから推測
      return editorEmail.split('@')[0];
    }
    
    const lastRow = memberSheet.getLastRow();
    if (lastRow < 2) {
      return editorEmail.split('@')[0];
    }
    
    // J列（メールアドレス）とM列（ID）を取得
    const emailRange = memberSheet.getRange(2, 10, lastRow - 1, 1);
    const idRange = memberSheet.getRange(2, 13, lastRow - 1, 1);
    const emails = emailRange.getValues();
    const ids = idRange.getValues();
    
    // メールアドレスに一致する行を探す
    for (let i = 0; i < emails.length; i++) {
      if (emails[i][0] === editorEmail) {
        // IDがあればIDを返す、なければメールアドレスから推測
        return ids[i][0] || editorEmail.split('@')[0];
      }
    }
    
    // 見つからない場合はメールアドレスから推測
    return editorEmail.split('@')[0];
    
  } catch (error) {
    console.error('パートナー名取得エラー:', error);
    return editorEmail.split('@')[0];
  }
}

/**
 * メールアドレスを検索して選択
 */
function selectEmailWithSearch() {
  const ui = SpreadsheetApp.getUi();
  
  // 検索キーワードを入力（直接入力も可能）
  const searchResponse = ui.prompt(
    'メールアドレス検索',
    '名前/メールアドレスの一部を入力（例: panko）、または完全なメールアドレスを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (searchResponse.getSelectedButton() !== ui.Button.OK) {
    return null;
  }
  
  const searchKeyword = searchResponse.getResponseText();
  
  // 完全なメールアドレスが入力された場合は、そのまま返す
  if (isValidEmail(searchKeyword)) {
    return searchKeyword;
  }
  
  // ユーザーを検索
  const users = searchUsers(searchKeyword);
  
  if (users.length === 0) {
    ui.alert('検索結果', '該当するユーザーが見つかりませんでした。', ui.ButtonSet.OK);
    return null;
  }
  
  if (users.length === 1) {
    // 1件のみの場合は自動選択
    return users[0].email;
  }
  
  // 複数件の場合は選択肢を表示
  let message = '該当するユーザー:\n\n';
  users.forEach((user, index) => {
    message += `${index + 1}. ${user.name} (${user.email})\n`;
  });
  message += '\n番号を入力してください:';
  
  const selectionResponse = ui.prompt(
    'ユーザーを選択',
    message,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (selectionResponse.getSelectedButton() !== ui.Button.OK) {
    return null;
  }
  
  const selectedIndex = parseInt(selectionResponse.getResponseText()) - 1;
  
  if (selectedIndex >= 0 && selectedIndex < users.length) {
    return users[selectedIndex].email;
  } else {
    ui.alert('エラー', '無効な番号が入力されました。', ui.ButtonSet.OK);
    return null;
  }
}