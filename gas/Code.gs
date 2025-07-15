/**
 * Google Apps Script版 1on1ドキュメント作成ツール
 * 
 * 使い方:
 * 1. このスクリプトをGoogle Apps Scriptにコピー
 * 2. createOneOnOneDoc() を実行
 * 3. 初回は認証が必要
 */

// グローバル設定オブジェクト
let CONFIG = null;

/**
 * 設定を読み込む
 */
function loadConfig() {
  if (CONFIG) return CONFIG;
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet) {
      // Configシートがない場合はデフォルト値を返す
      return getDefaultConfig();
    }
    
    // B列の値を読み込む（B2からB37まで）
    const values = sheet.getRange('B2:B37').getValues();
    
    CONFIG = {
      // 基本設定
      yourName: sheet.getRange('B2').getValue() || '',
      
      // ドキュメント設定
      documentTitleFormat: values[3][0] || '1on1 {yourName} / {partnerName}',
      partnerDefaultText: values[4][0] || '何かあれば',
      yourDefaultText: values[5][0] || 'あとで書きます！',
      
      // UIテキスト設定 - ダイアログ
      dialogTitle: values[8][0] || '1on1ドキュメント作成',
      namePrompt: values[9][0] || 'あなたの名前を入力してください:',
      emailSearchTitle: values[10][0] || 'メールアドレス検索',
      emailSearchPrompt: values[11][0] || '名前/メールアドレスの一部を入力（例: panko）、または完全なメールアドレスを入力してください:',
      userSelectTitle: values[12][0] || 'ユーザーを選択',
      
      // UIテキスト設定 - メッセージ
      successTitle: values[14][0] || '成功',
      successMessage: values[15][0] || 'ドキュメントが作成されました！\n\nURL: {url}\n\n{editorEmail} を編集者として招待しました。',
      errorTitle: values[16][0] || 'エラー',
      emailValidationError: values[17][0] || '有効なメールアドレスを入力してください。',
      documentCreationError: values[18][0] || 'ドキュメントの作成中にエラーが発生しました: ',
      noResultsMessage: values[19][0] || '該当するユーザーが見つかりませんでした。',
      invalidSelectionError: values[20][0] || '無効な番号が入力されました。',
      
      // メニュー設定
      menuName: values[22][0] || '1on1 Tool',
      menuItem1: values[23][0] || 'Create New Document',
      menuItem2: values[24][0] || 'Create with Interactive Search',
      
      // サイドバー設定
      sidebarTitle: values[27][0] || '1on1 Document Creator',
      nameLabel: values[28][0] || 'Your Name',
      emailLabel: values[29][0] || 'Invite Email Address',
      emailPlaceholder: values[30][0] || 'Type to search @quipper.com addresses',
      createButtonText: values[31][0] || 'Create Document',
      creatingText: values[32][0] || 'Creating...',
      searchingText: values[33][0] || 'Searching...',
      fieldsError: values[34][0] || 'Please fill in all fields',
      createSuccessMessage: values[35][0] || 'Document created successfully!'
    };
    
    return CONFIG;
  } catch (error) {
    console.error('設定の読み込みエラー:', error);
    return getDefaultConfig();
  }
}

/**
 * デフォルト設定を取得
 */
function getDefaultConfig() {
  return {
    yourName: '',
    documentTitleFormat: '1on1 {yourName} / {partnerName}',
    partnerDefaultText: '何かあれば',
    yourDefaultText: 'あとで書きます！',
    dialogTitle: '1on1ドキュメント作成',
    namePrompt: 'あなたの名前を入力してください:',
    emailSearchTitle: 'メールアドレス検索',
    emailSearchPrompt: '名前/メールアドレスの一部を入力（例: panko）、または完全なメールアドレスを入力してください:',
    userSelectTitle: 'ユーザーを選択',
    successTitle: '成功',
    successMessage: 'ドキュメントが作成されました！\n\nURL: {url}\n\n{editorEmail} を編集者として招待しました。',
    errorTitle: 'エラー',
    emailValidationError: '有効なメールアドレスを入力してください。',
    documentCreationError: 'ドキュメントの作成中にエラーが発生しました: ',
    noResultsMessage: '該当するユーザーが見つかりませんでした。',
    invalidSelectionError: '無効な番号が入力されました。',
    menuName: '1on1 Tool',
    menuItem1: 'Create New Document',
    menuItem2: 'Create with Interactive Search',
    sidebarTitle: '1on1 Document Creator',
    nameLabel: 'Your Name',
    emailLabel: 'Invite Email Address',
    emailPlaceholder: 'Type to search @quipper.com addresses',
    createButtonText: 'Create Document',
    creatingText: 'Creating...',
    searchingText: 'Searching...',
    fieldsError: 'Please fill in all fields',
    createSuccessMessage: 'Document created successfully!'
  };
}

/**
 * メイン関数 - 1on1ドキュメントを作成
 */
function createOneOnOneDoc() {
  // UIを取得
  const ui = SpreadsheetApp.getUi();
  
  // 設定を読み込む
  const config = loadConfig();
  
  // Configシートから名前を取得、なければ入力を求める
  let configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!configSheet) {
    // Configシートが存在しない場合は作成
    configSheet = createConfigSheet();
    CONFIG = null; // 設定をリセット
    config = loadConfig(); // 再読み込み
  }
  let yourName = config.yourName;
  
  if (!yourName || yourName.toString().trim() === '') {
    const yourNameResponse = ui.prompt(
      config.dialogTitle,
      config.namePrompt,
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
    ui.alert(config.errorTitle, config.emailValidationError, ui.ButtonSet.OK);
    return;
  }
  
  try {
    // ドキュメントの作成
    const result = createDocument(yourName, editorEmail);
    
    // 成功メッセージ
    const message = config.successMessage
      .replace('{url}', result.url)
      .replace('{editorEmail}', editorEmail);
    ui.alert(config.successTitle, message, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert(config.errorTitle, config.documentCreationError + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * ドキュメントを作成して共有
 */
function createDocument(yourName, editorEmail) {
  // 設定を読み込む
  const config = loadConfig();
  
  // パートナー名を取得（IDまたはメールアドレスから）
  const partnerName = getPartnerName(editorEmail);
  
  // ドキュメントのタイトル（設定から取得）
  const title = config.documentTitleFormat
    .replace('{yourName}', yourName)
    .replace('{partnerName}', partnerName);
  
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
  
  const subList1 = body.appendListItem(config.partnerDefaultText);
  subList1.setNestingLevel(1);
  subList1.setGlyphType(DocumentApp.GlyphType.BULLET);
  subList1.setListId(list1);
  
  const list2 = body.appendListItem(yourName);
  list2.setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const subList2 = body.appendListItem(config.yourDefaultText);
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
  const config = loadConfig();
  ui.createMenu(config.menuName)
    .addItem(config.menuItem1, 'createOneOnOneDoc')
    .addItem(config.menuItem2, 'showSidebar')
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
    
    // ===== 基本設定セクション =====
    sheet.getRange('A1').setValue('基本設定');
    sheet.getRange('A1').setFontWeight('bold').setFontSize(14).setBackground('#4285f4').setFontColor('#ffffff');
    sheet.getRange('A1:B1').merge();
    
    sheet.getRange('A2').setValue('Your Name:');
    sheet.getRange('B2').setValue(''); // ユーザーが名前を入力する場所
    
    // ===== ドキュメント設定セクション =====
    sheet.getRange('A4').setValue('ドキュメント設定');
    sheet.getRange('A4').setFontWeight('bold').setFontSize(14).setBackground('#4285f4').setFontColor('#ffffff');
    sheet.getRange('A4:B4').merge();
    
    sheet.getRange('A5').setValue('ドキュメントタイトル形式:');
    sheet.getRange('B5').setValue('1on1 {yourName} / {partnerName}');
    
    sheet.getRange('A6').setValue('相手のデフォルトテキスト:');
    sheet.getRange('B6').setValue('何かあれば');
    
    sheet.getRange('A7').setValue('自分のデフォルトテキスト:');
    sheet.getRange('B7').setValue('あとで書きます！');
    
    // ===== UIテキスト設定セクション =====
    sheet.getRange('A9').setValue('UIテキスト設定');
    sheet.getRange('A9').setFontWeight('bold').setFontSize(14).setBackground('#4285f4').setFontColor('#ffffff');
    sheet.getRange('A9:B9').merge();
    
    // ダイアログ関連
    sheet.getRange('A10').setValue('ドキュメント作成ダイアログタイトル:');
    sheet.getRange('B10').setValue('1on1ドキュメント作成');
    
    sheet.getRange('A11').setValue('名前入力プロンプト:');
    sheet.getRange('B11').setValue('あなたの名前を入力してください:');
    
    sheet.getRange('A12').setValue('メール検索ダイアログタイトル:');
    sheet.getRange('B12').setValue('メールアドレス検索');
    
    sheet.getRange('A13').setValue('メール検索プロンプト:');
    sheet.getRange('B13').setValue('名前/メールアドレスの一部を入力（例: panko）、または完全なメールアドレスを入力してください:');
    
    sheet.getRange('A14').setValue('ユーザー選択ダイアログタイトル:');
    sheet.getRange('B14').setValue('ユーザーを選択');
    
    // メッセージ関連
    sheet.getRange('A16').setValue('成功メッセージタイトル:');
    sheet.getRange('B16').setValue('成功');
    
    sheet.getRange('A17').setValue('成功メッセージ本文:');
    sheet.getRange('B17').setValue('ドキュメントが作成されました！\n\nURL: {url}\n\n{editorEmail} を編集者として招待しました。');
    
    sheet.getRange('A18').setValue('エラーメッセージタイトル:');
    sheet.getRange('B18').setValue('エラー');
    
    sheet.getRange('A19').setValue('メール検証エラー:');
    sheet.getRange('B19').setValue('有効なメールアドレスを入力してください。');
    
    sheet.getRange('A20').setValue('ドキュメント作成エラー:');
    sheet.getRange('B20').setValue('ドキュメントの作成中にエラーが発生しました: ');
    
    sheet.getRange('A21').setValue('検索結果なしメッセージ:');
    sheet.getRange('B21').setValue('該当するユーザーが見つかりませんでした。');
    
    sheet.getRange('A22').setValue('無効な選択エラー:');
    sheet.getRange('B22').setValue('無効な番号が入力されました。');
    
    // メニュー関連
    sheet.getRange('A24').setValue('メニュー名:');
    sheet.getRange('B24').setValue('1on1 Tool');
    
    sheet.getRange('A25').setValue('メニュー項目1:');
    sheet.getRange('B25').setValue('Create New Document');
    
    sheet.getRange('A26').setValue('メニュー項目2:');
    sheet.getRange('B26').setValue('Create with Interactive Search');
    
    // ===== サイドバー設定セクション =====
    sheet.getRange('A28').setValue('サイドバー設定');
    sheet.getRange('A28').setFontWeight('bold').setFontSize(14).setBackground('#4285f4').setFontColor('#ffffff');
    sheet.getRange('A28:B28').merge();
    
    sheet.getRange('A29').setValue('サイドバータイトル:');
    sheet.getRange('B29').setValue('1on1 Document Creator');
    
    sheet.getRange('A30').setValue('名前ラベル:');
    sheet.getRange('B30').setValue('Your Name');
    
    sheet.getRange('A31').setValue('メールアドレスラベル:');
    sheet.getRange('B31').setValue('Invite Email Address');
    
    sheet.getRange('A32').setValue('メール検索プレースホルダー:');
    sheet.getRange('B32').setValue('Type to search @quipper.com addresses');
    
    sheet.getRange('A33').setValue('作成ボタンテキスト:');
    sheet.getRange('B33').setValue('Create Document');
    
    sheet.getRange('A34').setValue('作成中テキスト:');
    sheet.getRange('B34').setValue('Creating...');
    
    sheet.getRange('A35').setValue('検索中テキスト:');
    sheet.getRange('B35').setValue('Searching...');
    
    sheet.getRange('A36').setValue('フィールド入力エラー:');
    sheet.getRange('B36').setValue('Please fill in all fields');
    
    sheet.getRange('A37').setValue('作成成功メッセージ:');
    sheet.getRange('B37').setValue('Document created successfully!');
    
    // 書式設定
    sheet.getRange('A2:A37').setFontWeight('bold');
    sheet.getRange('A2:A37').setBackground('#f3f3f3');
    
    // 列幅を調整
    sheet.setColumnWidth(1, 300);
    sheet.setColumnWidth(2, 400);
    
    // 使い方の説明を追加
    sheet.getRange('D1').setValue('使い方:');
    sheet.getRange('D1').setFontWeight('bold').setFontSize(12);
    sheet.getRange('D2').setValue('1. B2セルにあなたの名前を入力');
    sheet.getRange('D3').setValue('2. 各設定項目を必要に応じて変更');
    sheet.getRange('D4').setValue('3. {yourName}と{partnerName}は自動的に置換されます');
    sheet.getRange('D5').setValue('4. 変更後は画面をリロードしてください');
    
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
  const config = loadConfig();
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle(config.sidebarTitle)
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * サイドバーから呼ばれる：自分の名前を取得
 */
function getYourName() {
  const config = loadConfig();
  return config.yourName;
}

/**
 * サイドバーから呼ばれる：UI設定を取得
 */
function getUIConfig() {
  const config = loadConfig();
  return {
    nameLabel: config.nameLabel,
    emailLabel: config.emailLabel,
    emailPlaceholder: config.emailPlaceholder,
    createButtonText: config.createButtonText,
    creatingText: config.creatingText,
    searchingText: config.searchingText,
    fieldsError: config.fieldsError,
    createSuccessMessage: config.createSuccessMessage
  };
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
  const config = loadConfig();
  
  // 検索キーワードを入力（直接入力も可能）
  const searchResponse = ui.prompt(
    config.emailSearchTitle,
    config.emailSearchPrompt,
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
    ui.alert(config.errorTitle, config.noResultsMessage, ui.ButtonSet.OK);
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
    config.userSelectTitle,
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
    ui.alert(config.errorTitle, config.invalidSelectionError, ui.ButtonSet.OK);
    return null;
  }
}