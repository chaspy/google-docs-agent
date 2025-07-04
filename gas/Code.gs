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
  
  // あなたの名前を入力
  const yourNameResponse = ui.prompt(
    '1on1ドキュメント作成',
    'あなたの名前を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (yourNameResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const yourName = yourNameResponse.getResponseText();
  
  // 相手のメールアドレスを入力
  const emailResponse = ui.prompt(
    '1on1ドキュメント作成',
    '招待する相手のメールアドレスを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (emailResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const editorEmail = emailResponse.getResponseText();
  
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
  // 日付とパートナー名を取得
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const partnerName = editorEmail.split('@')[0];
  
  // ドキュメントのタイトル
  const title = `1on1 - ${date} - ${partnerName}`;
  
  // ドキュメントを作成
  const doc = DocumentApp.create(title);
  const body = doc.getBody();
  
  // 初期コンテンツを設定
  body.appendParagraph(`${date} * ${yourName} * ${partnerName}`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('');
  body.appendParagraph('話したいこと').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendListItem('');
  body.appendParagraph('');
  body.appendParagraph('メモ').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('');
  
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
  ui.createMenu('1on1ツール')
    .addItem('新規ドキュメント作成', 'createOneOnOneDoc')
    .addToUi();
}