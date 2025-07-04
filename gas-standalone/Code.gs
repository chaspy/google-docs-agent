/**
 * スタンドアロン版 Google Apps Script
 * 1on1ドキュメント作成ツール（Webアプリ版）
 */

/**
 * Webアプリのエントリーポイント
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('1on1 Document Creator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ドキュメントを作成（Webアプリから呼び出される）
 */
function createDocumentFromWeb(yourName, editorEmail) {
  // メールアドレスの検証
  if (!isValidEmail(editorEmail)) {
    throw new Error('有効なメールアドレスを入力してください。');
  }
  
  try {
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
      success: true,
      url: doc.getUrl(),
      id: docId,
      message: `ドキュメントを作成し、${editorEmail} を編集者として招待しました。`
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * メールアドレスの検証
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * HTMLファイルを読み込む（テンプレート用）
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}