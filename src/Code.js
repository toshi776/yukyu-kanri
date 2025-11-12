function doGet(e) {
  console.log('doGet called with parameters:', e.parameter);

  try {
    // ワンクリック承認処理
    var action = e.parameter.action || '';
    if (action === 'approve' || action === 'reject') {
      return handleOneClickApproval(e.parameter);
    }

    // 管理画面アクセス（管理者用）
    var adminParam = e.parameter.admin || '';
    if (adminParam.toLowerCase() === 'true') {
      return renderAdminPage();
    }

    // URLキーベースのアクセス（利用者用）
    var urlKey = e.parameter.key || '';
    if (urlKey) {
      return handleUserAccess(urlKey);
    }

    // 従来のユーザーIDアクセス（後方互換性）
    var userId = e.parameter.userID || e.parameter.userId || '';
    if (userId) {
      return handleLegacyUserAccess(userId);
    }

    // パラメータなしの場合はアクセス方法を案内
    return renderAccessGuide();

  } catch (error) {
    console.error('doGet エラー:', error);
    return renderErrorPage('システムエラーが発生しました: ' + error.message);
  }
}

/**
 * ワンクリック承認処理
 */
function handleOneClickApproval(params) {
  try {
    console.log('ワンクリック承認処理開始:', params);

    // トークンを検証
    var validationResult = validateApprovalToken(params);

    if (!validationResult || !validationResult.valid) {
      console.log('無効なトークン:', params);
      return renderApprovalResult(false, '無効なリンクです。リンクの有効期限が切れているか、既に処理済みの可能性があります。');
    }

    var action = params.action;
    var isApprove = action === 'approve';
    var rowNumber = validationResult.rowNumber;
    var application = validationResult.application;

    console.log('承認処理実行:', { action, rowNumber, application });

    // 承認または却下を実行
    var result;
    if (isApprove) {
      result = approveRecord({ rowNumber: rowNumber, status: 'Approved' });
    } else {
      result = rejectRecord({ rowNumber: rowNumber, reason: 'メールから却下' });
    }

    // デバッグ: resultオブジェクトの内容を確認
    console.log('承認処理結果:', JSON.stringify(result));

    // resultが正しいオブジェクトかチェック
    if (!result || typeof result !== 'object') {
      console.error('承認処理が不正な値を返しました:', result);
      return renderApprovalResult(false, 'システムエラーが発生しました。再度お試しください。');
    }

    if (result.success) {
      var message = isApprove ?
        `${application.userName}さんの${application.applyDate}の有給申請（${application.applyDays}日）を承認しました。` :
        `${application.userName}さんの${application.applyDate}の有給申請（${application.applyDays}日）を却下しました。`;

      console.log('承認処理成功:', message);
      return renderApprovalResult(true, message);
    } else {
      // エラーメッセージがundefinedの場合はデフォルトメッセージを使用
      var errorMessage = result.message || '不明なエラーが発生しました';
      console.error('承認処理失敗:', errorMessage);
      console.error('resultオブジェクト全体:', result);
      return renderApprovalResult(false, '処理中にエラーが発生しました: ' + errorMessage);
    }

  } catch (error) {
    console.error('ワンクリック承認処理エラー:', error);
    return renderApprovalResult(false, 'システムエラーが発生しました: ' + error.message);
  }
}

/**
 * 承認結果画面を表示
 */
function renderApprovalResult(success, message) {
  var color = success ? '#4CAF50' : '#f44336';
  var bgColor = success ? '#d4edda' : '#f8d7da';
  var borderColor = success ? '#c3e6cb' : '#f5c6cb';
  var icon = success ? '✓' : '✗';
  var title = success ? '処理完了' : 'エラー';

  var html = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>${title} - 有給管理システム</title>
    <style>
      body { font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }
      .container { max-width: 600px; margin: 50px auto; text-align: center; }
      .result {
        background: ${bgColor};
        border: 2px solid ${borderColor};
        padding: 30px;
        border-radius: 8px;
        color: ${color};
      }
      .icon {
        font-size: 64px;
        margin-bottom: 20px;
        color: ${color};
      }
      h1 { color: ${color}; margin-bottom: 20px; }
      .message {
        font-size: 16px;
        margin: 20px 0;
        color: #333;
      }
      .button {
        display: inline-block;
        margin-top: 30px;
        padding: 12px 24px;
        background: ${color};
        color: white;
        text-decoration: none;
        border-radius: 4px;
        font-weight: bold;
      }
      .button:hover {
        opacity: 0.9;
      }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="result">
        <div class="icon">${icon}</div>
        <h1>${title}</h1>
        <div class="message">${message}</div>
        <a href="?admin=true" class="button">管理画面へ</a>
      </div>
    </div>
  </body>
  </html>`;

  return HtmlService.createHtmlOutput(html)
    .setTitle(title + ' - 有給管理システム');
}

/**
 * 管理画面を表示
 */
function renderAdminPage() {
  try {
    var template = HtmlService.createTemplateFromFile('src/admin');
    return template.evaluate()
      .setTitle('有給管理システム - 管理画面')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    console.error('管理画面表示エラー:', error);
    return renderErrorPage('管理画面の表示に失敗しました');
  }
}

/**
 * URLキーベースの利用者アクセス処理
 */
function handleUserAccess(urlKey) {
  try {
    // URLキーから利用者番号を取得
    var userId = getUserIdFromUrlKey(urlKey);
    
    if (!userId) {
      console.log('無効なURLキーでのアクセス:', urlKey);
      return renderErrorPage('アクセスURLが無効です。正しいURLでアクセスしてください。');
    }
    
    // 利用者情報を取得
    var userInfo = getMasterRecord(userId);
    if (!userInfo) {
      console.log('利用者情報が見つからない:', userId);
      return renderErrorPage('利用者情報が見つかりません。管理者にお問い合わせください。');
    }
    
    console.log('正常なアクセス:', userId, userInfo.name);
    
    // 個人専用ページを表示
    return renderPersonalPage(userId, userInfo);
    
  } catch (error) {
    console.error('利用者アクセス処理エラー:', error);
    return renderErrorPage('アクセス処理中にエラーが発生しました');
  }
}

/**
 * 従来のユーザーIDアクセス（後方互換性）
 */
function handleLegacyUserAccess(userId) {
  try {
    console.log('従来方式でのアクセス:', userId);
    
    // 利用者情報を確認
    var userInfo = getMasterRecord(userId);
    if (!userInfo) {
      return renderErrorPage('利用者情報が見つかりません');
    }
    
    // 警告付きで申請フォームを表示
    var html = HtmlService.createHtmlOutputFromFile('src/form')
      .setTitle('有給申請フォーム（従来方式）')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    
    var content = html.getContent();
    
    // ユーザー情報とセキュリティ警告を埋め込み
    var userScript = `
    <script>
      var userIdFromGAS = "${userId}";
      var userNameFromGAS = "${userInfo.name}";
      var remainingDaysFromGAS = ${userInfo.remaining};
      var securityWarning = "このアクセス方式は将来廃止予定です。個人専用URLをご利用ください。";
      
      // 警告メッセージを表示
      window.addEventListener('load', function() {
        if (securityWarning) {
          var warningDiv = document.createElement('div');
          warningDiv.style.cssText = 'background:#fff3cd;border:1px solid #ffeaa7;padding:10px;margin:10px 0;border-radius:4px;color:#856404;';
          warningDiv.innerHTML = '<strong>お知らせ:</strong> ' + securityWarning;
          document.body.insertBefore(warningDiv, document.body.firstChild);
        }
      });
    </script>`;
    
    content = content.replace('<script>', userScript);
    html.setContent(content);
    
    return html;
    
  } catch (error) {
    console.error('従来アクセス処理エラー:', error);
    return renderErrorPage('アクセス処理中にエラーが発生しました');
  }
}

/**
 * 個人専用ページを表示
 */
function renderPersonalPage(userId, userInfo) {
  try {
    console.log('renderPersonalPage開始:', userId, userInfo);

    // 申請履歴を取得
    var applications = getUserApplications(userId);
    console.log('申請履歴取得完了:', applications);

    // 個人ページテンプレートを作成
    console.log('テンプレートファイル読み込み開始: src/personal');
    var template = HtmlService.createTemplateFromFile('src/personal');
    console.log('テンプレートファイル読み込み完了');

    // テンプレートに変数を設定
    template.userId = userId;
    template.userName = userInfo.name;
    template.remainingDays = userInfo.remaining;
    template.applications = applications || [];
    template.currentDate = Utilities.formatDate(new Date(), 'JST', 'yyyy年M月d日');
    console.log('テンプレート変数設定完了');

    console.log('テンプレート評価開始');
    var evaluated = template.evaluate();
    console.log('テンプレート評価完了');

    return evaluated
      .setTitle('有給管理システム - ' + userInfo.name + 'さんの専用ページ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  } catch (error) {
    console.error('個人ページ表示エラー:', error);
    console.error('エラースタック:', error.stack);
    return renderErrorPage('個人ページの表示に失敗しました: ' + error.message);
  }
}

/**
 * アクセス方法案内ページを表示
 */
function renderAccessGuide() {
  var html = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>有給管理システム - アクセス案内</title>
    <style>
      body { font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }
      .container { max-width: 600px; margin: 0 auto; }
      .header { background: #4CAF50; color: white; padding: 20px; text-align: center; border-radius: 8px; }
      .content { background: #f9f9f9; padding: 20px; margin: 20px 0; border-radius: 8px; }
      .warning { background: #fff3cd; border: 1px solid #ffeaa7; padding: 15px; border-radius: 4px; color: #856404; }
      .info { background: #d1ecf1; border: 1px solid #bee5eb; padding: 15px; border-radius: 4px; color: #0c5460; }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="header">
        <h1>有給管理システム</h1>
        <p>年次有給休暇申請・管理システム</p>
      </div>
      
      <div class="content">
        <h2>アクセス方法について</h2>
        
        <div class="info">
          <h3>🔐 個人専用URLでアクセス</h3>
          <p>セキュリティ向上のため、各利用者には専用のアクセスURLが発行されています。</p>
          <p><strong>例:</strong> https://script.google.com/.../exec?key=abc123...</p>
          <p>専用URLは人事部門または管理者にお問い合わせください。</p>
        </div>
        
        <div class="warning">
          <h3>⚠️ 管理者の方へ</h3>
          <p>管理画面にアクセスするには以下のURLをご利用ください：</p>
          <p><strong>管理画面:</strong> <a href="?admin=true">こちらをクリック</a></p>
        </div>
        
        <h3>📞 お問い合わせ</h3>
        <p>アクセスURLが不明な場合や、システムに関するご質問は人事部門までお問い合わせください。</p>
      </div>
    </div>
  </body>
  </html>`;
  
  return HtmlService.createHtmlOutput(html)
    .setTitle('有給管理システム - アクセス案内');
}

/**
 * エラーページを表示
 */
function renderErrorPage(message) {
  var html = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>エラー - 有給管理システム</title>
    <style>
      body { font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }
      .container { max-width: 600px; margin: 0 auto; text-align: center; }
      .error { background: #f8d7da; border: 1px solid #f5c6cb; padding: 20px; border-radius: 8px; color: #721c24; }
      .icon { font-size: 48px; margin-bottom: 20px; }
      h1 { color: #721c24; }
      .contact { margin-top: 30px; padding: 15px; background: #f9f9f9; border-radius: 4px; }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="error">
        <div class="icon">⚠️</div>
        <h1>アクセスエラー</h1>
        <p>${message}</p>
      </div>
      
      <div class="contact">
        <h3>お困りの場合は</h3>
        <p>正しいアクセスURLについては人事部門または管理者にお問い合わせください。</p>
        <p><a href="?admin=true">管理者の方はこちら</a></p>
      </div>
    </div>
  </body>
  </html>`;
  
  return HtmlService.createHtmlOutput(html)
    .setTitle('エラー - 有給管理システム');
}

/**
 * HTMLテンプレートの読み込み
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    console.error('HTMLファイル読み込みエラー:', e);
    return '';
  }
}

/**
 * フォーム送信処理
 */
function doPost(e) {
  try {
    appendApplication(e.parameter);
    var successHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>申請完了</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 20px; text-align: center; }
          .success { background: #d4edda; border: 1px solid #c3e6cb; padding: 20px; border-radius: 8px; color: #155724; max-width: 500px; margin: 50px auto; }
          .icon { font-size: 48px; margin-bottom: 20px; }
          h1 { color: #155724; }
          .button { display: inline-block; margin-top: 20px; padding: 10px 20px; background: #28a745; color: white; text-decoration: none; border-radius: 5px; }
        </style>
      </head>
      <body>
        <div class="success">
          <div class="icon">✅</div>
          <h1>申請完了</h1>
          <p>有給休暇の申請が正常に送信されました。</p>
          <p>承認結果をお待ちください。</p>
          <a href="javascript:window.close();" class="button">閉じる</a>
        </div>
      </body>
      </html>`;
    return HtmlService.createHtmlOutput(successHtml);
  } catch (error) {
    console.error('doPost エラー:', error);
    return HtmlService.createHtmlOutput('<h1>送信エラー</h1><p>' + error.toString() + '</p>');
  }
}

/**
 * 管理画面接続テスト用
 */
function testAuthorization() {
  try {
    getSpreadsheet();
    console.log('認証テスト成功');
    return true;
  } catch (e) {
    console.error('認証テストエラー:', e);
    return false;
  }
}

// テスト用関数：申請履歴の確認
function testUserApplications() {
  var userId = '123456'; // テストする利用者番号
  
  try {
    console.log('=== 申請履歴テスト開始 ===');
    
    // スプレッドシートの確認
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // シート一覧を確認
    var sheets = ss.getSheets();
    console.log('利用可能なシート:', sheets.map(function(s) { return s.getName(); }));
    
    // 申請シートの存在確認
    var applySheet = ss.getSheetByName('申請');
    if (!applySheet) {
      console.error('「申請」シートが見つかりません');
      // 他の名前の可能性をチェック
      var possibleNames = ['Applications', '申請一覧', '申請履歴'];
      possibleNames.forEach(function(name) {
        var sheet = ss.getSheetByName(name);
        if (sheet) {
          console.log('代替シート発見:', name);
        }
      });
      return;
    }
    
    console.log('申請シート発見');
    
    // データの確認
    var data = applySheet.getDataRange().getValues();
    console.log('データ行数:', data.length);
    
    if (data.length > 0) {
      console.log('ヘッダー行:', data[0]);
      
      // 最初の数行を表示
      for (var i = 1; i < Math.min(data.length, 4); i++) {
        console.log('行' + i + ':', data[i]);
      }
    }
    
    // getUserApplications関数のテスト
    console.log('\n=== getUserApplications関数実行 ===');
    var result = getUserApplications(userId);
    console.log('取得結果:', result);
    console.log('件数:', result ? result.length : 0);
    
  } catch (error) {
    console.error('テストエラー:', error);
    console.error('スタック:', error.stack);
  }
}

/**
 * 古いデータを削除する管理用関数
 * Google Apps Scriptエディタで実行してください
 */
function cleanupOldData() {
  var deletedCount = deleteOldApplications();
  console.log('削除完了:', deletedCount + '行の古いデータを削除しました');
  return '削除完了: ' + deletedCount + '行の古いデータを削除しました';
}

/**
 * 有給申請を送信（個人ページから呼び出される）
 */
function submitRequest(params) {
  try {
    console.log('submitRequest called with params:', params);

    // 利用者情報を取得
    var userInfo = getMasterRecord(params.userId);
    if (!userInfo) {
      return {
        success: false,
        message: '利用者情報が見つかりません'
      };
    }

    // 残日数チェック
    var applyDays = parseFloat(params.applyDays) || 1;
    if (userInfo.remaining < applyDays) {
      return {
        success: false,
        message: '有給残日数が不足しています（残り' + userInfo.remaining + '日）'
      };
    }

    // 申請データを追加
    var ss = getSpreadsheet();
    var appSheet = ss.getSheetByName('申請');

    if (!appSheet) {
      return {
        success: false,
        message: '申請シートが見つかりません'
      };
    }

    // 新しい残日数を計算
    var newRemaining = userInfo.remaining - applyDays;

    // 申請シートに追加
    appSheet.appendRow([
      params.userId,
      userInfo.name,
      newRemaining,
      params.applyDate,
      new Date(),
      'Pending',
      '',
      applyDays
    ]);

    // マスターシートの残日数を更新
    var masterSheet = ss.getSheetByName('マスター');
    if (masterSheet) {
      var finder = masterSheet.createTextFinder(params.userId).matchEntireCell(true);
      var foundCell = finder.findNext();
      if (foundCell) {
        var rowIndex = foundCell.getRow();
        var remCell = masterSheet.getRange(rowIndex, 3);
        remCell.setValue(newRemaining);
      }
    }

    console.log('申請成功:', params.userId, params.applyDate, applyDays + '日');

    // 承認者に通知を送信
    try {
      var application = {
        userId: params.userId,
        userName: userInfo.name,
        applyDate: params.applyDate,
        applyDays: applyDays,
        timestamp: new Date()
      };
      sendApplicationNotification(application);
      console.log('承認者への通知送信完了');
    } catch (notificationError) {
      console.error('通知送信エラー（申請は成功）:', notificationError);
      // 通知エラーでも申請は成功とする
    }

    return {
      success: true,
      message: '申請が完了しました',
      newRemaining: newRemaining
    };

  } catch (error) {
    console.error('submitRequest error:', error);
    return {
      success: false,
      message: 'エラーが発生しました: ' + error.message
    };
  }
}