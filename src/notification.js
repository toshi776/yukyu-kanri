// =============================
// 通知機能システム
// =============================

/**
 * メール通知の設定
 */
var NOTIFICATION_CONFIG = {
  FROM_EMAIL: 'noreply@company.com', // 実際の環境では適切なメールアドレスに変更
  SUBJECT_PREFIX: '[有給管理システム] ',
  MAX_DAILY_EMAILS: 50, // GASの制限を考慮
  TEST_MODE: false, // 開発時はtrue、本番時はfalse
  ENABLE_NOTIFICATIONS: true, // 通知機能の有効/無効切り替え
  NOTIFICATION_SETTINGS: {
    APPLICATION: true,    // 申請通知
    APPROVAL: true,       // 承認結果通知
    ALERT: true,         // 残日数アラート
    GRANT: true,         // 付与通知
    EXPIRY: true         // 失効通知
  }
};

/**
 * 申請通知メールを送信（承認者向け）
 * @param {Object} application - 申請情報
 * @return {Object} 送信結果
 */
function sendApplicationNotification(application) {
  try {
    console.log('申請通知メール送信開始:', application);
    
    // 通知機能が無効な場合
    if (!NOTIFICATION_CONFIG.ENABLE_NOTIFICATIONS || !NOTIFICATION_CONFIG.NOTIFICATION_SETTINGS.APPLICATION) {
      return {
        success: true,
        message: '申請通知が無効になっています',
        disabled: true
      };
    }
    
    if (NOTIFICATION_CONFIG.TEST_MODE) {
      console.log('TEST MODE: メール送信をスキップします');
      return {
        success: true,
        message: 'テストモード: メール送信をスキップしました',
        testMode: true
      };
    }
    
    // 承認者のメールアドレスを取得（実際の環境では適切なロジックに変更）
    var approverEmail = getApproverEmail(application.userId);
    if (!approverEmail) {
      throw new Error('承認者のメールアドレスが見つかりません');
    }
    
    var subject = NOTIFICATION_CONFIG.SUBJECT_PREFIX + '有給申請の承認依頼';
    var htmlBody = generateApplicationNotificationHtml(application);
    var plainBody = generateApplicationNotificationText(application);
    
    // メール送信
    GmailApp.sendEmail(
      approverEmail,
      subject,
      plainBody,
      {
        htmlBody: htmlBody,
        name: '有給管理システム'
      }
    );
    
    console.log('申請通知メール送信完了:', approverEmail);
    
    return {
      success: true,
      message: '申請通知メールを送信しました',
      recipient: approverEmail
    };
    
  } catch (error) {
    console.error('申請通知メール送信エラー:', error);
    return {
      success: false,
      message: 'メール送信に失敗しました: ' + error.message
    };
  }
}

/**
 * 承認結果通知メールを送信（申請者向け）
 * @param {Object} application - 申請情報
 * @param {string} status - 承認結果 ('Approved' or 'Rejected')
 * @return {Object} 送信結果
 */
function sendApprovalResultNotification(application, status) {
  try {
    console.log('承認結果通知メール送信開始:', { application, status });
    
    // 通知機能が無効な場合
    if (!NOTIFICATION_CONFIG.ENABLE_NOTIFICATIONS || !NOTIFICATION_CONFIG.NOTIFICATION_SETTINGS.APPROVAL) {
      return {
        success: true,
        message: '承認結果通知が無効になっています',
        disabled: true
      };
    }
    
    if (NOTIFICATION_CONFIG.TEST_MODE) {
      console.log('TEST MODE: メール送信をスキップします');
      return {
        success: true,
        message: 'テストモード: メール送信をスキップしました',
        testMode: true
      };
    }
    
    // 申請者のメールアドレスを取得
    var applicantEmail = getEmployeeEmail(application.userId);
    if (!applicantEmail) {
      throw new Error('申請者のメールアドレスが見つかりません');
    }
    
    var isApproved = status === 'Approved';
    var subject = NOTIFICATION_CONFIG.SUBJECT_PREFIX + 
      (isApproved ? '有給申請が承認されました' : '有給申請について');
    
    var htmlBody = generateApprovalResultNotificationHtml(application, status);
    var plainBody = generateApprovalResultNotificationText(application, status);
    
    // メール送信
    GmailApp.sendEmail(
      applicantEmail,
      subject,
      plainBody,
      {
        htmlBody: htmlBody,
        name: '有給管理システム'
      }
    );
    
    console.log('承認結果通知メール送信完了:', applicantEmail);
    
    return {
      success: true,
      message: '承認結果通知メールを送信しました',
      recipient: applicantEmail
    };
    
  } catch (error) {
    console.error('承認結果通知メール送信エラー:', error);
    return {
      success: false,
      message: 'メール送信に失敗しました: ' + error.message
    };
  }
}

/**
 * 残日数アラートメールを送信
 * @param {Object} employee - 社員情報
 * @return {Object} 送信結果
 */
function sendLowRemainingDaysAlert(employee) {
  try {
    console.log('残日数アラートメール送信開始:', employee);
    
    // 通知機能が無効な場合
    if (!NOTIFICATION_CONFIG.ENABLE_NOTIFICATIONS || !NOTIFICATION_CONFIG.NOTIFICATION_SETTINGS.ALERT) {
      return {
        success: true,
        message: '残日数アラートが無効になっています',
        disabled: true
      };
    }
    
    if (NOTIFICATION_CONFIG.TEST_MODE) {
      console.log('TEST MODE: メール送信をスキップします');
      return {
        success: true,
        message: 'テストモード: メール送信をスキップしました',
        testMode: true
      };
    }
    
    var employeeEmail = getEmployeeEmail(employee.userId);
    if (!employeeEmail) {
      throw new Error('社員のメールアドレスが見つかりません');
    }
    
    var subject = NOTIFICATION_CONFIG.SUBJECT_PREFIX + '有給残日数のお知らせ';
    var htmlBody = generateLowRemainingDaysAlertHtml(employee);
    var plainBody = generateLowRemainingDaysAlertText(employee);
    
    // メール送信
    GmailApp.sendEmail(
      employeeEmail,
      subject,
      plainBody,
      {
        htmlBody: htmlBody,
        name: '有給管理システム'
      }
    );
    
    console.log('残日数アラートメール送信完了:', employeeEmail);
    
    return {
      success: true,
      message: '残日数アラートメールを送信しました',
      recipient: employeeEmail
    };
    
  } catch (error) {
    console.error('残日数アラートメール送信エラー:', error);
    return {
      success: false,
      message: 'メール送信に失敗しました: ' + error.message
    };
  }
}

// =============================
// メールテンプレート生成関数
// =============================

/**
 * 申請通知のHTMLメールテンプレート生成
 */
function generateApplicationNotificationHtml(application) {
  var daysText = application.applyDays === 0.5 ? '0.5日（半日）' : '1日';
  
  var html = `
  <html>
  <body style="font-family: Arial, sans-serif; color: #333;">
    <h2 style="color: #4CAF50;">有給申請の承認依頼</h2>
    
    <p>以下の有給申請について承認をお願いします。</p>
    
    <table style="border-collapse: collapse; margin: 20px 0;">
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f5;">申請者</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${application.userName}（${application.userId}）</td>
      </tr>
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f5;">申請日</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${application.applyDate}</td>
      </tr>
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f5;">申請日数</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${daysText}</td>
      </tr>
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f5;">申請時刻</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${application.timestamp}</td>
      </tr>
    </table>
    
    <p style="margin-top: 30px;">
      <strong>操作:</strong><br>
      管理画面にアクセスして承認・却下の処理をお願いします。
    </p>
    
    <p style="color: #666; font-size: 12px; margin-top: 40px;">
      このメールは有給管理システムから自動送信されています。<br>
      返信はできませんので、ご質問がある場合は人事部までお問い合わせください。
    </p>
  </body>
  </html>`;
  
  return html;
}

/**
 * 申請通知のテキストメール生成
 */
function generateApplicationNotificationText(application) {
  var daysText = application.applyDays === 0.5 ? '0.5日（半日）' : '1日';
  
  return `有給申請の承認依頼

以下の有給申請について承認をお願いします。

申請者: ${application.userName}（${application.userId}）
申請日: ${application.applyDate}
申請日数: ${daysText}
申請時刻: ${application.timestamp}

操作:
管理画面にアクセスして承認・却下の処理をお願いします。

※このメールは有給管理システムから自動送信されています。
返信はできませんので、ご質問がある場合は人事部までお問い合わせください。`;
}

/**
 * 承認結果通知のHTMLメール生成
 */
function generateApprovalResultNotificationHtml(application, status) {
  var isApproved = status === 'Approved';
  var statusText = isApproved ? '承認' : '却下';
  var color = isApproved ? '#4CAF50' : '#f44336';
  var daysText = application.applyDays === 0.5 ? '0.5日（半日）' : '1日';
  
  var html = `
  <html>
  <body style="font-family: Arial, sans-serif; color: #333;">
    <h2 style="color: ${color};">有給申請が${statusText}されました</h2>
    
    <p>以下の有給申請が${statusText}されました。</p>
    
    <table style="border-collapse: collapse; margin: 20px 0;">
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f5;">申請日</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${application.applyDate}</td>
      </tr>
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f5;">申請日数</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${daysText}</td>
      </tr>
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f5;">結果</td>
        <td style="padding: 8px; border: 1px solid #ddd; color: ${color}; font-weight: bold;">${statusText}</td>
      </tr>
    </table>
    
    ${isApproved ? 
      '<p style="color: #4CAF50;"><strong>有給申請が承認されました。予定通り休暇を取得してください。</strong></p>' :
      '<p style="color: #f44336;"><strong>申請内容について確認が必要です。詳細は上司または人事部にお問い合わせください。</strong></p>'
    }
    
    <p style="color: #666; font-size: 12px; margin-top: 40px;">
      このメールは有給管理システムから自動送信されています。<br>
      返信はできませんので、ご質問がある場合は人事部までお問い合わせください。
    </p>
  </body>
  </html>`;
  
  return html;
}

/**
 * 承認結果通知のテキストメール生成
 */
function generateApprovalResultNotificationText(application, status) {
  var isApproved = status === 'Approved';
  var statusText = isApproved ? '承認' : '却下';
  var daysText = application.applyDays === 0.5 ? '0.5日（半日）' : '1日';
  
  return `有給申請が${statusText}されました

以下の有給申請が${statusText}されました。

申請日: ${application.applyDate}
申請日数: ${daysText}
結果: ${statusText}

${isApproved ? 
  '有給申請が承認されました。予定通り休暇を取得してください。' :
  '申請内容について確認が必要です。詳細は上司または人事部にお問い合わせください。'
}

※このメールは有給管理システムから自動送信されています。
返信はできませんので、ご質問がある場合は人事部までお問い合わせください。`;
}

/**
 * 残日数アラートのHTMLメール生成
 */
function generateLowRemainingDaysAlertHtml(employee) {
  var urgencyLevel = employee.remaining <= 3 ? '緊急' : '注意';
  var color = employee.remaining <= 3 ? '#f44336' : '#ff9800';
  
  var html = `
  <html>
  <body style="font-family: Arial, sans-serif; color: #333;">
    <h2 style="color: ${color};">【${urgencyLevel}】有給残日数のお知らせ</h2>
    
    <p>あなたの有給残日数が少なくなっています。</p>
    
    <table style="border-collapse: collapse; margin: 20px 0;">
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f5;">氏名</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${employee.userName}</td>
      </tr>
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f5;">現在の残日数</td>
        <td style="padding: 8px; border: 1px solid #ddd; color: ${color}; font-weight: bold; font-size: 18px;">${employee.remaining}日</td>
      </tr>
    </table>
    
    ${employee.remaining <= 3 ?
      '<p style="color: #f44336;"><strong>残日数が3日以下です。失効前に計画的に取得することをお勧めします。</strong></p>' :
      '<p style="color: #ff9800;"><strong>残日数が5日以下です。今後の取得予定を検討してください。</strong></p>'
    }
    
    <p>
      有給申請は管理システムから行えます。<br>
      ご不明な点がございましたら、人事部までお問い合わせください。
    </p>
    
    <p style="color: #666; font-size: 12px; margin-top: 40px;">
      このメールは有給管理システムから自動送信されています。<br>
      返信はできませんので、ご質問がある場合は人事部までお問い合わせください。
    </p>
  </body>
  </html>`;
  
  return html;
}

/**
 * 残日数アラートのテキストメール生成
 */
function generateLowRemainingDaysAlertText(employee) {
  var urgencyLevel = employee.remaining <= 3 ? '緊急' : '注意';
  
  return `【${urgencyLevel}】有給残日数のお知らせ

あなたの有給残日数が少なくなっています。

氏名: ${employee.userName}
現在の残日数: ${employee.remaining}日

${employee.remaining <= 3 ?
  '残日数が3日以下です。失効前に計画的に取得することをお勧めします。' :
  '残日数が5日以下です。今後の取得予定を検討してください。'
}

有給申請は管理システムから行えます。
ご不明な点がございましたら、人事部までお問い合わせください。

※このメールは有給管理システムから自動送信されています。
返信はできませんので、ご質問がある場合は人事部までお問い合わせください。`;
}

// =============================
// ヘルパー関数
// =============================

/**
 * 承認者のメールアドレスを取得
 */
function getApproverEmail(userId) {
  // 実際の環境では適切なロジックに変更
  // 例: マスターシートから上司情報を取得
  
  if (NOTIFICATION_CONFIG.TEST_MODE) {
    return 'approver@test.com';
  }
  
  // 事業所別の承認者設定（デモ版）
  var approvers = {
    'R': 'rise-manager@company.com',
    'P': 'paron-manager@company.com', 
    'S': 'ciel-manager@company.com',
    'E': 'ebisu-manager@company.com'
  };
  
  var division = userId.substring(0, 1);
  return approvers[division] || 'hr@company.com';
}

/**
 * 社員のメールアドレスを取得
 */
function getEmployeeEmail(userId) {
  // 実際の環境では適切なロジックに変更
  // 例: マスターシートからメールアドレスを取得
  
  if (NOTIFICATION_CONFIG.TEST_MODE) {
    return 'employee@test.com';
  }
  
  // 仮の実装（実際は適切なデータソースから取得）
  return userId.toLowerCase() + '@company.com';
}

/**
 * 通知機能のテスト
 */
function testNotificationSystem() {
  try {
    console.log('=== 通知機能テスト開始 ===');
    
    // テスト用の申請データ
    var testApplication = {
      userId: 'R01',
      userName: '福島英昭',
      applyDate: '2024/08/15',
      applyDays: 1,
      timestamp: '2024/08/11 10:30:00'
    };
    
    var testEmployee = {
      userId: 'P01',
      userName: '牧野洋一',
      remaining: 2
    };
    
    var results = [];
    
    // 申請通知テスト
    var appNotifyResult = sendApplicationNotification(testApplication);
    results.push({
      test: '申請通知',
      result: appNotifyResult
    });
    
    // 承認結果通知テスト
    var approvalResult = sendApprovalResultNotification(testApplication, 'Approved');
    results.push({
      test: '承認結果通知',
      result: approvalResult
    });
    
    // 残日数アラートテスト
    var alertResult = sendLowRemainingDaysAlert(testEmployee);
    results.push({
      test: '残日数アラート',
      result: alertResult
    });
    
    console.log('=== 通知機能テスト結果 ===');
    results.forEach(function(item, index) {
      console.log((index + 1) + '. ' + item.test + ':', item.result);
    });
    
    var successCount = results.filter(function(item) {
      return item.result.success;
    }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功',
      message: NOTIFICATION_CONFIG.TEST_MODE ? 
        'テストモードで実行しました（実際のメール送信は行われていません）' :
        'すべての通知機能をテストしました'
    };
    
  } catch (error) {
    console.error('通知機能テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 残日数アラート対象者を抽出して一括通知
 */
function sendBulkLowRemainingDaysAlerts() {
  try {
    console.log('=== 残日数アラート一括送信開始 ===');
    
    // マスターシートから残日数が少ない社員を抽出
    var lowRemainingEmployees = getLowRemainingDaysEmployees();
    
    if (lowRemainingEmployees.length === 0) {
      return {
        success: true,
        message: '残日数アラート対象者はいません',
        sentCount: 0
      };
    }
    
    var results = [];
    var successCount = 0;
    
    lowRemainingEmployees.forEach(function(employee) {
      var result = sendLowRemainingDaysAlert(employee);
      results.push({
        employee: employee,
        result: result
      });
      
      if (result.success) {
        successCount++;
      }
    });
    
    console.log('残日数アラート一括送信完了:', successCount + '/' + lowRemainingEmployees.length);
    
    return {
      success: true,
      message: successCount + '/' + lowRemainingEmployees.length + '人にアラートを送信しました',
      results: results,
      sentCount: successCount,
      totalTargets: lowRemainingEmployees.length
    };
    
  } catch (error) {
    console.error('残日数アラート一括送信エラー:', error);
    return {
      success: false,
      message: 'アラート送信中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 残日数が少ない社員を抽出
 */
function getLowRemainingDaysEmployees() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var masterSheet = ss.getSheetByName('マスター');
    
    if (!masterSheet) {
      throw new Error('マスターシートが見つかりません');
    }
    
    var data = masterSheet.getDataRange().getValues();
    var lowRemainingEmployees = [];
    
    // ヘッダーをスキップして処理
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue; // 利用者番号が空の場合はスキップ
      
      var userId = String(row[0]);
      var userName = String(row[1] || '');
      var remaining = Number(row[2] || 0);
      
      // 残日数が5日以下の場合はアラート対象
      if (remaining <= 5 && remaining > 0) {
        lowRemainingEmployees.push({
          userId: userId,
          userName: userName,
          remaining: remaining
        });
      }
    }
    
    console.log('残日数アラート対象者:', lowRemainingEmployees.length + '人');
    return lowRemainingEmployees;
    
  } catch (error) {
    console.error('残日数アラート対象者抽出エラー:', error);
    return [];
  }
}

/**
 * 付与通知メールを送信
 * @param {Object} grantData - 付与データ
 * @return {Object} 送信結果
 */
function sendGrantNotification(grantData) {
  try {
    console.log('付与通知メール送信開始:', grantData);
    
    // 通知機能が無効な場合
    if (!NOTIFICATION_CONFIG.ENABLE_NOTIFICATIONS || !NOTIFICATION_CONFIG.NOTIFICATION_SETTINGS.GRANT) {
      return {
        success: true,
        message: '付与通知が無効になっています',
        disabled: true
      };
    }
    
    if (NOTIFICATION_CONFIG.TEST_MODE) {
      console.log('TEST MODE: メール送信をスキップします');
      return {
        success: true,
        message: 'テストモード: メール送信をスキップしました',
        testMode: true
      };
    }
    
    var employeeEmail = getEmployeeEmail(grantData.userId);
    if (!employeeEmail) {
      throw new Error('社員のメールアドレスが見つかりません');
    }
    
    var subject = NOTIFICATION_CONFIG.SUBJECT_PREFIX + '有給付与のお知らせ';
    var htmlBody = generateGrantNotificationHtml(grantData);
    var plainBody = generateGrantNotificationText(grantData);
    
    // メール送信
    GmailApp.sendEmail(
      employeeEmail,
      subject,
      plainBody,
      {
        htmlBody: htmlBody,
        name: '有給管理システム'
      }
    );
    
    console.log('付与通知メール送信完了:', employeeEmail);
    
    return {
      success: true,
      message: '付与通知メールを送信しました',
      recipient: employeeEmail
    };
    
  } catch (error) {
    console.error('付与通知メール送信エラー:', error);
    return {
      success: false,
      message: 'メール送信に失敗しました: ' + error.message
    };
  }
}

/**
 * 失効通知メールを送信
 * @param {Array} expiredUsers - 失効対象者リスト
 * @return {Object} 送信結果
 */
function sendExpiryNotifications(expiredUsers) {
  try {
    console.log('失効通知メール一括送信開始:', expiredUsers.length + '名');
    
    // 通知機能が無効な場合
    if (!NOTIFICATION_CONFIG.ENABLE_NOTIFICATIONS || !NOTIFICATION_CONFIG.NOTIFICATION_SETTINGS.EXPIRY) {
      return {
        success: true,
        message: '失効通知が無効になっています',
        disabled: true
      };
    }
    
    if (NOTIFICATION_CONFIG.TEST_MODE) {
      console.log('TEST MODE: メール送信をスキップします');
      return {
        success: true,
        message: 'テストモード: メール送信をスキップしました',
        testMode: true
      };
    }
    
    var results = [];
    var successCount = 0;
    
    expiredUsers.forEach(function(expiredUser) {
      try {
        var employeeEmail = getEmployeeEmail(expiredUser.userId);
        if (!employeeEmail) {
          results.push({
            userId: expiredUser.userId,
            success: false,
            reason: 'メールアドレスが見つかりません'
          });
          return;
        }
        
        var subject = NOTIFICATION_CONFIG.SUBJECT_PREFIX + '有給失効のお知らせ';
        var htmlBody = generateExpiryNotificationHtml(expiredUser);
        var plainBody = generateExpiryNotificationText(expiredUser);
        
        // メール送信
        GmailApp.sendEmail(
          employeeEmail,
          subject,
          plainBody,
          {
            htmlBody: htmlBody,
            name: '有給管理システム'
          }
        );
        
        successCount++;
        results.push({
          userId: expiredUser.userId,
          success: true,
          recipient: employeeEmail
        });
        
        console.log('失効通知メール送信完了:', expiredUser.userId, employeeEmail);
        
      } catch (error) {
        console.error('失効通知メール送信エラー:', expiredUser.userId, error);
        results.push({
          userId: expiredUser.userId,
          success: false,
          reason: error.message
        });
      }
    });
    
    return {
      success: true,
      totalTargets: expiredUsers.length,
      successCount: successCount,
      results: results,
      message: successCount + '/' + expiredUsers.length + '名に失効通知を送信しました'
    };
    
  } catch (error) {
    console.error('失効通知一括送信エラー:', error);
    return {
      success: false,
      message: '失効通知送信中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 付与通知のHTMLメール生成
 */
function generateGrantNotificationHtml(grantData) {
  var html = `
  <html>
  <body style="font-family: Arial, sans-serif; color: #333;">
    <h2 style="color: #4CAF50;">有給付与のお知らせ</h2>
    
    <p>有給休暇の付与をお知らせします。</p>
    
    <table style="border-collapse: collapse; margin: 20px 0;">
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f5;">氏名</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${grantData.userName}</td>
      </tr>
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f5;">付与タイプ</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${grantData.grantType}</td>
      </tr>
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f5;">付与日数</td>
        <td style="padding: 8px; border: 1px solid #ddd; color: #4CAF50; font-weight: bold; font-size: 18px;">${grantData.grantDays}日</td>
      </tr>
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f5;">付与日</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${grantData.grantDate}</td>
      </tr>
    </table>
    
    <p style="color: #4CAF50;"><strong>計画的な有給取得をお勧めします。</strong></p>
    
    <p style="color: #666; font-size: 12px; margin-top: 40px;">
      このメールは有給管理システムから自動送信されています。<br>
      返信はできませんので、ご質問がある場合は人事部までお問い合わせください。
    </p>
  </body>
  </html>`;
  
  return html;
}

/**
 * 付与通知のテキストメール生成
 */
function generateGrantNotificationText(grantData) {
  return `有給付与のお知らせ

有給休暇の付与をお知らせします。

氏名: ${grantData.userName}
付与タイプ: ${grantData.grantType}
付与日数: ${grantData.grantDays}日
付与日: ${grantData.grantDate}

計画的な有給取得をお勧めします。

※このメールは有給管理システムから自動送信されています。
返信はできませんので、ご質問がある場合は人事部までお問い合わせください。`;
}

/**
 * 失効通知のHTMLメール生成
 */
function generateExpiryNotificationHtml(expiredUser) {
  var html = `
  <html>
  <body style="font-family: Arial, sans-serif; color: #333;">
    <h2 style="color: #f44336;">有給失効のお知らせ</h2>
    
    <p>以下の有給休暇が失効しました。</p>
    
    <table style="border-collapse: collapse; margin: 20px 0;">
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f5;">氏名</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${expiredUser.userId}</td>
      </tr>
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f5;">失効日数</td>
        <td style="padding: 8px; border: 1px solid #ddd; color: #f44336; font-weight: bold; font-size: 18px;">${expiredUser.expiredDays}日</td>
      </tr>
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f5;">失効日</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${Utilities.formatDate(expiredUser.expiryDate, 'JST', 'yyyy年M月d日')}</td>
      </tr>
    </table>
    
    <p style="color: #f44336;"><strong>今後は計画的な有給取得をお勧めします。</strong></p>
    
    <p style="color: #666; font-size: 12px; margin-top: 40px;">
      このメールは有給管理システムから自動送信されています。<br>
      返信はできませんので、ご質問がある場合は人事部までお問い合わせください。
    </p>
  </body>
  </html>`;
  
  return html;
}

/**
 * 失効通知のテキストメール生成
 */
function generateExpiryNotificationText(expiredUser) {
  return `有給失効のお知らせ

以下の有給休暇が失効しました。

氏名: ${expiredUser.userId}
失効日数: ${expiredUser.expiredDays}日
失効日: ${Utilities.formatDate(expiredUser.expiryDate, 'JST', 'yyyy年M月d日')}

今後は計画的な有給取得をお勧めします。

※このメールは有給管理システムから自動送信されています。
返信はできませんので、ご質問がある場合は人事部までお問い合わせください。`;
}

/**
 * 通知設定を管理する関数
 */
function updateNotificationSettings(settings) {
  try {
    console.log('通知設定更新:', settings);
    
    // 設定をPropertiesServiceに保存
    PropertiesService.getScriptProperties().setProperty(
      'NOTIFICATION_SETTINGS',
      JSON.stringify(settings)
    );
    
    return {
      success: true,
      message: '通知設定を更新しました',
      settings: settings
    };
    
  } catch (error) {
    console.error('通知設定更新エラー:', error);
    return {
      success: false,
      message: '通知設定の更新に失敗しました: ' + error.message
    };
  }
}

/**
 * 現在の通知設定を取得
 */
function getNotificationSettings() {
  try {
    var savedSettings = PropertiesService.getScriptProperties().getProperty('NOTIFICATION_SETTINGS');
    
    if (savedSettings) {
      var settings = JSON.parse(savedSettings);
      return {
        success: true,
        settings: settings,
        message: '通知設定を取得しました'
      };
    } else {
      return {
        success: true,
        settings: NOTIFICATION_CONFIG.NOTIFICATION_SETTINGS,
        message: 'デフォルト通知設定を取得しました'
      };
    }
    
  } catch (error) {
    console.error('通知設定取得エラー:', error);
    return {
      success: false,
      message: '通知設定の取得に失敗しました: ' + error.message
    };
  }
}