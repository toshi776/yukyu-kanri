// =============================
// 通知機能本格運用化テスト
// =============================

/**
 * 通知機能本格運用化の全テストを実行
 */
function runNotificationProductionTests() {
  console.log('=== 通知機能本格運用化テスト開始 ===');
  
  var testResults = [];
  
  try {
    // テスト1: 通知設定機能のテスト
    var settingsTest = testNotificationSettings();
    testResults.push({
      testName: '通知設定機能',
      result: settingsTest
    });
    
    // テスト2: 付与通知機能のテスト
    var grantNotificationTest = testGrantNotification();
    testResults.push({
      testName: '付与通知機能',
      result: grantNotificationTest
    });
    
    // テスト3: 失効通知機能のテスト
    var expiryNotificationTest = testExpiryNotification();
    testResults.push({
      testName: '失効通知機能',
      result: expiryNotificationTest
    });
    
    // テスト4: 通知無効化機能のテスト
    var disableTest = testNotificationDisable();
    testResults.push({
      testName: '通知無効化機能',
      result: disableTest
    });
    
    // テスト5: 統合通知フローのテスト
    var integrationTest = testNotificationIntegration();
    testResults.push({
      testName: '統合通知フロー',
      result: integrationTest
    });
    
  } catch (error) {
    console.error('通知機能テスト実行エラー:', error);
    testResults.push({
      testName: 'テスト実行',
      result: {
        success: false,
        message: 'テスト実行中にエラーが発生しました: ' + error.message
      }
    });
  }
  
  // 結果サマリーを表示
  console.log('\n=== 通知機能テスト結果サマリー ===');
  var totalTests = testResults.length;
  var passedTests = 0;
  
  testResults.forEach(function(test, index) {
    var status = test.result.success ? '✅ PASS' : '❌ FAIL';
    console.log((index + 1) + '. ' + test.testName + ': ' + status);
    
    if (test.result.success) {
      passedTests++;
    } else {
      console.log('   エラー: ' + test.result.message);
    }
  });
  
  console.log('\n総合結果: ' + passedTests + '/' + totalTests + ' テスト成功');
  console.log('成功率: ' + Math.round((passedTests / totalTests) * 100) + '%');
  
  return {
    success: passedTests === totalTests,
    totalTests: totalTests,
    passedTests: passedTests,
    failedTests: totalTests - passedTests,
    results: testResults
  };
}

/**
 * 通知設定機能のテスト
 */
function testNotificationSettings() {
  try {
    console.log('通知設定機能テスト開始');
    
    var results = [];
    
    // テスト1: デフォルト設定取得
    console.log('1. デフォルト設定取得テスト');
    var getResult = getNotificationSettings();
    
    results.push({
      test: 'デフォルト設定取得',
      result: getResult.success ? 'SUCCESS' : 'FAILED',
      detail: getResult.message
    });
    
    // テスト2: 設定更新
    console.log('2. 設定更新テスト');
    var testSettings = {
      APPLICATION: true,
      APPROVAL: true,
      ALERT: false,    // アラートを無効にしてテスト
      GRANT: true,
      EXPIRY: true
    };
    
    var updateResult = updateNotificationSettings(testSettings);
    
    results.push({
      test: '設定更新',
      result: updateResult.success ? 'SUCCESS' : 'FAILED',
      detail: updateResult.message
    });
    
    // テスト3: 更新された設定の確認
    console.log('3. 更新設定確認テスト');
    var getUpdatedResult = getNotificationSettings();
    var isCorrect = getUpdatedResult.success && 
                   getUpdatedResult.settings.ALERT === false;
    
    results.push({
      test: '更新設定確認',
      result: isCorrect ? 'SUCCESS' : 'FAILED',
      detail: 'アラート設定: ' + getUpdatedResult.settings.ALERT
    });
    
    // テスト4: 設定を元に戻す
    console.log('4. 設定復元テスト');
    var defaultSettings = {
      APPLICATION: true,
      APPROVAL: true,
      ALERT: true,
      GRANT: true,
      EXPIRY: true
    };
    
    var restoreResult = updateNotificationSettings(defaultSettings);
    
    results.push({
      test: '設定復元',
      result: restoreResult.success ? 'SUCCESS' : 'FAILED',
      detail: restoreResult.message
    });
    
    // 結果集計
    console.log('=== 通知設定機能テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功'
    };
    
  } catch (error) {
    console.error('通知設定機能テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 付与通知機能のテスト
 */
function testGrantNotification() {
  try {
    console.log('付与通知機能テスト開始');
    
    var testGrantData = {
      userId: 'R01',
      userName: '福島英昭',
      grantDays: 11,
      grantDate: '2024年8月18日',
      grantType: '年次付与（2年目）'
    };
    
    var results = [];
    
    // テスト1: 付与通知送信（本格運用モード）
    console.log('1. 付与通知送信テスト');
    var grantResult = sendGrantNotification(testGrantData);
    
    results.push({
      test: '付与通知送信',
      result: grantResult.success ? 'SUCCESS' : 'FAILED',
      detail: grantResult.message
    });
    
    // テスト2: HTMLテンプレート生成
    console.log('2. HTMLテンプレート生成テスト');
    var htmlBody = generateGrantNotificationHtml(testGrantData);
    var hasUserName = htmlBody.indexOf(testGrantData.userName) !== -1;
    var hasGrantDays = htmlBody.indexOf(testGrantData.grantDays) !== -1;
    
    results.push({
      test: 'HTMLテンプレート生成',
      result: (hasUserName && hasGrantDays) ? 'SUCCESS' : 'FAILED',
      detail: '必要項目の埋め込み: ' + (hasUserName && hasGrantDays ? '完了' : '不完全')
    });
    
    // テスト3: テキストテンプレート生成
    console.log('3. テキストテンプレート生成テスト');
    var textBody = generateGrantNotificationText(testGrantData);
    var hasUserNameText = textBody.indexOf(testGrantData.userName) !== -1;
    var hasGrantDaysText = textBody.indexOf(testGrantData.grantDays) !== -1;
    
    results.push({
      test: 'テキストテンプレート生成',
      result: (hasUserNameText && hasGrantDaysText) ? 'SUCCESS' : 'FAILED',
      detail: '必要項目の埋め込み: ' + (hasUserNameText && hasGrantDaysText ? '完了' : '不完全')
    });
    
    // 結果集計
    console.log('=== 付与通知機能テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功'
    };
    
  } catch (error) {
    console.error('付与通知機能テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 失効通知機能のテスト
 */
function testExpiryNotification() {
  try {
    console.log('失効通知機能テスト開始');
    
    var testExpiredUsers = [
      {
        userId: 'P01',
        expiredDays: 5,
        expiryDate: new Date('2024-08-15')
      },
      {
        userId: 'S01',
        expiredDays: 3,
        expiryDate: new Date('2024-08-15')
      }
    ];
    
    var results = [];
    
    // テスト1: 失効通知一括送信
    console.log('1. 失効通知一括送信テスト');
    var expiryResult = sendExpiryNotifications(testExpiredUsers);
    
    results.push({
      test: '失効通知一括送信',
      result: expiryResult.success ? 'SUCCESS' : 'FAILED',
      detail: expiryResult.message
    });
    
    // テスト2: 失効通知HTMLテンプレート
    console.log('2. 失効通知HTMLテンプレートテスト');
    var expiredUser = testExpiredUsers[0];
    var htmlBody = generateExpiryNotificationHtml(expiredUser);
    var hasUserId = htmlBody.indexOf(expiredUser.userId) !== -1;
    var hasExpiredDays = htmlBody.indexOf(expiredUser.expiredDays) !== -1;
    
    results.push({
      test: '失効通知HTMLテンプレート',
      result: (hasUserId && hasExpiredDays) ? 'SUCCESS' : 'FAILED',
      detail: '必要項目の埋め込み: ' + (hasUserId && hasExpiredDays ? '完了' : '不完全')
    });
    
    // テスト3: 失効通知テキストテンプレート
    console.log('3. 失効通知テキストテンプレートテスト');
    var textBody = generateExpiryNotificationText(expiredUser);
    var hasUserIdText = textBody.indexOf(expiredUser.userId) !== -1;
    var hasExpiredDaysText = textBody.indexOf(expiredUser.expiredDays) !== -1;
    
    results.push({
      test: '失効通知テキストテンプレート',
      result: (hasUserIdText && hasExpiredDaysText) ? 'SUCCESS' : 'FAILED',
      detail: '必要項目の埋め込み: ' + (hasUserIdText && hasExpiredDaysText ? '完了' : '不完全')
    });
    
    // 結果集計
    console.log('=== 失効通知機能テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功'
    };
    
  } catch (error) {
    console.error('失効通知機能テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 通知無効化機能のテスト
 */
function testNotificationDisable() {
  try {
    console.log('通知無効化機能テスト開始');
    
    var results = [];
    
    // テスト1: アラート通知を無効にして送信
    console.log('1. アラート通知無効化テスト');
    
    // アラートを無効に設定
    var disableSettings = {
      APPLICATION: true,
      APPROVAL: true,
      ALERT: false,  // アラートを無効
      GRANT: true,
      EXPIRY: true
    };
    
    updateNotificationSettings(disableSettings);
    
    var testEmployee = {
      userId: 'TEST001',
      userName: 'テスト太郎',
      remaining: 3
    };
    
    var alertResult = sendLowRemainingDaysAlert(testEmployee);
    var isDisabled = alertResult.disabled === true;
    
    results.push({
      test: 'アラート通知無効化',
      result: isDisabled ? 'SUCCESS' : 'FAILED',
      detail: alertResult.message
    });
    
    // テスト2: 付与通知を無効にして送信
    console.log('2. 付与通知無効化テスト');
    
    var disableGrantSettings = {
      APPLICATION: true,
      APPROVAL: true,
      ALERT: true,
      GRANT: false,  // 付与通知を無効
      EXPIRY: true
    };
    
    updateNotificationSettings(disableGrantSettings);
    
    var testGrantData = {
      userId: 'TEST002',
      userName: 'テスト花子',
      grantDays: 10
    };
    
    var grantResult = sendGrantNotification(testGrantData);
    var isGrantDisabled = grantResult.disabled === true;
    
    results.push({
      test: '付与通知無効化',
      result: isGrantDisabled ? 'SUCCESS' : 'FAILED',
      detail: grantResult.message
    });
    
    // テスト3: 設定を元に戻す
    console.log('3. 通知設定復元テスト');
    var enableAllSettings = {
      APPLICATION: true,
      APPROVAL: true,
      ALERT: true,
      GRANT: true,
      EXPIRY: true
    };
    
    var restoreResult = updateNotificationSettings(enableAllSettings);
    
    results.push({
      test: '通知設定復元',
      result: restoreResult.success ? 'SUCCESS' : 'FAILED',
      detail: restoreResult.message
    });
    
    // 結果集計
    console.log('=== 通知無効化機能テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功'
    };
    
  } catch (error) {
    console.error('通知無効化機能テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 統合通知フローのテスト
 */
function testNotificationIntegration() {
  try {
    console.log('統合通知フローテスト開始');
    
    var results = [];
    
    // テスト1: 申請通知フロー
    console.log('1. 申請通知フローテスト');
    var testApplication = {
      userId: 'R01',
      userName: '福島英昭',
      applyDate: '2024/08/15',
      applyDays: 1,
      timestamp: '2024/08/11 10:30:00'
    };
    
    var appResult = sendApplicationNotification(testApplication);
    
    results.push({
      test: '申請通知フロー',
      result: appResult.success ? 'SUCCESS' : 'FAILED',
      detail: appResult.message
    });
    
    // テスト2: 承認結果通知フロー
    console.log('2. 承認結果通知フローテスト');
    var approvalResult = sendApprovalResultNotification(testApplication, 'Approved');
    
    results.push({
      test: '承認結果通知フロー',
      result: approvalResult.success ? 'SUCCESS' : 'FAILED',
      detail: approvalResult.message
    });
    
    // テスト3: 一括アラート送信フロー
    console.log('3. 一括アラート送信フローテスト');
    var bulkAlertResult = sendBulkLowRemainingDaysAlerts();
    
    results.push({
      test: '一括アラート送信フロー',
      result: bulkAlertResult.success ? 'SUCCESS' : 'FAILED',
      detail: bulkAlertResult.message
    });
    
    // テスト4: 通知機能全体の設定確認
    console.log('4. 通知機能全体設定確認テスト');
    var settingsCheck = getNotificationSettings();
    var allEnabled = settingsCheck.success && 
                    settingsCheck.settings.APPLICATION &&
                    settingsCheck.settings.APPROVAL &&
                    settingsCheck.settings.ALERT &&
                    settingsCheck.settings.GRANT &&
                    settingsCheck.settings.EXPIRY;
    
    results.push({
      test: '通知機能全体設定確認',
      result: allEnabled ? 'SUCCESS' : 'FAILED',
      detail: '全通知機能: ' + (allEnabled ? '有効' : '一部無効')
    });
    
    // 結果集計
    console.log('=== 統合通知フローテスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功'
    };
    
  } catch (error) {
    console.error('統合通知フローテストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}
