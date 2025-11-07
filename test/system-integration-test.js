// =============================
// システム統合テスト
// =============================

/**
 * システム全体の統合テストを実行
 */
function runSystemIntegrationTests() {
  console.log('=== システム統合テスト開始 ===');
  
  var testResults = [];
  var startTime = new Date();
  
  try {
    // テスト1: 基盤システムの統合テスト
    console.log('\n1. 基盤システム統合テスト');
    var foundationTest = testFoundationSystemIntegration();
    testResults.push({
      testName: '基盤システム統合',
      result: foundationTest
    });
    
    // テスト2: 法定付与システムの統合テスト
    console.log('\n2. 法定付与システム統合テスト');
    var legalGrantTest = testLegalGrantSystemIntegration();
    testResults.push({
      testName: '法定付与システム統合',
      result: legalGrantTest
    });
    
    // テスト3: 通知システムの統合テスト
    console.log('\n3. 通知システム統合テスト');
    var notificationTest = testNotificationSystemIntegration();
    testResults.push({
      testName: '通知システム統合',
      result: notificationTest
    });
    
    // テスト4: 統計レポートシステムの統合テスト
    console.log('\n4. 統計レポートシステム統合テスト');
    var reportTest = testReportSystemIntegration();
    testResults.push({
      testName: '統計レポートシステム統合',
      result: reportTest
    });
    
    // テスト5: データ整合性の統合テスト
    console.log('\n5. データ整合性統合テスト');
    var dataIntegrityTest = testDataIntegritySystem();
    testResults.push({
      testName: 'データ整合性システム',
      result: dataIntegrityTest
    });
    
    // テスト6: パフォーマンステスト
    console.log('\n6. パフォーマンステスト');
    var performanceTest = testSystemPerformance();
    testResults.push({
      testName: 'システムパフォーマンス',
      result: performanceTest
    });
    
    // テスト7: セキュリティテスト
    console.log('\n7. セキュリティテスト');
    var securityTest = testSystemSecurity();
    testResults.push({
      testName: 'システムセキュリティ',
      result: securityTest
    });
    
  } catch (error) {
    console.error('統合テスト実行エラー:', error);
    testResults.push({
      testName: 'テスト実行',
      result: {
        success: false,
        message: 'テスト実行中にエラーが発生しました: ' + error.message
      }
    });
  }
  
  var endTime = new Date();
  var duration = Math.round((endTime - startTime) / 1000);
  
  // 結果サマリーを表示
  console.log('\n=== システム統合テスト結果サマリー ===');
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
  console.log('実行時間: ' + duration + '秒');
  
  // 統合テスト結果をPropertiesServiceに保存
  var integrationResult = {
    testDate: new Date().toISOString(),
    totalTests: totalTests,
    passedTests: passedTests,
    failedTests: totalTests - passedTests,
    successRate: Math.round((passedTests / totalTests) * 100),
    duration: duration,
    results: testResults
  };
  
  PropertiesService.getScriptProperties().setProperty(
    'LAST_INTEGRATION_TEST',
    JSON.stringify(integrationResult)
  );
  
  return {
    success: passedTests === totalTests,
    totalTests: totalTests,
    passedTests: passedTests,
    failedTests: totalTests - passedTests,
    successRate: Math.round((passedTests / totalTests) * 100),
    duration: duration,
    results: testResults
  };
}

/**
 * 基盤システムの統合テスト
 */
function testFoundationSystemIntegration() {
  try {
    console.log('基盤システム統合テスト開始');
    
    var results = [];
    
    // URL管理システム
    var urlStatus = getTriggerStatus();
    results.push({
      test: 'URL管理システム',
      result: urlStatus.success ? 'SUCCESS' : 'FAILED',
      detail: 'URL管理機能: ' + (urlStatus.success ? '正常' : 'エラー')
    });
    
    // アクセス制御システム
    var testUrlKey = generateUrlKey();
    var accessControlWorking = typeof testUrlKey === 'string' && testUrlKey.length === 32;
    results.push({
      test: 'アクセス制御システム',
      result: accessControlWorking ? 'SUCCESS' : 'FAILED',
      detail: 'URLキー生成: ' + (accessControlWorking ? '正常' : 'エラー')
    });
    
    // データベース接続
    var ss = getSpreadsheet();
    var dbConnection = ss !== null;
    results.push({
      test: 'データベース接続',
      result: dbConnection ? 'SUCCESS' : 'FAILED',
      detail: 'スプレッドシート接続: ' + (dbConnection ? '成功' : '失敗')
    });
    
    // マスターデータ構造
    if (dbConnection) {
      var masterSheet = ss.getSheetByName('マスター');
      var masterStructure = masterSheet && masterSheet.getLastColumn() >= 4;
      results.push({
        test: 'マスターデータ構造',
        result: masterStructure ? 'SUCCESS' : 'FAILED',
        detail: 'マスターシート構造: ' + (masterStructure ? '適切' : '不適切')
      });
    }
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' 基盤システムテスト成功'
    };
    
  } catch (error) {
    console.error('基盤システム統合テストエラー:', error);
    return {
      success: false,
      message: '基盤システムテスト中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 法定付与システムの統合テスト
 */
function testLegalGrantSystemIntegration() {
  try {
    console.log('法定付与システム統合テスト開始');
    
    var results = [];
    
    // 付与ルール計算システム
    var testGrantDays = calculateLeaveDays(1.5, 5, false);
    var grantCalculation = testGrantDays === 11;
    results.push({
      test: '付与ルール計算システム',
      result: grantCalculation ? 'SUCCESS' : 'FAILED',
      detail: '1.5年勤務・週5日: ' + testGrantDays + '日付与 (期待値: 11日)'
    });
    
    // 付与履歴管理システム
    var ss = getSpreadsheet();
    var grantHistorySheet = getOrCreateGrantHistorySheet(ss);
    var historyManagement = grantHistorySheet !== null;
    results.push({
      test: '付与履歴管理システム',
      result: historyManagement ? 'SUCCESS' : 'FAILED',
      detail: '付与履歴シート: ' + (historyManagement ? '作成・取得成功' : '失敗')
    });
    
    // FIFO消費システム
    var testRemaining = calculateEffectiveRemainingDays('TEST_USER');
    var fifoSystem = typeof testRemaining === 'number' && testRemaining >= 0;
    results.push({
      test: 'FIFO消費システム',
      result: fifoSystem ? 'SUCCESS' : 'FAILED',
      detail: '残日数計算: ' + (fifoSystem ? '正常動作' : 'エラー')
    });
    
    // 失効処理システム
    var expiredResult = processExpiredLeaves();
    var expirySystem = expiredResult.success;
    results.push({
      test: '失効処理システム',
      result: expirySystem ? 'SUCCESS' : 'FAILED',
      detail: '失効処理: ' + expiredResult.message
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' 法定付与システムテスト成功'
    };
    
  } catch (error) {
    console.error('法定付与システム統合テストエラー:', error);
    return {
      success: false,
      message: '法定付与システムテスト中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 通知システムの統合テスト
 */
function testNotificationSystemIntegration() {
  try {
    console.log('通知システム統合テスト開始');
    
    var results = [];
    
    // 通知設定管理
    var settingsResult = getNotificationSettings();
    var settingsManagement = settingsResult.success;
    results.push({
      test: '通知設定管理',
      result: settingsManagement ? 'SUCCESS' : 'FAILED',
      detail: '設定管理: ' + settingsResult.message
    });
    
    // メールテンプレート生成
    var testGrantData = {
      userId: 'TEST001',
      userName: 'テスト太郎',
      grantDays: 10,
      grantDate: '2024年8月18日',
      grantType: 'テスト付与'
    };
    
    var htmlTemplate = generateGrantNotificationHtml(testGrantData);
    var templateGeneration = typeof htmlTemplate === 'string' && htmlTemplate.length > 0;
    results.push({
      test: 'メールテンプレート生成',
      result: templateGeneration ? 'SUCCESS' : 'FAILED',
      detail: 'HTMLテンプレート: ' + (templateGeneration ? '生成成功' : '生成失敗')
    });
    
    // 通知送信システム（テストモード）
    var notifyResult = sendGrantNotification(testGrantData);
    var notificationSending = notifyResult.success;
    results.push({
      test: '通知送信システム',
      result: notificationSending ? 'SUCCESS' : 'FAILED',
      detail: '通知送信: ' + notifyResult.message
    });
    
    // 一括通知システム
    var bulkAlertResult = sendBulkLowRemainingDaysAlerts();
    var bulkNotification = bulkAlertResult.success;
    results.push({
      test: '一括通知システム',
      result: bulkNotification ? 'SUCCESS' : 'FAILED',
      detail: '一括通知: ' + bulkAlertResult.message
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' 通知システムテスト成功'
    };
    
  } catch (error) {
    console.error('通知システム統合テストエラー:', error);
    return {
      success: false,
      message: '通知システムテスト中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 統計レポートシステムの統合テスト
 */
function testReportSystemIntegration() {
  try {
    console.log('統計レポートシステム統合テスト開始');
    
    var results = [];
    
    // 基本統計計算
    var basicStats = calculateBasicStatistics(new Date());
    var statsCalculation = basicStats.hasOwnProperty('totalEmployees') && 
                          basicStats.hasOwnProperty('averageRemainingDays');
    results.push({
      test: '基本統計計算',
      result: statsCalculation ? 'SUCCESS' : 'FAILED',
      detail: '統計計算: 従業員数' + basicStats.totalEmployees + '名, 平均残日数' + basicStats.averageRemainingDays + '日'
    });
    
    // 月次レポート生成
    var monthlyReport = generateMonthlyReport(2024, 8);
    var reportGeneration = monthlyReport.success;
    results.push({
      test: '月次レポート生成',
      result: reportGeneration ? 'SUCCESS' : 'FAILED',
      detail: 'レポート生成: ' + monthlyReport.message
    });
    
    // HTMLレポート生成
    if (reportGeneration) {
      var htmlReport = generateReportHTML(monthlyReport.reportData);
      var htmlGeneration = typeof htmlReport === 'string' && htmlReport.length > 1000;
      results.push({
        test: 'HTMLレポート生成',
        result: htmlGeneration ? 'SUCCESS' : 'FAILED',
        detail: 'HTMLレポート: ' + htmlReport.length + '文字生成'
      });
    }
    
    // 年5日取得義務監視
    var fiveDayStats = calculateFiveDayObligationStats(new Date('2024-04-01'), new Date('2025-03-31'));
    var obligationMonitoring = fiveDayStats.hasOwnProperty('complianceRate');
    results.push({
      test: '年5日取得義務監視',
      result: obligationMonitoring ? 'SUCCESS' : 'FAILED',
      detail: '義務監視: 対象者' + fiveDayStats.targetEmployees + '名, 達成率' + fiveDayStats.complianceRate + '%'
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' レポートシステムテスト成功'
    };
    
  } catch (error) {
    console.error('統計レポートシステム統合テストエラー:', error);
    return {
      success: false,
      message: '統計レポートシステムテスト中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * データ整合性システムのテスト
 */
function testDataIntegritySystem() {
  try {
    console.log('データ整合性システムテスト開始');
    
    var results = [];
    
    // データ整合性チェック
    var integrityCheck = checkDataIntegrity();
    var dataIntegrity = integrityCheck.success;
    results.push({
      test: 'データ整合性チェック',
      result: dataIntegrity ? 'SUCCESS' : 'FAILED',
      detail: 'データ整合性: ' + integrityCheck.message
    });
    
    // マスターと履歴の整合性
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');
    var grantHistorySheet = ss.getSheetByName('付与履歴');
    var sheetsExist = masterSheet && grantHistorySheet;
    results.push({
      test: 'シート構造整合性',
      result: sheetsExist ? 'SUCCESS' : 'FAILED',
      detail: '必要シート: ' + (sheetsExist ? '全て存在' : '不足あり')
    });
    
    // 計算ロジック整合性
    var testCalculations = [
      { years: 0.5, days: 5, initial: true, expected: 10 },
      { years: 1.5, days: 5, initial: false, expected: 11 },
      { years: 3.5, days: 5, initial: false, expected: 14 }
    ];
    
    var calculationConsistency = true;
    testCalculations.forEach(function(test) {
      var result = calculateLeaveDays(test.years, test.days, test.initial);
      if (result !== test.expected) {
        calculationConsistency = false;
      }
    });
    
    results.push({
      test: '計算ロジック整合性',
      result: calculationConsistency ? 'SUCCESS' : 'FAILED',
      detail: '付与日数計算: ' + (calculationConsistency ? '全て正確' : '不整合あり')
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' データ整合性テスト成功'
    };
    
  } catch (error) {
    console.error('データ整合性システムテストエラー:', error);
    return {
      success: false,
      message: 'データ整合性システムテスト中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * システムパフォーマンステスト
 */
function testSystemPerformance() {
  try {
    console.log('システムパフォーマンステスト開始');
    
    var results = [];
    
    // レポート生成時間測定
    var startTime = new Date();
    var monthlyReport = generateMonthlyReport(2024, 8);
    var endTime = new Date();
    var reportGenerationTime = endTime - startTime;
    
    var reportPerformance = reportGenerationTime < 10000; // 10秒以内
    results.push({
      test: 'レポート生成時間',
      result: reportPerformance ? 'SUCCESS' : 'FAILED',
      detail: '生成時間: ' + Math.round(reportGenerationTime / 1000) + '秒 (目標: 10秒以内)'
    });
    
    // 統計計算時間測定
    startTime = new Date();
    var basicStats = calculateBasicStatistics(new Date());
    endTime = new Date();
    var statsCalculationTime = endTime - startTime;
    
    var statsPerformance = statsCalculationTime < 5000; // 5秒以内
    results.push({
      test: '統計計算時間',
      result: statsPerformance ? 'SUCCESS' : 'FAILED',
      detail: '計算時間: ' + Math.round(statsCalculationTime / 1000) + '秒 (目標: 5秒以内)'
    });
    
    // データアクセス時間測定
    startTime = new Date();
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');
    if (masterSheet) {
      var data = masterSheet.getDataRange().getValues();
    }
    endTime = new Date();
    var dataAccessTime = endTime - startTime;
    
    var accessPerformance = dataAccessTime < 3000; // 3秒以内
    results.push({
      test: 'データアクセス時間',
      result: accessPerformance ? 'SUCCESS' : 'FAILED',
      detail: 'アクセス時間: ' + Math.round(dataAccessTime / 1000) + '秒 (目標: 3秒以内)'
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' パフォーマンステスト成功'
    };
    
  } catch (error) {
    console.error('システムパフォーマンステストエラー:', error);
    return {
      success: false,
      message: 'システムパフォーマンステスト中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * システムセキュリティテスト
 */
function testSystemSecurity() {
  try {
    console.log('システムセキュリティテスト開始');
    
    var results = [];
    
    // URLキー生成セキュリティ
    var urlKey1 = generateUrlKey();
    var urlKey2 = generateUrlKey();
    var keyUniqueness = urlKey1 !== urlKey2 && urlKey1.length === 32;
    results.push({
      test: 'URLキー生成セキュリティ',
      result: keyUniqueness ? 'SUCCESS' : 'FAILED',
      detail: 'URLキー: ' + (keyUniqueness ? '一意性・長さ確保' : 'セキュリティ問題')
    });
    
    // データアクセス制御
    var testUserId = 'INVALID_USER';
    var invalidAccess = getUserIdFromUrlKey('invalid_key_test_123456789012');
    var accessControl = invalidAccess === null;
    results.push({
      test: 'データアクセス制御',
      result: accessControl ? 'SUCCESS' : 'FAILED',
      detail: '不正アクセス: ' + (accessControl ? '適切に拒否' : 'セキュリティ問題')
    });
    
    // PropertiesService暗号化
    try {
      PropertiesService.getScriptProperties().setProperty('TEST_SECURITY', 'test_value');
      var retrievedValue = PropertiesService.getScriptProperties().getProperty('TEST_SECURITY');
      var propertiesWorking = retrievedValue === 'test_value';
      PropertiesService.getScriptProperties().deleteProperty('TEST_SECURITY');
      
      results.push({
        test: 'PropertiesService保護',
        result: propertiesWorking ? 'SUCCESS' : 'FAILED',
        detail: 'データ保護: ' + (propertiesWorking ? '正常動作' : 'アクセス問題')
      });
    } catch (error) {
      results.push({
        test: 'PropertiesService保護',
        result: 'FAILED',
        detail: 'データ保護: エラー発生'
      });
    }
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' セキュリティテスト成功'
    };
    
  } catch (error) {
    console.error('システムセキュリティテストエラー:', error);
    return {
      success: false,
      message: 'システムセキュリティテスト中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 前回の統合テスト結果を取得
 */
function getLastIntegrationTestResult() {
  try {
    var lastResult = PropertiesService.getScriptProperties().getProperty('LAST_INTEGRATION_TEST');
    
    if (lastResult) {
      return {
        success: true,
        result: JSON.parse(lastResult)
      };
    } else {
      return {
        success: false,
        message: '前回の統合テスト結果が見つかりません'
      };
    }
    
  } catch (error) {
    console.error('統合テスト結果取得エラー:', error);
    return {
      success: false,
      message: '統合テスト結果の取得に失敗しました: ' + error.message
    };
  }
}

/**
 * システム全体のヘルスチェック
 */
function systemHealthCheck() {
  try {
    console.log('=== システムヘルスチェック開始 ===');
    
    var healthStatus = {
      timestamp: new Date().toISOString(),
      overallHealth: 'HEALTHY',
      components: {}
    };
    
    // 基盤システム
    try {
      var ss = getSpreadsheet();
      healthStatus.components.database = ss ? 'HEALTHY' : 'UNHEALTHY';
    } catch (error) {
      healthStatus.components.database = 'UNHEALTHY';
    }
    
    // URL管理システム
    try {
      var urlKey = generateUrlKey();
      healthStatus.components.urlManagement = (typeof urlKey === 'string' && urlKey.length === 32) ? 'HEALTHY' : 'UNHEALTHY';
    } catch (error) {
      healthStatus.components.urlManagement = 'UNHEALTHY';
    }
    
    // 通知システム
    try {
      var notificationSettings = getNotificationSettings();
      healthStatus.components.notification = notificationSettings.success ? 'HEALTHY' : 'UNHEALTHY';
    } catch (error) {
      healthStatus.components.notification = 'UNHEALTHY';
    }
    
    // 統計システム
    try {
      var basicStats = calculateBasicStatistics(new Date());
      healthStatus.components.statistics = basicStats.hasOwnProperty('totalEmployees') ? 'HEALTHY' : 'UNHEALTHY';
    } catch (error) {
      healthStatus.components.statistics = 'UNHEALTHY';
    }
    
    // トリガーシステム
    try {
      var triggerStatus = getTriggerStatus();
      healthStatus.components.triggers = triggerStatus.success ? 'HEALTHY' : 'UNHEALTHY';
    } catch (error) {
      healthStatus.components.triggers = 'UNHEALTHY';
    }
    
    // 全体ヘルス判定
    var unhealthyComponents = Object.keys(healthStatus.components).filter(function(key) {
      return healthStatus.components[key] === 'UNHEALTHY';
    });
    
    if (unhealthyComponents.length > 0) {
      healthStatus.overallHealth = 'DEGRADED';
      if (unhealthyComponents.length > 2) {
        healthStatus.overallHealth = 'UNHEALTHY';
      }
    }
    
    // ヘルス状況をPropertiesServiceに保存
    PropertiesService.getScriptProperties().setProperty(
      'SYSTEM_HEALTH_STATUS',
      JSON.stringify(healthStatus)
    );
    
    console.log('システムヘルスチェック完了:', healthStatus.overallHealth);
    
    return {
      success: true,
      healthStatus: healthStatus,
      message: 'システムヘルス: ' + healthStatus.overallHealth
    };
    
  } catch (error) {
    console.error('システムヘルスチェックエラー:', error);
    return {
      success: false,
      message: 'ヘルスチェックに失敗しました: ' + error.message
    };
  }
}