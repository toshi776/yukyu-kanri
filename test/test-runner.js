// =============================
// テスト実行スクリプト
// =============================

/**
 * 全テストを実行
 */
function runAllTests() {
  console.log('=== 全テスト開始 ===');
  
  var testResults = [];
  
  try {
    // 1. URL管理システムのテスト
    console.log('\n1. URL管理システムテスト');
    var urlTest = testUrlManagement();
    testResults.push({
      testName: 'URL管理システム',
      result: urlTest
    });
    
    // 2. 付与ルール計算のテスト
    console.log('\n2. 付与ルール計算テスト');
    var grantTest = testLeaveGrantCalculation();
    testResults.push({
      testName: '付与ルール計算',
      result: grantTest
    });
    
    // 3. 付与履歴管理のテスト
    console.log('\n3. 付与履歴管理テスト');
    var historyTest = testGrantHistoryManagement();
    testResults.push({
      testName: '付与履歴管理',
      result: historyTest
    });
    
    // 4. FIFO消費のテスト
    console.log('\n4. FIFO消費テスト');
    var fifoTest = testFifoConsumption();
    testResults.push({
      testName: 'FIFO消費',
      result: fifoTest
    });
    
    // 5. 失効処理のテスト
    console.log('\n5. 失効処理テスト');
    var expireTest = testExpiryProcess();
    testResults.push({
      testName: '失効処理',
      result: expireTest
    });
    
    // 6. 通知システムのテスト
    console.log('\n6. 通知システムテスト');
    var notificationTest = testNotificationSystem();
    testResults.push({
      testName: '通知システム',
      result: notificationTest
    });
    
    // 7. 年次付与処理のテスト
    console.log('\n7. 年次付与処理テスト');
    var annualGrantTest = testAnnualGrantProcess();
    testResults.push({
      testName: '年次付与処理',
      result: annualGrantTest
    });
    
    // 8. トリガー管理システムのテスト
    console.log('\n8. トリガー管理システムテスト');
    var triggerTest = testTriggerManagement();
    testResults.push({
      testName: 'トリガー管理システム',
      result: triggerTest
    });
    
    // 9. 統計レポート機能のテスト
    console.log('\n9. 統計レポート機能テスト');
    var statisticsTest = testStatisticsReportSystem();
    testResults.push({
      testName: '統計レポート機能',
      result: statisticsTest
    });

    // 10. エッジケース・境界値テスト
    console.log('\n10. エッジケース・境界値テスト');
    var edgeCaseTest = runEdgeCaseTests();
    testResults.push({
      testName: 'エッジケース・境界値',
      result: edgeCaseTest
    });

    // 11. データ整合性詳細確認テスト
    console.log('\n11. データ整合性詳細確認テスト');
    var dataIntegrityTest = runDataIntegrityTests();
    testResults.push({
      testName: 'データ整合性詳細確認',
      result: dataIntegrityTest
    });

  } catch (error) {
    console.error('テスト実行中にエラー:', error);
    testResults.push({
      testName: 'テスト実行',
      result: {
        success: false,
        message: 'テスト実行中にエラーが発生しました: ' + error.message
      }
    });
  }
  
  // 結果サマリーを表示
  console.log('\n=== テスト結果サマリー ===');
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
 * 付与履歴管理システムのテスト
 */
function testGrantHistoryManagement() {
  try {
    console.log('=== 付与履歴管理テスト開始 ===');
    
    var testUserId = 'TEST001';
    var results = [];
    
    // テスト1: 付与履歴シートの作成
    console.log('1. 付与履歴シートの作成テスト');
    var ss = getSpreadsheet();
    var grantHistorySheet = getOrCreateGrantHistorySheet(ss);
    
    results.push({
      test: '付与履歴シート作成',
      result: grantHistorySheet ? 'SUCCESS' : 'FAILED',
      detail: grantHistorySheet ? 'シート作成成功' : 'シート作成失敗'
    });
    
    // テスト2: 有給付与のテスト
    console.log('2. 有給付与テスト');
    var grantDate = new Date('2024-08-18');
    var grantResult = grantLeave(testUserId, grantDate, 10, '初回', 0.5);
    
    results.push({
      test: '有給付与',
      result: grantResult.success ? 'SUCCESS' : 'FAILED',
      detail: grantResult.message
    });
    
    // テスト3: 付与履歴取得のテスト
    console.log('3. 付与履歴取得テスト');
    var history = getUserGrantHistory(testUserId);
    
    results.push({
      test: '付与履歴取得',
      result: history.length > 0 ? 'SUCCESS' : 'FAILED',
      detail: '取得件数: ' + history.length + '件'
    });
    
    // テスト4: 残日数計算のテスト
    console.log('4. 残日数計算テスト');
    var effectiveRemaining = calculateEffectiveRemainingDays(testUserId);
    
    results.push({
      test: '残日数計算',
      result: effectiveRemaining === 10 ? 'SUCCESS' : 'FAILED',
      detail: '計算結果: ' + effectiveRemaining + '日'
    });
    
    // 結果集計
    console.log('=== 付与履歴管理テスト結果 ===');
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
    console.error('付与履歴管理テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * FIFO消費システムのテスト
 */
function testFifoConsumption() {
  try {
    console.log('=== FIFO消費テスト開始 ===');
    
    var testUserId = 'TEST002';
    var results = [];
    
    // 準備: 複数回の付与を行う
    console.log('準備: 複数回の付与');
    var grant1 = grantLeave(testUserId, new Date('2023-08-18'), 10, '初回', 0.5);
    var grant2 = grantLeave(testUserId, new Date('2024-04-01'), 11, '年次', 1.5);
    
    if (!grant1.success || !grant2.success) {
      throw new Error('付与準備に失敗しました');
    }
    
    // テスト1: 部分消費（古い付与分から消費されるか）
    console.log('1. 部分消費テスト（5日消費）');
    var consume1 = consumeLeave(testUserId, 5);
    
    results.push({
      test: '部分消費（5日）',
      result: consume1.success ? 'SUCCESS' : 'FAILED',
      detail: consume1.message
    });
    
    // テスト2: 跨ぎ消費（複数の付与分に跨がる消費）
    console.log('2. 跨ぎ消費テスト（8日消費）');
    var consume2 = consumeLeave(testUserId, 8);
    
    results.push({
      test: '跨ぎ消費（8日）',
      result: consume2.success ? 'SUCCESS' : 'FAILED',
      detail: consume2.message
    });
    
    // テスト3: 残日数確認
    console.log('3. 残日数確認テスト');
    var remaining = calculateEffectiveRemainingDays(testUserId);
    var expectedRemaining = 21 - 5 - 8; // 8日残るはず
    
    results.push({
      test: '残日数確認',
      result: remaining === expectedRemaining ? 'SUCCESS' : 'FAILED',
      detail: '残日数: ' + remaining + '日（期待値: ' + expectedRemaining + '日）'
    });
    
    // テスト4: 過剰消費エラー
    console.log('4. 過剰消費エラーテスト');
    var consume3 = consumeLeave(testUserId, 100); // 100日消費（エラーになるはず）
    
    results.push({
      test: '過剰消費エラー',
      result: !consume3.success ? 'SUCCESS' : 'FAILED',
      detail: consume3.message
    });
    
    // 結果集計
    console.log('=== FIFO消費テスト結果 ===');
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
    console.error('FIFO消費テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 失効処理のテスト
 */
function testExpiryProcess() {
  try {
    console.log('=== 失効処理テスト開始 ===');
    
    var testUserId = 'TEST003';
    var results = [];
    
    // 準備: 失効対象の付与を作成（過去の日付で付与）
    console.log('準備: 失効対象データ作成');
    var expiredDate = new Date('2022-08-18'); // 2年以上前
    var grant = grantLeave(testUserId, expiredDate, 5, '初回', 0.5);
    
    if (!grant.success) {
      throw new Error('失効テスト用データ作成に失敗しました');
    }
    
    // 付与履歴を手動で失効日を過去に設定
    var ss = getSpreadsheet();
    var grantHistorySheet = getOrCreateGrantHistorySheet(ss);
    var data = grantHistorySheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === testUserId) {
        // 失効日を昨日に設定
        var yesterday = new Date();
        yesterday.setDate(yesterday.getDate() - 1);
        grantHistorySheet.getRange(i + 1, 4).setValue(yesterday);
        break;
      }
    }
    
    // テスト1: 失効処理の実行
    console.log('1. 失効処理実行テスト');
    var expireResult = processExpiredLeaves();
    
    results.push({
      test: '失効処理実行',
      result: expireResult.success ? 'SUCCESS' : 'FAILED',
      detail: expireResult.message
    });
    
    // テスト2: 失効後の残日数確認
    console.log('2. 失効後残日数確認テスト');
    var remainingAfterExpiry = calculateEffectiveRemainingDays(testUserId);
    
    results.push({
      test: '失効後残日数確認',
      result: remainingAfterExpiry === 0 ? 'SUCCESS' : 'FAILED',
      detail: '失効後残日数: ' + remainingAfterExpiry + '日'
    });
    
    // 結果集計
    console.log('=== 失効処理テスト結果 ===');
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
    console.error('失効処理テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 統合テスト用のテストデータクリーンアップ
 */
function cleanupTestData() {
  try {
    console.log('=== テストデータクリーンアップ開始 ===');
    
    var ss = getSpreadsheet();
    
    // 付与履歴からテストデータを削除
    var grantHistorySheet = ss.getSheetByName('付与履歴');
    if (grantHistorySheet) {
      var data = grantHistorySheet.getDataRange().getValues();
      var rowsToDelete = [];
      
      for (var i = 1; i < data.length; i++) {
        var userId = String(data[i][0]);
        if (userId.startsWith('TEST')) {
          rowsToDelete.push(i + 1);
        }
      }
      
      // 後ろから削除（行番号が変わらないようにするため）
      for (var i = rowsToDelete.length - 1; i >= 0; i--) {
        grantHistorySheet.deleteRow(rowsToDelete[i]);
      }
      
      console.log('付与履歴からテストデータを削除:', rowsToDelete.length + '行');
    }
    
    // URL管理からテストデータを削除
    var urlSheet = ss.getSheetByName('URL管理');
    if (urlSheet) {
      var urlData = urlSheet.getDataRange().getValues();
      var urlRowsToDelete = [];
      
      for (var i = 1; i < urlData.length; i++) {
        var userId = String(urlData[i][0]);
        if (userId.startsWith('TEST')) {
          urlRowsToDelete.push(i + 1);
        }
      }
      
      // 後ろから削除
      for (var i = urlRowsToDelete.length - 1; i >= 0; i--) {
        urlSheet.deleteRow(urlRowsToDelete[i]);
      }
      
      console.log('URL管理からテストデータを削除:', urlRowsToDelete.length + '行');
    }
    
    console.log('テストデータクリーンアップ完了');
    
    return {
      success: true,
      message: 'テストデータをクリーンアップしました'
    };
    
  } catch (error) {
    console.error('テストデータクリーンアップエラー:', error);
    return {
      success: false,
      message: 'クリーンアップに失敗しました: ' + error.message
    };
  }
}

/**
 * 年次付与処理のテスト
 */
function testAnnualGrantProcess() {
  try {
    console.log('=== 年次付与処理テスト開始 ===');
    
    var results = [];
    
    // テスト1: 年次付与対象者抽出テスト
    console.log('1. 年次付与対象者抽出テスト');
    var targets = getAnnualGrantTargets();
    
    results.push({
      test: '年次付与対象者抽出',
      result: 'SUCCESS',
      detail: '対象者数: ' + targets.length + '名'
    });
    
    // テスト2: 付与日数計算テスト（年次）
    console.log('2. 年次付与日数計算テスト');
    var testCases = [
      { workYears: 1.5, weekDays: 5, expected: 11 },
      { workYears: 2.5, weekDays: 5, expected: 12 },
      { workYears: 3.5, weekDays: 5, expected: 14 },
      { workYears: 6.5, weekDays: 5, expected: 20 }
    ];
    
    var calculationSuccess = true;
    testCases.forEach(function(testCase) {
      var result = calculateLeaveDays(testCase.workYears, testCase.weekDays, false);
      if (result !== testCase.expected) {
        calculationSuccess = false;
      }
    });
    
    results.push({
      test: '年次付与日数計算',
      result: calculationSuccess ? 'SUCCESS' : 'FAILED',
      detail: '全' + testCases.length + 'ケース' + (calculationSuccess ? '成功' : '失敗')
    });
    
    // テスト3: 出勤率チェック機能テスト
    console.log('3. 出勤率チェック機能テスト');
    var testUserId = 'TEST_ANNUAL';
    var testFromDate = new Date('2023-04-01');
    var testToDate = new Date('2024-03-31');
    
    var attendanceRate = checkAnnualAttendanceRate(testUserId, testFromDate, testToDate);
    
    results.push({
      test: '出勤率チェック機能',
      result: attendanceRate >= 0 && attendanceRate <= 1 ? 'SUCCESS' : 'FAILED',
      detail: '出勤率: ' + Math.round(attendanceRate * 100) + '%'
    });
    
    // テスト4: マスターシート設定確認
    console.log('4. マスターシート設定確認');
    var setupResult = setupMasterSheetColumns();
    
    results.push({
      test: 'マスターシート設定',
      result: setupResult.success ? 'SUCCESS' : 'FAILED',
      detail: setupResult.message
    });
    
    // 結果集計
    console.log('=== 年次付与処理テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功',
      message: '年次付与処理の基盤機能をテストしました'
    };
    
  } catch (error) {
    console.error('年次付与処理テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * トリガー管理システムのテスト
 */
function testTriggerManagement() {
  try {
    console.log('=== トリガー管理システムテスト開始 ===');
    
    var results = [];
    
    // テスト1: 現在のトリガー状況確認
    console.log('1. 現在のトリガー状況確認');
    var statusResult = getTriggerStatus();
    
    results.push({
      test: 'トリガー状況確認',
      result: statusResult.success ? 'SUCCESS' : 'FAILED',
      detail: statusResult.message
    });
    
    // テスト2: データ整合性チェック機能
    console.log('2. データ整合性チェック機能テスト');
    var integrityResult = checkDataIntegrity();
    
    results.push({
      test: 'データ整合性チェック',
      result: integrityResult.success ? 'SUCCESS' : 'FAILED',
      detail: integrityResult.message
    });
    
    // テスト3: 残日数アラートチェック機能
    console.log('3. 残日数アラートチェック機能テスト');
    var alertResult = checkLowRemainingDaysAlerts();
    
    results.push({
      test: '残日数アラートチェック',
      result: alertResult.success ? 'SUCCESS' : 'FAILED',
      detail: alertResult.message
    });
    
    // テスト4: 年度切り替え処理テスト
    console.log('4. 年度切り替え処理テスト');
    var fiscalYearResult = processFiscalYearChange();
    
    results.push({
      test: '年度切り替え処理',
      result: fiscalYearResult.success ? 'SUCCESS' : 'FAILED',
      detail: fiscalYearResult.message
    });
    
    // 結果集計
    console.log('=== トリガー管理システムテスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功',
      message: 'トリガー管理システムの基盤機能をテストしました'
    };
    
  } catch (error) {
    console.error('トリガー管理システムテストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 統計レポート機能システムのテスト
 */
function testStatisticsReportSystem() {
  try {
    console.log('=== 統計レポート機能システムテスト開始 ===');
    
    var results = [];
    
    // テスト1: 基本統計計算機能
    console.log('1. 基本統計計算機能テスト');
    var basicStatsResult = testBasicStatisticsCalculation();
    
    results.push({
      test: '基本統計計算機能',
      result: basicStatsResult.success ? 'SUCCESS' : 'FAILED',
      detail: basicStatsResult.summary
    });
    
    // テスト2: 月次レポート生成機能
    console.log('2. 月次レポート生成機能テスト');
    var monthlyReportResult = generateMonthlyReport(2024, 8);
    
    results.push({
      test: '月次レポート生成機能',
      result: monthlyReportResult.success ? 'SUCCESS' : 'FAILED',
      detail: monthlyReportResult.message
    });
    
    // テスト3: HTMLレポート生成機能
    console.log('3. HTMLレポート生成機能テスト');
    if (monthlyReportResult.success) {
      var htmlReport = generateReportHTML(monthlyReportResult.reportData);
      var htmlValid = typeof htmlReport === 'string' && htmlReport.length > 0;
      
      results.push({
        test: 'HTMLレポート生成機能',
        result: htmlValid ? 'SUCCESS' : 'FAILED',
        detail: 'HTML生成: ' + (htmlValid ? '成功' : '失敗') + ' (長さ: ' + htmlReport.length + '文字)'
      });
    } else {
      results.push({
        test: 'HTMLレポート生成機能',
        result: 'FAILED',
        detail: '月次レポート生成失敗により実行不可'
      });
    }
    
    // テスト4: 年5日取得義務統計
    console.log('4. 年5日取得義務統計テスト');
    var fiveDayStats = calculateFiveDayObligationStats(new Date('2024-04-01'), new Date('2025-03-31'));
    var fiveDayValid = fiveDayStats.hasOwnProperty('targetEmployees') && 
                      fiveDayStats.hasOwnProperty('complianceRate');
    
    results.push({
      test: '年5日取得義務統計',
      result: fiveDayValid ? 'SUCCESS' : 'FAILED',
      detail: '対象者: ' + fiveDayStats.targetEmployees + '名, 達成率: ' + fiveDayStats.complianceRate + '%'
    });
    
    // テスト5: 事業所別統計
    console.log('5. 事業所別統計テスト');
    var divisionStats = calculateDivisionStatistics(new Date('2024-08-01'), new Date('2024-09-01'));
    var divisionValid = divisionStats.hasOwnProperty('R') && 
                       divisionStats.hasOwnProperty('P') &&
                       divisionStats.hasOwnProperty('S') && 
                       divisionStats.hasOwnProperty('E');
    
    results.push({
      test: '事業所別統計',
      result: divisionValid ? 'SUCCESS' : 'FAILED',
      detail: '事業所数: ' + Object.keys(divisionStats).length + '箇所'
    });
    
    // 結果集計
    console.log('=== 統計レポート機能システムテスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功',
      message: '統計レポート機能システムをテストしました'
    };
    
  } catch (error) {
    console.error('統計レポート機能システムテストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 開発進捗の確認
 */
function checkDevelopmentProgress() {
  console.log('=== 開発進捗確認 ===');
  
  var completedFeatures = [
    '✅ URL管理システム',
    '✅ 個人専用ページ',
    '✅ アクセス制御機能',
    '✅ 付与履歴管理',
    '✅ FIFO消費システム',
    '✅ 失効処理',
    '✅ 付与ルール計算',
    '✅ 通知システム基盤',
    '✅ 6ヶ月付与処理',
    '✅ 年次付与処理',
    '✅ GASトリガー管理システム',
    '✅ 通知機能の本格運用',
    '✅ 統計レポート機能',
    '✅ エッジケース・境界値テスト',
    '✅ データ整合性詳細確認テスト'
  ];
  
  var pendingFeatures = [
    '🚧 年5日義務の監視強化',
    '🚧 管理画面の拡張',
    '🚧 エラーログ管理',
    '🚧 データエクスポート機能',
    '🚧 高度な分析機能'
  ];
  
  console.log('\n完了済み機能:');
  completedFeatures.forEach(function(feature) {
    console.log('  ' + feature);
  });
  
  console.log('\n実装予定機能:');
  pendingFeatures.forEach(function(feature) {
    console.log('  ' + feature);
  });
  
  var totalFeatures = completedFeatures.length + pendingFeatures.length;
  var completedCount = completedFeatures.length;
  var progressRate = Math.round((completedCount / totalFeatures) * 100);
  
  console.log('\n進捗率: ' + progressRate + '% (' + completedCount + '/' + totalFeatures + ')');
  
  return {
    progressRate: progressRate,
    completedFeatures: completedFeatures,
    pendingFeatures: pendingFeatures
  };
}