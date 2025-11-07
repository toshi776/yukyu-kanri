// =============================
// 年次付与処理テスト専用スクリプト
// =============================

/**
 * 年次付与処理の単体テスト
 */
function runAnnualGrantTests() {
  console.log('=== 年次付与処理テスト開始 ===');
  
  var testResults = [];
  
  try {
    // テスト1: 付与日数計算の確認
    var calculationTest = testAnnualGrantCalculation();
    testResults.push({
      testName: '年次付与日数計算',
      result: calculationTest
    });
    
    // テスト2: 対象者抽出の確認
    var targetTest = testAnnualTargetExtraction();
    testResults.push({
      testName: '年次付与対象者抽出',
      result: targetTest
    });
    
    // テスト3: 出勤率計算の確認
    var attendanceTest = testAttendanceRateCalculation();
    testResults.push({
      testName: '出勤率計算',
      result: attendanceTest
    });
    
    // テスト4: マスターシート設定の確認
    var setupTest = testMasterSheetSetup();
    testResults.push({
      testName: 'マスターシート設定',
      result: setupTest
    });
    
  } catch (error) {
    console.error('年次付与テスト実行エラー:', error);
    testResults.push({
      testName: 'テスト実行',
      result: {
        success: false,
        message: 'テスト実行中にエラーが発生しました: ' + error.message
      }
    });
  }
  
  // 結果サマリーを表示
  console.log('\n=== 年次付与テスト結果サマリー ===');
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
 * 年次付与日数計算のテスト
 */
function testAnnualGrantCalculation() {
  try {
    console.log('年次付与日数計算テスト開始');
    
    var testCases = [
      // 週5日勤務者のテスト
      { workYears: 1.0, weekDays: 5, expected: 11, desc: '1年・週5日' },
      { workYears: 1.5, weekDays: 5, expected: 11, desc: '1.5年・週5日' },
      { workYears: 2.0, weekDays: 5, expected: 12, desc: '2年・週5日' },
      { workYears: 3.0, weekDays: 5, expected: 14, desc: '3年・週5日' },
      { workYears: 4.0, weekDays: 5, expected: 16, desc: '4年・週5日' },
      { workYears: 5.0, weekDays: 5, expected: 18, desc: '5年・週5日' },
      { workYears: 6.0, weekDays: 5, expected: 20, desc: '6年・週5日' },
      
      // 週4日勤務者のテスト
      { workYears: 1.0, weekDays: 4, expected: 8, desc: '1年・週4日' },
      { workYears: 6.0, weekDays: 4, expected: 15, desc: '6年・週4日' },
      
      // 週3日勤務者のテスト
      { workYears: 1.0, weekDays: 3, expected: 6, desc: '1年・週3日' },
      { workYears: 6.0, weekDays: 3, expected: 11, desc: '6年・週3日' }
    ];
    
    var failedTests = [];
    var passedCount = 0;
    
    testCases.forEach(function(testCase) {
      var result = calculateLeaveDays(testCase.workYears, testCase.weekDays, false);
      
      if (result === testCase.expected) {
        passedCount++;
        console.log('✅ ' + testCase.desc + ': ' + result + '日 (期待値: ' + testCase.expected + '日)');
      } else {
        failedTests.push({
          case: testCase.desc,
          actual: result,
          expected: testCase.expected
        });
        console.log('❌ ' + testCase.desc + ': ' + result + '日 (期待値: ' + testCase.expected + '日)');
      }
    });
    
    if (failedTests.length === 0) {
      return {
        success: true,
        message: '全' + testCases.length + 'ケースが正常に計算されました',
        passedCount: passedCount,
        totalCount: testCases.length
      };
    } else {
      return {
        success: false,
        message: failedTests.length + '/' + testCases.length + 'ケースが失敗しました',
        failedTests: failedTests,
        passedCount: passedCount,
        totalCount: testCases.length
      };
    }
    
  } catch (error) {
    console.error('年次付与日数計算テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 年次付与対象者抽出のテスト
 */
function testAnnualTargetExtraction() {
  try {
    console.log('年次付与対象者抽出テスト開始');
    
    // 対象者抽出関数をテスト
    var targets = getAnnualGrantTargets();
    
    console.log('抽出された対象者数:', targets.length + '名');
    
    // 各対象者の情報を確認
    targets.forEach(function(target, index) {
      console.log((index + 1) + '. ' + target.userId + ' (' + target.name + ') - 勤続' + Math.floor(target.workYears) + '年');
    });
    
    return {
      success: true,
      message: targets.length + '名の年次付与対象者を抽出しました',
      targetCount: targets.length,
      targets: targets
    };
    
  } catch (error) {
    console.error('年次付与対象者抽出テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 出勤率計算のテスト
 */
function testAttendanceRateCalculation() {
  try {
    console.log('出勤率計算テスト開始');
    
    // テスト用の期間設定
    var testFromDate = new Date('2023-04-01');
    var testToDate = new Date('2024-03-31');
    var testUserId = 'TEST_ATTENDANCE';
    
    // 出勤率計算を実行
    var attendanceRate = checkAnnualAttendanceRate(testUserId, testFromDate, testToDate);
    
    console.log('テスト結果:');
    console.log('- 対象期間: ' + Utilities.formatDate(testFromDate, 'JST', 'yyyy/MM/dd') + ' - ' + Utilities.formatDate(testToDate, 'JST', 'yyyy/MM/dd'));
    console.log('- 利用者ID: ' + testUserId);
    console.log('- 出勤率: ' + Math.round(attendanceRate * 100) + '%');
    
    // 出勤率が0-1の範囲内かチェック
    var isValidRange = attendanceRate >= 0 && attendanceRate <= 1;
    
    if (isValidRange) {
      return {
        success: true,
        message: '出勤率計算が正常に実行されました (' + Math.round(attendanceRate * 100) + '%)',
        attendanceRate: attendanceRate
      };
    } else {
      return {
        success: false,
        message: '出勤率が不正な値です: ' + attendanceRate,
        attendanceRate: attendanceRate
      };
    }
    
  } catch (error) {
    console.error('出勤率計算テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * マスターシート設定のテスト
 */
function testMasterSheetSetup() {
  try {
    console.log('マスターシート設定テスト開始');
    
    // マスターシート設定関数をテスト
    var setupResult = setupMasterSheetColumns();
    
    if (setupResult.success) {
      console.log('✅ マスターシート設定成功:', setupResult.message);
      
      // 実際のシート構造を確認
      var ss = getSpreadsheet();
      var masterSheet = ss.getSheetByName('マスター');
      
      if (masterSheet) {
        var lastColumn = masterSheet.getLastColumn();
        var headerRow = masterSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
        
        console.log('現在の列数:', lastColumn);
        console.log('ヘッダー行:', headerRow);
        
        return {
          success: true,
          message: 'マスターシート設定完了 (' + lastColumn + '列)',
          columns: lastColumn,
          headers: headerRow
        };
      } else {
        return {
          success: false,
          message: 'マスターシートが見つかりません'
        };
      }
    } else {
      return setupResult;
    }
    
  } catch (error) {
    console.error('マスターシート設定テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 年次付与処理の統合テスト（実際の付与は行わない）
 */
function testAnnualGrantProcessDryRun() {
  try {
    console.log('=== 年次付与処理統合テスト（ドライラン） ===');
    
    // 対象者抽出
    var targets = getAnnualGrantTargets();
    
    if (targets.length === 0) {
      return {
        success: true,
        message: '現在年次付与対象者はいません',
        processedCount: 0
      };
    }
    
    console.log('年次付与対象者:', targets.length + '名');
    
    var simulationResults = [];
    
    // 各対象者をシミュレーション
    targets.forEach(function(target) {
      // 出勤率チェック
      var attendanceRate = checkAnnualAttendanceRate(
        target.userId, 
        target.initialGrantDate, 
        new Date()
      );
      
      // 付与日数計算
      var grantDays = calculateLeaveDays(target.workYears, target.weeklyWorkDays, false);
      
      var willGrant = attendanceRate >= 0.8;
      
      simulationResults.push({
        userId: target.userId,
        name: target.name,
        workYears: Math.floor(target.workYears),
        attendanceRate: Math.round(attendanceRate * 100) + '%',
        grantDays: grantDays,
        willGrant: willGrant,
        reason: willGrant ? '付与予定' : '出勤率不足'
      });
      
      console.log('- ' + target.userId + ' (' + target.name + '): ' + 
                  (willGrant ? grantDays + '日付与予定' : '出勤率不足により見送り'));
    });
    
    var grantCount = simulationResults.filter(function(r) { return r.willGrant; }).length;
    
    return {
      success: true,
      message: '年次付与シミュレーション完了 (' + grantCount + '/' + targets.length + '名が付与対象)',
      totalTargets: targets.length,
      grantCount: grantCount,
      simulationResults: simulationResults
    };
    
  } catch (error) {
    console.error('年次付与統合テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}