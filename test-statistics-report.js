// =============================
// 統計レポート機能テスト
// =============================

/**
 * 統計レポート機能の全テストを実行
 */
function runStatisticsReportTests() {
  console.log('=== 統計レポート機能テスト開始 ===');
  
  var testResults = [];
  
  try {
    // テスト1: 基本統計計算のテスト
    var basicStatsTest = testBasicStatisticsCalculation();
    testResults.push({
      testName: '基本統計計算',
      result: basicStatsTest
    });
    
    // テスト2: 申請統計計算のテスト
    var applicationStatsTest = testApplicationStatisticsCalculation();
    testResults.push({
      testName: '申請統計計算',
      result: applicationStatsTest
    });
    
    // テスト3: 付与統計計算のテスト
    var grantStatsTest = testGrantStatisticsCalculation();
    testResults.push({
      testName: '付与統計計算',
      result: grantStatsTest
    });
    
    // テスト4: 年5日取得義務統計のテスト
    var fiveDayObligationTest = testFiveDayObligationStatistics();
    testResults.push({
      testName: '年5日取得義務統計',
      result: fiveDayObligationTest
    });
    
    // テスト5: 月次レポート生成のテスト
    var monthlyReportTest = testMonthlyReportGeneration();
    testResults.push({
      testName: '月次レポート生成',
      result: monthlyReportTest
    });
    
    // テスト6: HTMLレポート生成のテスト
    var htmlReportTest = testHTMLReportGeneration();
    testResults.push({
      testName: 'HTMLレポート生成',
      result: htmlReportTest
    });
    
  } catch (error) {
    console.error('統計レポートテスト実行エラー:', error);
    testResults.push({
      testName: 'テスト実行',
      result: {
        success: false,
        message: 'テスト実行中にエラーが発生しました: ' + error.message
      }
    });
  }
  
  // 結果サマリーを表示
  console.log('\n=== 統計レポートテスト結果サマリー ===');
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
 * 基本統計計算のテスト
 */
function testBasicStatisticsCalculation() {
  try {
    console.log('基本統計計算テスト開始');
    
    var testDate = new Date('2024-08-01');
    var basicStats = calculateBasicStatistics(testDate);
    
    var results = [];
    
    // テスト1: 統計データの構造確認
    var hasRequiredFields = basicStats.hasOwnProperty('totalEmployees') &&
                           basicStats.hasOwnProperty('totalRemainingDays') &&
                           basicStats.hasOwnProperty('averageRemainingDays') &&
                           basicStats.hasOwnProperty('lowRemainingCount') &&
                           basicStats.hasOwnProperty('divisionCounts');
    
    results.push({
      test: '統計データ構造確認',
      result: hasRequiredFields ? 'SUCCESS' : 'FAILED',
      detail: '必要フィールド: ' + (hasRequiredFields ? '完備' : '不足')
    });
    
    // テスト2: 数値データの妥当性確認
    var isValidNumbers = typeof basicStats.totalEmployees === 'number' &&
                        typeof basicStats.averageRemainingDays === 'number' &&
                        basicStats.totalEmployees >= 0 &&
                        basicStats.averageRemainingDays >= 0;
    
    results.push({
      test: '数値データ妥当性',
      result: isValidNumbers ? 'SUCCESS' : 'FAILED',
      detail: '総従業員数: ' + basicStats.totalEmployees + '名, 平均残日数: ' + basicStats.averageRemainingDays + '日'
    });
    
    // テスト3: 事業所別カウントの確認
    var divisionCounts = basicStats.divisionCounts;
    var hasAllDivisions = divisionCounts.hasOwnProperty('R') &&
                         divisionCounts.hasOwnProperty('P') &&
                         divisionCounts.hasOwnProperty('S') &&
                         divisionCounts.hasOwnProperty('E');
    
    results.push({
      test: '事業所別カウント',
      result: hasAllDivisions ? 'SUCCESS' : 'FAILED',
      detail: 'R:' + divisionCounts.R + ', P:' + divisionCounts.P + ', S:' + divisionCounts.S + ', E:' + divisionCounts.E
    });
    
    // 結果集計
    console.log('=== 基本統計計算テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功',
      basicStats: basicStats
    };
    
  } catch (error) {
    console.error('基本統計計算テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 申請統計計算のテスト
 */
function testApplicationStatisticsCalculation() {
  try {
    console.log('申請統計計算テスト開始');
    
    var ss = getSpreadsheet();
    var applySheet = ss.getSheetByName('申請');
    var startDate = new Date('2024-08-01');
    var endDate = new Date('2024-09-01');
    
    var applicationStats = calculateApplicationStatistics(applySheet, startDate, endDate);
    
    var results = [];
    
    // テスト1: 申請統計データの構造確認
    var hasRequiredFields = applicationStats.hasOwnProperty('totalApplications') &&
                           applicationStats.hasOwnProperty('approvedApplications') &&
                           applicationStats.hasOwnProperty('rejectedApplications') &&
                           applicationStats.hasOwnProperty('totalAppliedDays') &&
                           applicationStats.hasOwnProperty('approvalRate');
    
    results.push({
      test: '申請統計構造確認',
      result: hasRequiredFields ? 'SUCCESS' : 'FAILED',
      detail: '必要フィールド: ' + (hasRequiredFields ? '完備' : '不足')
    });
    
    // テスト2: 数値データの妥当性確認
    var isValidNumbers = typeof applicationStats.totalApplications === 'number' &&
                        typeof applicationStats.approvalRate === 'number' &&
                        applicationStats.totalApplications >= 0 &&
                        applicationStats.approvalRate >= 0 &&
                        applicationStats.approvalRate <= 100;
    
    results.push({
      test: '申請統計数値妥当性',
      result: isValidNumbers ? 'SUCCESS' : 'FAILED',
      detail: '総申請: ' + applicationStats.totalApplications + '件, 承認率: ' + applicationStats.approvalRate + '%'
    });
    
    // テスト3: 申請合計の整合性確認
    var totalCheck = applicationStats.totalApplications === 
                    (applicationStats.approvedApplications + applicationStats.rejectedApplications + applicationStats.pendingApplications);
    
    results.push({
      test: '申請合計整合性',
      result: totalCheck ? 'SUCCESS' : 'FAILED',
      detail: '総申請数と内訳の合計: ' + (totalCheck ? '一致' : '不一致')
    });
    
    // 結果集計
    console.log('=== 申請統計計算テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功',
      applicationStats: applicationStats
    };
    
  } catch (error) {
    console.error('申請統計計算テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 付与統計計算のテスト
 */
function testGrantStatisticsCalculation() {
  try {
    console.log('付与統計計算テスト開始');
    
    var ss = getSpreadsheet();
    var grantHistorySheet = ss.getSheetByName('付与履歴');
    var startDate = new Date('2024-08-01');
    var endDate = new Date('2024-09-01');
    
    var grantStats = calculateGrantStatistics(grantHistorySheet, startDate, endDate);
    
    var results = [];
    
    // テスト1: 付与統計データの構造確認
    var hasRequiredFields = grantStats.hasOwnProperty('totalGrants') &&
                           grantStats.hasOwnProperty('totalGrantedDays') &&
                           grantStats.hasOwnProperty('initialGrants') &&
                           grantStats.hasOwnProperty('annualGrants') &&
                           grantStats.hasOwnProperty('averageGrantDays');
    
    results.push({
      test: '付与統計構造確認',
      result: hasRequiredFields ? 'SUCCESS' : 'FAILED',
      detail: '必要フィールド: ' + (hasRequiredFields ? '完備' : '不足')
    });
    
    // テスト2: 数値データの妥当性確認
    var isValidNumbers = typeof grantStats.totalGrants === 'number' &&
                        typeof grantStats.averageGrantDays === 'number' &&
                        grantStats.totalGrants >= 0 &&
                        grantStats.averageGrantDays >= 0;
    
    results.push({
      test: '付与統計数値妥当性',
      result: isValidNumbers ? 'SUCCESS' : 'FAILED',
      detail: '総付与数: ' + grantStats.totalGrants + '件, 平均付与日数: ' + grantStats.averageGrantDays + '日'
    });
    
    // テスト3: 付与タイプ合計の整合性確認
    var typeCheck = grantStats.totalGrants === (grantStats.initialGrants + grantStats.annualGrants);
    
    results.push({
      test: '付与タイプ合計整合性',
      result: typeCheck ? 'SUCCESS' : 'FAILED',
      detail: '初回: ' + grantStats.initialGrants + ', 年次: ' + grantStats.annualGrants + ', 合計: ' + grantStats.totalGrants
    });
    
    // 結果集計
    console.log('=== 付与統計計算テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功',
      grantStats: grantStats
    };
    
  } catch (error) {
    console.error('付与統計計算テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 年5日取得義務統計のテスト
 */
function testFiveDayObligationStatistics() {
  try {
    console.log('年5日取得義務統計テスト開始');
    
    var startDate = new Date('2024-04-01');
    var endDate = new Date('2025-03-31');
    
    var fiveDayStats = calculateFiveDayObligationStats(startDate, endDate);
    
    var results = [];
    
    // テスト1: 年5日義務統計データの構造確認
    var hasRequiredFields = fiveDayStats.hasOwnProperty('targetEmployees') &&
                           fiveDayStats.hasOwnProperty('compliantEmployees') &&
                           fiveDayStats.hasOwnProperty('complianceRate') &&
                           fiveDayStats.hasOwnProperty('details');
    
    results.push({
      test: '年5日義務統計構造確認',
      result: hasRequiredFields ? 'SUCCESS' : 'FAILED',
      detail: '必要フィールド: ' + (hasRequiredFields ? '完備' : '不足')
    });
    
    // テスト2: 達成率計算の妥当性確認
    var isValidRate = typeof fiveDayStats.complianceRate === 'number' &&
                     fiveDayStats.complianceRate >= 0 &&
                     fiveDayStats.complianceRate <= 100;
    
    results.push({
      test: '達成率計算妥当性',
      result: isValidRate ? 'SUCCESS' : 'FAILED',
      detail: '対象者: ' + fiveDayStats.targetEmployees + '名, 達成率: ' + fiveDayStats.complianceRate + '%'
    });
    
    // テスト3: 詳細データの形式確認
    var detailsValid = Array.isArray(fiveDayStats.details);
    var sampleDetail = fiveDayStats.details.length > 0 ? fiveDayStats.details[0] : null;
    var detailStructure = sampleDetail && 
                         sampleDetail.hasOwnProperty('userId') &&
                         sampleDetail.hasOwnProperty('takenDays') &&
                         sampleDetail.hasOwnProperty('isCompliant');
    
    results.push({
      test: '詳細データ形式',
      result: (detailsValid && (fiveDayStats.details.length === 0 || detailStructure)) ? 'SUCCESS' : 'FAILED',
      detail: '詳細データ件数: ' + fiveDayStats.details.length + '件'
    });
    
    // 結果集計
    console.log('=== 年5日取得義務統計テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功',
      fiveDayStats: fiveDayStats
    };
    
  } catch (error) {
    console.error('年5日取得義務統計テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 月次レポート生成のテスト
 */
function testMonthlyReportGeneration() {
  try {
    console.log('月次レポート生成テスト開始');
    
    var testYear = 2024;
    var testMonth = 8;
    
    var monthlyReport = generateMonthlyReport(testYear, testMonth);
    
    var results = [];
    
    // テスト1: レポート生成成功確認
    results.push({
      test: 'レポート生成成功',
      result: monthlyReport.success ? 'SUCCESS' : 'FAILED',
      detail: monthlyReport.message
    });
    
    if (monthlyReport.success) {
      var reportData = monthlyReport.reportData;
      
      // テスト2: レポートデータ構造確認
      var hasReportStructure = reportData.hasOwnProperty('reportType') &&
                              reportData.hasOwnProperty('period') &&
                              reportData.hasOwnProperty('basicStatistics') &&
                              reportData.hasOwnProperty('applicationStatistics');
      
      results.push({
        test: 'レポートデータ構造',
        result: hasReportStructure ? 'SUCCESS' : 'FAILED',
        detail: 'レポートタイプ: ' + reportData.reportType
      });
      
      // テスト3: 期間情報確認
      var periodCorrect = reportData.period.year === testYear &&
                         reportData.period.month === testMonth;
      
      results.push({
        test: '期間情報正確性',
        result: periodCorrect ? 'SUCCESS' : 'FAILED',
        detail: reportData.period.displayName
      });
      
      // テスト4: 統計データ存在確認
      var hasStats = reportData.basicStatistics && 
                    reportData.applicationStatistics &&
                    reportData.grantStatistics;
      
      results.push({
        test: '統計データ存在',
        result: hasStats ? 'SUCCESS' : 'FAILED',
        detail: '全統計データ: ' + (hasStats ? '存在' : '不足')
      });
    }
    
    // 結果集計
    console.log('=== 月次レポート生成テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功',
      monthlyReport: monthlyReport
    };
    
  } catch (error) {
    console.error('月次レポート生成テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * HTMLレポート生成のテスト
 */
function testHTMLReportGeneration() {
  try {
    console.log('HTMLレポート生成テスト開始');
    
    // テスト用のレポートデータを作成
    var testReportData = {
      reportType: 'MONTHLY',
      period: {
        year: 2024,
        month: 8,
        displayName: '2024年8月'
      },
      generatedAt: new Date(),
      basicStatistics: {
        totalEmployees: 100,
        totalRemainingDays: 1500,
        averageRemainingDays: 15.0,
        lowRemainingCount: 5,
        zeroRemainingCount: 2
      },
      applicationStatistics: {
        totalApplications: 25,
        approvedApplications: 23,
        rejectedApplications: 1,
        pendingApplications: 1,
        approvalRate: 92
      },
      grantStatistics: {
        totalGrants: 10,
        totalGrantedDays: 150,
        initialGrants: 3,
        annualGrants: 7,
        averageGrantDays: 15.0
      },
      usageStatistics: {
        overallUsageRate: 65,
        fiveDayObligation: {
          targetEmployees: 80,
          compliantEmployees: 70,
          complianceRate: 87.5
        }
      },
      divisionStatistics: {
        'R': { name: 'ライズ', employees: 9, averageRemaining: 14.5, applications: 5 },
        'P': { name: 'パロン', employees: 48, averageRemaining: 15.2, applications: 12 },
        'S': { name: 'シエル', employees: 25, averageRemaining: 16.1, applications: 6 },
        'E': { name: 'EBISU', employees: 18, averageRemaining: 13.8, applications: 2 }
      }
    };
    
    var results = [];
    
    // テスト1: HTML生成実行
    var htmlReport = generateReportHTML(testReportData);
    var isValidHTML = typeof htmlReport === 'string' && htmlReport.length > 0;
    
    results.push({
      test: 'HTML生成実行',
      result: isValidHTML ? 'SUCCESS' : 'FAILED',
      detail: 'HTML長: ' + htmlReport.length + '文字'
    });
    
    // テスト2: 必要なHTMLタグ存在確認
    var hasRequiredTags = htmlReport.indexOf('<div') !== -1 &&
                         htmlReport.indexOf('<h1') !== -1 &&
                         htmlReport.indexOf('<h3') !== -1;
    
    results.push({
      test: '必要HTMLタグ存在',
      result: hasRequiredTags ? 'SUCCESS' : 'FAILED',
      detail: '基本HTMLタグ: ' + (hasRequiredTags ? '存在' : '不足')
    });
    
    // テスト3: データ埋め込み確認
    var hasEmbeddedData = htmlReport.indexOf('2024年8月') !== -1 &&
                         htmlReport.indexOf('100名') !== -1 &&
                         htmlReport.indexOf('25件') !== -1;
    
    results.push({
      test: 'データ埋め込み確認',
      result: hasEmbeddedData ? 'SUCCESS' : 'FAILED',
      detail: 'テストデータ埋め込み: ' + (hasEmbeddedData ? '成功' : '失敗')
    });
    
    // テスト4: スタイル適用確認
    var hasStyles = htmlReport.indexOf('style=') !== -1 &&
                   htmlReport.indexOf('color:') !== -1;
    
    results.push({
      test: 'スタイル適用確認',
      result: hasStyles ? 'SUCCESS' : 'FAILED',
      detail: 'CSSスタイル: ' + (hasStyles ? '適用済み' : '未適用')
    });
    
    // 結果集計
    console.log('=== HTMLレポート生成テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });
    
    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;
    
    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功',
      htmlReport: isValidHTML ? htmlReport.substring(0, 200) + '...' : 'HTML生成失敗'
    };
    
  } catch (error) {
    console.error('HTMLレポート生成テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}