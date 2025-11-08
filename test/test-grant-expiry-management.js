// =============================
// 付与管理・失効管理機能テストスクリプト
// =============================

/**
 * 付与管理・失効管理機能の統合テスト
 */
function runGrantExpiryManagementTests() {
  console.log('=== 付与管理・失効管理機能テスト開始 ===');

  var testResults = [];

  try {
    // テスト1: 6ヶ月付与対象者確認
    var sixMonthTest = testSixMonthGrantTargets();
    testResults.push({
      testName: '6ヶ月付与対象者確認',
      result: sixMonthTest
    });

    // テスト2: 年次付与対象者確認
    var annualTest = testAnnualGrantTargets();
    testResults.push({
      testName: '年次付与対象者確認',
      result: annualTest
    });

    // テスト3: 付与履歴取得
    var historyTest = testGrantHistory();
    testResults.push({
      testName: '付与履歴取得',
      result: historyTest
    });

    // テスト4: 失効予定確認（1週間）
    var expiryWeekTest = testExpiringLeaves(7);
    testResults.push({
      testName: '失効予定確認（1週間）',
      result: expiryWeekTest
    });

    // テスト5: 失効予定確認（1ヶ月）
    var expiryMonthTest = testExpiringLeaves(30);
    testResults.push({
      testName: '失効予定確認（1ヶ月）',
      result: expiryMonthTest
    });

    // テスト6: 失効予定確認（3ヶ月）
    var expiry3MonthTest = testExpiringLeaves(90);
    testResults.push({
      testName: '失効予定確認（3ヶ月）',
      result: expiry3MonthTest
    });

  } catch (error) {
    console.error('付与管理・失効管理テスト実行エラー:', error);
    testResults.push({
      testName: 'テスト実行',
      result: {
        success: false,
        message: 'テスト実行中にエラーが発生しました: ' + error.message
      }
    });
  }

  // 結果サマリーを表示
  console.log('\n=== 付与管理・失効管理テスト結果サマリー ===');
  var totalTests = testResults.length;
  var passedTests = 0;

  testResults.forEach(function(test, index) {
    var status = test.result.success ? '✅ PASS' : '❌ FAIL';
    console.log((index + 1) + '. ' + test.testName + ': ' + status);

    if (test.result.success) {
      passedTests++;
      if (test.result.message) {
        console.log('   ' + test.result.message);
      }
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
 * 6ヶ月付与対象者確認テスト
 */
function testSixMonthGrantTargets() {
  try {
    console.log('\n--- 6ヶ月付与対象者確認テスト ---');

    var targets = getSixMonthGrantTargets();

    console.log('抽出された対象者数:', targets.length + '名');

    if (targets.length > 0) {
      console.log('対象者リスト:');
      targets.forEach(function(target, index) {
        console.log('  ' + (index + 1) + '. ' + target.userId + ' (' + target.userName + ')');
        console.log('     入社日: ' + target.joinDate);
        console.log('     6ヶ月経過日: ' + target.sixMonthDate);
        console.log('     付与日数: ' + target.grantDays + '日');
      });
    } else {
      console.log('現在6ヶ月付与対象者はいません');
    }

    return {
      success: true,
      message: targets.length + '名の対象者を確認',
      targetCount: targets.length,
      targets: targets
    };

  } catch (error) {
    console.error('6ヶ月付与対象者確認テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 年次付与対象者確認テスト
 */
function testAnnualGrantTargets() {
  try {
    console.log('\n--- 年次付与対象者確認テスト ---');

    var targets = getAnnualGrantTargets();

    console.log('抽出された対象者数:', targets.length + '名');

    if (targets.length > 0) {
      console.log('対象者リスト:');
      targets.forEach(function(target, index) {
        console.log('  ' + (index + 1) + '. ' + target.userId + ' (' + target.name + ')');
        console.log('     入社日: ' + Utilities.formatDate(target.joinDate, 'JST', 'yyyy/MM/dd'));
        console.log('     勤続年数: ' + Math.floor(target.workYears) + '年');
        console.log('     付与日数: ' + target.grantDays + '日');
      });
    } else {
      console.log('現在年次付与対象者はいません');
    }

    return {
      success: true,
      message: targets.length + '名の対象者を確認',
      targetCount: targets.length,
      targets: targets
    };

  } catch (error) {
    console.error('年次付与対象者確認テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 付与履歴取得テスト
 */
function testGrantHistory() {
  try {
    console.log('\n--- 付与履歴取得テスト ---');

    var limit = 10;
    var history = getRecentGrantHistory(limit);

    console.log('取得した履歴件数:', history.length + '件');

    if (history.length > 0) {
      console.log('最新の付与履歴:');
      history.forEach(function(record, index) {
        console.log('  ' + (index + 1) + '. ' + record.grantDate + ' - ' +
                    record.userId + ' (' + record.userName + ')');
        console.log('     ' + record.grantType + ' - ' + record.days + '日');
      });
    } else {
      console.log('付与履歴が見つかりませんでした');
    }

    return {
      success: true,
      message: history.length + '件の履歴を取得',
      historyCount: history.length,
      history: history
    };

  } catch (error) {
    console.error('付与履歴取得テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 失効予定確認テスト
 * @param {number} days - 確認期間（日数）
 */
function testExpiringLeaves(days) {
  try {
    console.log('\n--- 失効予定確認テスト (' + days + '日以内) ---');

    var expiringLeaves = getExpiringLeaves(days);

    console.log('失効予定件数:', expiringLeaves.length + '件');

    if (expiringLeaves.length > 0) {
      console.log('失効予定リスト:');
      expiringLeaves.forEach(function(leave, index) {
        var today = new Date();
        var expiryDate = new Date(leave.expiryDate);
        var daysUntil = Math.ceil((expiryDate - today) / (1000 * 60 * 60 * 24));
        var urgency = daysUntil <= 7 ? '🔴緊急' : (daysUntil <= 30 ? '🟠注意' : '🟡要確認');

        console.log('  ' + (index + 1) + '. ' + leave.userId + ' (' + leave.userName + ') ' + urgency);
        console.log('     失効日: ' + leave.expiryDate + ' (残り' + daysUntil + '日)');
        console.log('     失効予定日数: ' + leave.days + '日');
        console.log('     付与日: ' + leave.grantDate);
      });
    } else {
      console.log(days + '日以内に失効予定の有給はありません');
    }

    return {
      success: true,
      message: expiringLeaves.length + '件の失効予定を確認',
      expiringCount: expiringLeaves.length,
      expiringLeaves: expiringLeaves
    };

  } catch (error) {
    console.error('失効予定確認テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 管理画面機能の動作確認（読み取り専用）
 */
function testAdminFunctionsReadOnly() {
  console.log('=== 管理画面機能の動作確認（読み取り専用） ===\n');

  var allTests = [];

  // 1. 6ヶ月付与対象者確認
  console.log('【1. 6ヶ月付与対象者確認】');
  var sixMonthTargets = getSixMonthGrantTargets();
  console.log('→ 対象者: ' + sixMonthTargets.length + '名\n');
  allTests.push({ name: '6ヶ月付与対象者確認', success: true, count: sixMonthTargets.length });

  // 2. 年次付与対象者確認
  console.log('【2. 年次付与対象者確認】');
  var annualTargets = getAnnualGrantTargets();
  console.log('→ 対象者: ' + annualTargets.length + '名\n');
  allTests.push({ name: '年次付与対象者確認', success: true, count: annualTargets.length });

  // 3. 付与履歴（最新50件）
  console.log('【3. 付与履歴取得】');
  var history = getRecentGrantHistory(50);
  console.log('→ 履歴件数: ' + history.length + '件\n');
  allTests.push({ name: '付与履歴取得', success: true, count: history.length });

  // 4. 失効予定（1週間以内）
  console.log('【4. 失効予定確認（1週間）】');
  var expiry7 = getExpiringLeaves(7);
  console.log('→ 失効予定: ' + expiry7.length + '件\n');
  allTests.push({ name: '失効予定（1週間）', success: true, count: expiry7.length });

  // 5. 失効予定（1ヶ月以内）
  console.log('【5. 失効予定確認（1ヶ月）】');
  var expiry30 = getExpiringLeaves(30);
  console.log('→ 失効予定: ' + expiry30.length + '件\n');
  allTests.push({ name: '失効予定（1ヶ月）', success: true, count: expiry30.length });

  // 6. 失効予定（3ヶ月以内）
  console.log('【6. 失効予定確認（3ヶ月）】');
  var expiry90 = getExpiringLeaves(90);
  console.log('→ 失効予定: ' + expiry90.length + '件\n');
  allTests.push({ name: '失効予定（3ヶ月）', success: true, count: expiry90.length });

  // 7. 失効予定（6ヶ月以内）
  console.log('【7. 失効予定確認（6ヶ月）】');
  var expiry180 = getExpiringLeaves(180);
  console.log('→ 失効予定: ' + expiry180.length + '件\n');
  allTests.push({ name: '失効予定（6ヶ月）', success: true, count: expiry180.length });

  // サマリー
  console.log('=== テスト結果サマリー ===');
  allTests.forEach(function(test, index) {
    console.log((index + 1) + '. ' + test.name + ': ✅ PASS (' + test.count + '件)');
  });
  console.log('\n全 ' + allTests.length + ' 項目のテストが成功しました！');

  return {
    success: true,
    tests: allTests,
    summary: {
      sixMonthTargets: sixMonthTargets.length,
      annualTargets: annualTargets.length,
      grantHistory: history.length,
      expiring7Days: expiry7.length,
      expiring30Days: expiry30.length,
      expiring90Days: expiry90.length,
      expiring180Days: expiry180.length
    }
  };
}
