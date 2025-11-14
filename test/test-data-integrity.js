// =============================
// データ整合性詳細確認テスト
// =============================

/**
 * データ整合性詳細確認テストの全実行
 */
function runDataIntegrityTests() {
  console.log('=== データ整合性詳細確認テスト開始 ===');

  var testResults = [];

  try {
    // テスト1: マスターシートと付与履歴の整合性
    console.log('\n1. マスターシートと付与履歴の整合性テスト');
    var masterHistoryTest = testMasterHistoryConsistency();
    testResults.push({
      testName: 'マスターシートと付与履歴の整合性',
      result: masterHistoryTest
    });

    // テスト2: 計算残日数と実残日数の一致
    console.log('\n2. 計算残日数と実残日数の一致テスト');
    var remainingCalcTest = testRemainingDaysCalculationAccuracy();
    testResults.push({
      testName: '計算残日数と実残日数の一致',
      result: remainingCalcTest
    });

    // テスト3: 申請日数と消費日数の整合性
    console.log('\n3. 申請日数と消費日数の整合性テスト');
    var applicationConsumptionTest = testApplicationConsumptionConsistency();
    testResults.push({
      testName: '申請日数と消費日数の整合性',
      result: applicationConsumptionTest
    });

    // テスト4: 付与日数の法定ルール準拠
    console.log('\n4. 付与日数の法定ルール準拠テスト');
    var legalComplianceTest = testLegalGrantCompliance();
    testResults.push({
      testName: '付与日数の法定ルール準拠',
      result: legalComplianceTest
    });

    // テスト5: 失効日の正確性
    console.log('\n5. 失効日の正確性テスト');
    var expiryDateTest = testExpiryDateAccuracy();
    testResults.push({
      testName: '失効日の正確性',
      result: expiryDateTest
    });

    // テスト6: URL管理とマスターの同期
    console.log('\n6. URL管理とマスターの同期テスト');
    var urlMasterSyncTest = testUrlMasterSync();
    testResults.push({
      testName: 'URL管理とマスターの同期',
      result: urlMasterSyncTest
    });

  } catch (error) {
    console.error('データ整合性テスト実行エラー:', error);
    testResults.push({
      testName: 'テスト実行',
      result: {
        success: false,
        message: 'テスト実行中にエラーが発生しました: ' + error.message
      }
    });
  }

  // 結果サマリーを表示
  console.log('\n=== データ整合性テスト結果サマリー ===');
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
 * マスターシートと付与履歴の整合性テスト
 */
function testMasterHistoryConsistency() {
  try {
    console.log('マスターシートと付与履歴の整合性テスト開始');

    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');
    var grantHistorySheet = ss.getSheetByName('付与履歴');

    if (!masterSheet || !grantHistorySheet) {
      throw new Error('必要なシートが見つかりません');
    }

    var results = [];

    // マスターシートデータ取得
    var masterData = masterSheet.getDataRange().getValues();
    var historyData = grantHistorySheet.getDataRange().getValues();

    // テスト1: 全マスターユーザーの付与履歴存在確認
    console.log('テスト1: 全マスターユーザーの付与履歴存在確認');
    var usersWithoutHistory = [];

    for (var i = 1; i < masterData.length; i++) {
      var userId = String(masterData[i][0]);
      if (!userId) continue;

      var hasHistory = false;
      for (var j = 1; j < historyData.length; j++) {
        if (String(historyData[j][0]) === userId) {
          hasHistory = true;
          break;
        }
      }

      if (!hasHistory) {
        usersWithoutHistory.push(userId);
      }
    }

    results.push({
      test: '全ユーザーに付与履歴',
      result: usersWithoutHistory.length === 0 ? 'SUCCESS' : 'FAILED',
      detail: '履歴なしユーザー: ' + usersWithoutHistory.length + '名' + (usersWithoutHistory.length > 0 ? ' (' + usersWithoutHistory.join(', ') + ')' : '')
    });

    // テスト2: 付与履歴の全ユーザーがマスターに存在
    console.log('テスト2: 付与履歴の全ユーザーがマスターに存在');
    var historyUsersNotInMaster = [];
    var historyUserIds = {};

    for (var j = 1; j < historyData.length; j++) {
      var historyUserId = String(historyData[j][0]);
      if (!historyUserId) continue;

      historyUserIds[historyUserId] = true;
    }

    for (var userId in historyUserIds) {
      var foundInMaster = false;
      for (var i = 1; i < masterData.length; i++) {
        if (String(masterData[i][0]) === userId) {
          foundInMaster = true;
          break;
        }
      }

      if (!foundInMaster && !userId.startsWith('TEST')) {
        historyUsersNotInMaster.push(userId);
      }
    }

    results.push({
      test: '付与履歴ユーザーがマスターに存在',
      result: historyUsersNotInMaster.length === 0 ? 'SUCCESS' : 'FAILED',
      detail: 'マスター未登録: ' + historyUsersNotInMaster.length + '名' + (historyUsersNotInMaster.length > 0 ? ' (' + historyUsersNotInMaster.join(', ') + ')' : '')
    });

    // テスト3: 残日数の範囲チェック（0以上）
    console.log('テスト3: 残日数の範囲チェック');
    var invalidRemainingDays = [];

    for (var i = 1; i < masterData.length; i++) {
      var userId = String(masterData[i][0]);
      if (!userId) continue;

      var remaining = calculateEffectiveRemainingDays(userId);
      if (remaining < 0) {
        invalidRemainingDays.push({ userId: userId, remaining: remaining });
      }
    }

    results.push({
      test: '残日数が0以上',
      result: invalidRemainingDays.length === 0 ? 'SUCCESS' : 'FAILED',
      detail: 'マイナス残日数: ' + invalidRemainingDays.length + '件'
    });

    // 結果集計
    console.log('=== マスターシートと付与履歴の整合性テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });

    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;

    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功',
      usersWithoutHistory: usersWithoutHistory,
      historyUsersNotInMaster: historyUsersNotInMaster,
      invalidRemainingDays: invalidRemainingDays
    };

  } catch (error) {
    console.error('マスターシートと付与履歴の整合性テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 計算残日数と実残日数の一致テスト
 */
function testRemainingDaysCalculationAccuracy() {
  try {
    console.log('計算残日数と実残日数の一致テスト開始');

    var ss = getSpreadsheet();
    var grantHistorySheet = ss.getSheetByName('付与履歴');

    if (!grantHistorySheet) {
      throw new Error('付与履歴シートが見つかりません');
    }

    var results = [];
    var historyData = grantHistorySheet.getDataRange().getValues();

    // ユーザーごとに計算
    var userTotals = {};

    for (var i = 1; i < historyData.length; i++) {
      var userId = String(historyData[i][0]);
      if (!userId || userId.startsWith('TEST')) continue;

      if (!userTotals[userId]) {
        userTotals[userId] = {
          totalGranted: 0,
          totalRemaining: 0,
          records: 0
        };
      }

      var grantDays = parseFloat(historyData[i][2]) || 0;
      var remainingDays = parseFloat(historyData[i][5]) || 0;
      var expiryDate = new Date(historyData[i][3]);
      var today = new Date();

      // 失効していないもののみカウント
      if (expiryDate > today) {
        userTotals[userId].totalGranted += grantDays;
        userTotals[userId].totalRemaining += remainingDays;
        userTotals[userId].records++;
      }
    }

    // 各ユーザーの計算結果を検証
    console.log('ユーザーごとの残日数検証');
    var mismatches = [];

    for (var userId in userTotals) {
      var calculatedRemaining = calculateEffectiveRemainingDays(userId);
      var summedRemaining = userTotals[userId].totalRemaining;

      // 小数点以下の誤差を許容（0.01日以内）
      var diff = Math.abs(calculatedRemaining - summedRemaining);

      if (diff > 0.01) {
        mismatches.push({
          userId: userId,
          calculated: calculatedRemaining,
          summed: summedRemaining,
          diff: diff
        });
      }
    }

    results.push({
      test: '計算残日数と合計の一致',
      result: mismatches.length === 0 ? 'SUCCESS' : 'FAILED',
      detail: '不一致ユーザー: ' + mismatches.length + '名 / 全' + Object.keys(userTotals).length + '名'
    });

    if (mismatches.length > 0) {
      console.log('不一致の詳細:');
      mismatches.forEach(function(m) {
        console.log('  - ' + m.userId + ': 計算=' + m.calculated + '日, 合計=' + m.summed + '日, 差=' + m.diff + '日');
      });
    }

    // 結果集計
    console.log('=== 計算残日数と実残日数の一致テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });

    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;

    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功',
      mismatches: mismatches
    };

  } catch (error) {
    console.error('計算残日数と実残日数の一致テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 申請日数と消費日数の整合性テスト
 */
function testApplicationConsumptionConsistency() {
  try {
    console.log('申請日数と消費日数の整合性テスト開始');

    var ss = getSpreadsheet();
    var applySheet = ss.getSheetByName('申請');
    var grantHistorySheet = ss.getSheetByName('付与履歴');

    if (!applySheet || !grantHistorySheet) {
      throw new Error('必要なシートが見つかりません');
    }

    var results = [];

    // 承認済み申請の合計日数を計算
    var applyData = applySheet.getDataRange().getValues();
    var userAppliedDays = {};

    for (var i = 1; i < applyData.length; i++) {
      var userId = String(applyData[i][1]);
      var applyDays = parseFloat(applyData[i][4]) || 0;
      var status = String(applyData[i][5]);

      if (status === 'Approved' && !userId.startsWith('TEST')) {
        if (!userAppliedDays[userId]) {
          userAppliedDays[userId] = 0;
        }
        userAppliedDays[userId] += applyDays;
      }
    }

    // 付与履歴から消費日数を計算
    var historyData = grantHistorySheet.getDataRange().getValues();
    var userConsumedDays = {};

    for (var i = 1; i < historyData.length; i++) {
      var userId = String(historyData[i][0]);
      if (!userId || userId.startsWith('TEST')) continue;

      var grantDays = parseFloat(historyData[i][2]) || 0;
      var remainingDays = parseFloat(historyData[i][5]) || 0;
      var consumedDays = grantDays - remainingDays;

      if (!userConsumedDays[userId]) {
        userConsumedDays[userId] = 0;
      }
      userConsumedDays[userId] += consumedDays;
    }

    // 比較
    console.log('申請日数と消費日数の比較');
    var inconsistencies = [];

    for (var userId in userAppliedDays) {
      var appliedDays = userAppliedDays[userId];
      var consumedDays = userConsumedDays[userId] || 0;

      // 小数点以下の誤差を許容（0.01日以内）
      var diff = Math.abs(appliedDays - consumedDays);

      if (diff > 0.01) {
        inconsistencies.push({
          userId: userId,
          applied: appliedDays,
          consumed: consumedDays,
          diff: diff
        });
      }
    }

    results.push({
      test: '申請日数と消費日数の一致',
      result: inconsistencies.length === 0 ? 'SUCCESS' : 'WARNING',
      detail: '不一致ユーザー: ' + inconsistencies.length + '名 / 全' + Object.keys(userAppliedDays).length + '名（手動消費等あり）'
    });

    if (inconsistencies.length > 0) {
      console.log('不一致の詳細（要確認）:');
      inconsistencies.slice(0, 5).forEach(function(m) {
        console.log('  - ' + m.userId + ': 申請=' + m.applied + '日, 消費=' + m.consumed + '日, 差=' + m.diff + '日');
      });
    }

    // 結果集計
    console.log('=== 申請日数と消費日数の整合性テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });

    var successCount = results.filter(function(r) { return r.result === 'SUCCESS' || r.result === 'WARNING'; }).length;

    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功',
      inconsistencies: inconsistencies
    };

  } catch (error) {
    console.error('申請日数と消費日数の整合性テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 付与日数の法定ルール準拠テスト
 */
function testLegalGrantCompliance() {
  try {
    console.log('付与日数の法定ルール準拠テスト開始');

    var ss = getSpreadsheet();
    var grantHistorySheet = ss.getSheetByName('付与履歴');

    if (!grantHistorySheet) {
      throw new Error('付与履歴シートが見つかりません');
    }

    var results = [];
    var historyData = grantHistorySheet.getDataRange().getValues();

    // 法定範囲外の付与を検出
    console.log('法定範囲外の付与検出');
    var outOfRangeGrants = [];

    for (var i = 1; i < historyData.length; i++) {
      var userId = String(historyData[i][0]);
      if (!userId || userId.startsWith('TEST')) continue;

      var grantDays = parseFloat(historyData[i][2]) || 0;
      var grantType = String(historyData[i][6]) || '';

      // 法定範囲チェック
      // 初回: 10日（週5日）、最大20日（6年以上）
      if (grantType.indexOf('初回') !== -1) {
        if (grantDays < 1 || grantDays > 15) {
          outOfRangeGrants.push({
            userId: userId,
            grantDays: grantDays,
            grantType: grantType,
            reason: '初回付与の範囲外（1-15日）'
          });
        }
      } else {
        // 年次付与: 11-20日
        if (grantDays < 10 || grantDays > 20) {
          outOfRangeGrants.push({
            userId: userId,
            grantDays: grantDays,
            grantType: grantType,
            reason: '年次付与の範囲外（10-20日）'
          });
        }
      }
    }

    results.push({
      test: '付与日数の法定範囲',
      result: outOfRangeGrants.length === 0 ? 'SUCCESS' : 'WARNING',
      detail: '範囲外付与: ' + outOfRangeGrants.length + '件（要確認）'
    });

    if (outOfRangeGrants.length > 0) {
      console.log('範囲外付与の詳細:');
      outOfRangeGrants.slice(0, 5).forEach(function(g) {
        console.log('  - ' + g.userId + ': ' + g.grantDays + '日 (' + g.grantType + ') - ' + g.reason);
      });
    }

    // 結果集計
    console.log('=== 付与日数の法定ルール準拠テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });

    var successCount = results.filter(function(r) { return r.result === 'SUCCESS' || r.result === 'WARNING'; }).length;

    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功',
      outOfRangeGrants: outOfRangeGrants
    };

  } catch (error) {
    console.error('付与日数の法定ルール準拠テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 失効日の正確性テスト
 */
function testExpiryDateAccuracy() {
  try {
    console.log('失効日の正確性テスト開始');

    var ss = getSpreadsheet();
    var grantHistorySheet = ss.getSheetByName('付与履歴');

    if (!grantHistorySheet) {
      throw new Error('付与履歴シートが見つかりません');
    }

    var results = [];
    var historyData = grantHistorySheet.getDataRange().getValues();

    // 失効日の正確性チェック
    console.log('失効日の正確性チェック');
    var incorrectExpiryDates = [];

    for (var i = 1; i < historyData.length; i++) {
      var userId = String(historyData[i][0]);
      if (!userId || userId.startsWith('TEST')) continue;

      var grantDate = new Date(historyData[i][1]);
      var expiryDate = new Date(historyData[i][3]);
      var workYears = parseFloat(historyData[i][7]) || 0;

      // 失効日は付与日から2年後の前日であるべき
      var expectedExpiryDate = new Date(grantDate);
      expectedExpiryDate.setFullYear(expectedExpiryDate.getFullYear() + 2);
      expectedExpiryDate.setDate(expectedExpiryDate.getDate() - 1);

      // 日付の比較（時刻を無視）
      var expiryDateOnly = new Date(expiryDate.getFullYear(), expiryDate.getMonth(), expiryDate.getDate());
      var expectedExpiryDateOnly = new Date(expectedExpiryDate.getFullYear(), expectedExpiryDate.getMonth(), expectedExpiryDate.getDate());

      if (expiryDateOnly.getTime() !== expectedExpiryDateOnly.getTime()) {
        incorrectExpiryDates.push({
          userId: userId,
          grantDate: Utilities.formatDate(grantDate, 'JST', 'yyyy-MM-dd'),
          expiryDate: Utilities.formatDate(expiryDate, 'JST', 'yyyy-MM-dd'),
          expectedExpiryDate: Utilities.formatDate(expectedExpiryDate, 'JST', 'yyyy-MM-dd')
        });
      }
    }

    results.push({
      test: '失効日の正確性',
      result: incorrectExpiryDates.length === 0 ? 'SUCCESS' : 'WARNING',
      detail: '不正確な失効日: ' + incorrectExpiryDates.length + '件（要確認）'
    });

    if (incorrectExpiryDates.length > 0) {
      console.log('不正確な失効日の詳細:');
      incorrectExpiryDates.slice(0, 5).forEach(function(e) {
        console.log('  - ' + e.userId + ': 付与=' + e.grantDate + ', 失効=' + e.expiryDate + ', 期待=' + e.expectedExpiryDate);
      });
    }

    // 結果集計
    console.log('=== 失効日の正確性テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });

    var successCount = results.filter(function(r) { return r.result === 'SUCCESS' || r.result === 'WARNING'; }).length;

    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功',
      incorrectExpiryDates: incorrectExpiryDates
    };

  } catch (error) {
    console.error('失効日の正確性テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * URL管理とマスターの同期テスト
 */
function testUrlMasterSync() {
  try {
    console.log('URL管理とマスターの同期テスト開始');

    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');
    var urlSheet = ss.getSheetByName('URL管理');

    if (!masterSheet || !urlSheet) {
      throw new Error('必要なシートが見つかりません');
    }

    var results = [];

    var masterData = masterSheet.getDataRange().getValues();
    var urlData = urlSheet.getDataRange().getValues();

    // テスト1: 全マスターユーザーにURLが存在
    console.log('テスト1: 全マスターユーザーにURLが存在');
    var usersWithoutUrl = [];

    for (var i = 1; i < masterData.length; i++) {
      var userId = String(masterData[i][0]);
      if (!userId) continue;

      var hasUrl = false;
      for (var j = 1; j < urlData.length; j++) {
        if (String(urlData[j][0]) === userId) {
          hasUrl = true;
          break;
        }
      }

      if (!hasUrl) {
        usersWithoutUrl.push(userId);
      }
    }

    results.push({
      test: '全ユーザーにURL存在',
      result: usersWithoutUrl.length === 0 ? 'SUCCESS' : 'FAILED',
      detail: 'URLなしユーザー: ' + usersWithoutUrl.length + '名' + (usersWithoutUrl.length > 0 ? ' (' + usersWithoutUrl.join(', ') + ')' : '')
    });

    // テスト2: URL管理の全ユーザーがマスターに存在
    console.log('テスト2: URL管理の全ユーザーがマスターに存在');
    var urlUsersNotInMaster = [];

    for (var j = 1; j < urlData.length; j++) {
      var urlUserId = String(urlData[j][0]);
      if (!urlUserId || urlUserId.startsWith('TEST')) continue;

      var foundInMaster = false;
      for (var i = 1; i < masterData.length; i++) {
        if (String(masterData[i][0]) === urlUserId) {
          foundInMaster = true;
          break;
        }
      }

      if (!foundInMaster) {
        urlUsersNotInMaster.push(urlUserId);
      }
    }

    results.push({
      test: 'URLユーザーがマスターに存在',
      result: urlUsersNotInMaster.length === 0 ? 'SUCCESS' : 'FAILED',
      detail: 'マスター未登録: ' + urlUsersNotInMaster.length + '名' + (urlUsersNotInMaster.length > 0 ? ' (' + urlUsersNotInMaster.join(', ') + ')' : '')
    });

    // 結果集計
    console.log('=== URL管理とマスターの同期テスト結果 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.test + ': ' + result.result + ' (' + result.detail + ')');
    });

    var successCount = results.filter(function(r) { return r.result === 'SUCCESS'; }).length;

    return {
      success: successCount === results.length,
      results: results,
      summary: successCount + '/' + results.length + ' テスト成功',
      usersWithoutUrl: usersWithoutUrl,
      urlUsersNotInMaster: urlUsersNotInMaster
    };

  } catch (error) {
    console.error('URL管理とマスターの同期テストエラー:', error);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + error.message
    };
  }
}
