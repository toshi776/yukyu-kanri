// =============================
// テストデータ準備・クリーンアップ
// =============================

/**
 * テストデータ準備関数
 * 全てのテストに必要なテストユーザーをマスターシートに追加
 */
function setupTestData() {
  console.log('=== テストデータ準備開始 ===');

  try {
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');

    if (!masterSheet) {
      throw new Error('マスターシートが見つかりません');
    }

    // 既存データを取得
    var data = masterSheet.getDataRange().getValues();
    var existingUsers = {};

    for (var i = 1; i < data.length; i++) {
      var userId = String(data[i][0]);
      if (userId) {
        existingUsers[userId] = true;
      }
    }

    // テストユーザーのリスト
    var testUsers = [
      // 基本テスト用
      {
        userId: 'TEST001',
        name: 'テスト太郎（付与履歴管理）',
        remaining: 0,
        remarks: 'テストユーザー',
        hireDate: new Date('2024-01-01'),
        weeklyWorkDays: 5,
        initialGrantDate: '',
        lastAnnualGrantDate: ''
      },
      {
        userId: 'TEST002',
        name: 'テスト花子（FIFO消費）',
        remaining: 0,
        remarks: 'テストユーザー',
        hireDate: new Date('2023-01-01'),
        weeklyWorkDays: 5,
        initialGrantDate: '',
        lastAnnualGrantDate: ''
      },
      {
        userId: 'TEST003',
        name: 'テスト次郎（失効処理）',
        remaining: 0,
        remarks: 'テストユーザー',
        hireDate: new Date('2022-01-01'),
        weeklyWorkDays: 5,
        initialGrantDate: '',
        lastAnnualGrantDate: ''
      },

      // エッジケーステスト用
      {
        userId: 'TEST_BOUNDARY_001',
        name: 'テスト境界値001',
        remaining: 0,
        remarks: 'エッジケーステスト用',
        hireDate: new Date('2024-06-01'),
        weeklyWorkDays: 5,
        initialGrantDate: '',
        lastAnnualGrantDate: ''
      },
      {
        userId: 'TEST_EXPIRY_TODAY',
        name: 'テスト失効当日',
        remaining: 0,
        remarks: 'エッジケーステスト用',
        hireDate: new Date('2023-01-01'),
        weeklyWorkDays: 5,
        initialGrantDate: '',
        lastAnnualGrantDate: ''
      },
      {
        userId: 'TEST_PAST_GRANT',
        name: 'テスト過去付与',
        remaining: 0,
        remarks: 'エッジケーステスト用',
        hireDate: new Date('2019-01-01'),
        weeklyWorkDays: 5,
        initialGrantDate: '',
        lastAnnualGrantDate: ''
      },
      {
        userId: 'TEST_FUTURE_GRANT',
        name: 'テスト未来付与',
        remaining: 0,
        remarks: 'エッジケーステスト用',
        hireDate: new Date('2029-01-01'),
        weeklyWorkDays: 5,
        initialGrantDate: '',
        lastAnnualGrantDate: ''
      },
      {
        userId: 'TEST_COMPLEX_FIFO',
        name: 'テスト複雑FIFO',
        remaining: 0,
        remarks: 'エッジケーステスト用',
        hireDate: new Date('2022-01-01'),
        weeklyWorkDays: 5,
        initialGrantDate: '',
        lastAnnualGrantDate: ''
      },
      {
        userId: 'TEST_EXPIRY_BEFORE',
        name: 'テスト失効前日',
        remaining: 0,
        remarks: 'エッジケーステスト用',
        hireDate: new Date('2023-01-01'),
        weeklyWorkDays: 5,
        initialGrantDate: '',
        lastAnnualGrantDate: ''
      },
      {
        userId: 'TEST_EXPIRY_AFTER',
        name: 'テスト失効翌日',
        remaining: 0,
        remarks: 'エッジケーステスト用',
        hireDate: new Date('2023-01-01'),
        weeklyWorkDays: 5,
        initialGrantDate: '',
        lastAnnualGrantDate: ''
      },
      {
        userId: 'TEST_ATTENDANCE',
        name: 'テスト出勤率',
        remaining: 0,
        remarks: 'エッジケーステスト用',
        hireDate: new Date('2023-04-01'),
        weeklyWorkDays: 5,
        initialGrantDate: '',
        lastAnnualGrantDate: ''
      },
      {
        userId: 'TEST_ANNUAL',
        name: 'テスト年次付与',
        remaining: 0,
        remarks: 'エッジケーステスト用',
        hireDate: new Date('2023-04-01'),
        weeklyWorkDays: 5,
        initialGrantDate: '',
        lastAnnualGrantDate: ''
      }
    ];

    var addedCount = 0;
    var skippedCount = 0;

    testUsers.forEach(function(user) {
      if (existingUsers[user.userId]) {
        console.log('スキップ（既に存在）: ' + user.userId);
        skippedCount++;
        return;
      }

      // マスターシートに追加
      var newRow = [
        user.userId,
        user.name,
        user.remaining,
        user.remarks,
        user.hireDate,
        user.weeklyWorkDays,
        user.initialGrantDate,
        user.lastAnnualGrantDate
      ];

      masterSheet.appendRow(newRow);
      console.log('追加: ' + user.userId + ' (' + user.name + ')');
      addedCount++;
    });

    console.log('=== テストデータ準備完了 ===');
    console.log('追加: ' + addedCount + '名');
    console.log('スキップ: ' + skippedCount + '名');

    return {
      success: true,
      message: 'テストデータを準備しました',
      added: addedCount,
      skipped: skippedCount,
      total: testUsers.length
    };

  } catch (error) {
    console.error('テストデータ準備エラー:', error);
    return {
      success: false,
      message: 'テストデータ準備に失敗しました: ' + error.message
    };
  }
}

/**
 * テストデータクリーンアップ関数（改善版）
 * テストユーザーと不整合データを全て削除
 */
function cleanupTestData() {
  console.log('=== テストデータクリーンアップ開始 ===');

  try {
    var ss = getSpreadsheet();
    var results = {
      master: 0,
      grantHistory: 0,
      urlManagement: 0,
      applications: 0
    };

    // 1. マスターシートからテストユーザーを削除
    console.log('1. マスターシートのクリーンアップ');
    var masterSheet = ss.getSheetByName('マスター');
    if (masterSheet) {
      var data = masterSheet.getDataRange().getValues();
      var rowsToDelete = [];

      for (var i = 1; i < data.length; i++) {
        var userId = String(data[i][0]);
        // TESTで始まるユーザーまたは不明なユーザー（123456等）を削除
        if (userId.startsWith('TEST') || userId === '123456' || userId === '123457' || userId === '123458') {
          rowsToDelete.push(i + 1);
        }
      }

      // 後ろから削除（行番号が変わらないようにするため）
      for (var i = rowsToDelete.length - 1; i >= 0; i--) {
        masterSheet.deleteRow(rowsToDelete[i]);
        results.master++;
      }

      console.log('マスターシートから削除: ' + results.master + '行');
    }

    // 2. 付与履歴からテストデータを削除
    console.log('2. 付与履歴のクリーンアップ');
    var grantHistorySheet = ss.getSheetByName('付与履歴');
    if (grantHistorySheet) {
      var data = grantHistorySheet.getDataRange().getValues();
      var rowsToDelete = [];

      for (var i = 1; i < data.length; i++) {
        var userId = String(data[i][0]);
        if (userId.startsWith('TEST') || userId === '123456' || userId === '123457' || userId === '123458') {
          rowsToDelete.push(i + 1);
        }
      }

      for (var i = rowsToDelete.length - 1; i >= 0; i--) {
        grantHistorySheet.deleteRow(rowsToDelete[i]);
        results.grantHistory++;
      }

      console.log('付与履歴から削除: ' + results.grantHistory + '行');
    }

    // 3. URL管理からテストデータを削除
    console.log('3. URL管理のクリーンアップ');
    var urlSheet = ss.getSheetByName('URL管理');
    if (urlSheet) {
      var urlData = urlSheet.getDataRange().getValues();
      var urlRowsToDelete = [];

      for (var i = 1; i < urlData.length; i++) {
        var userId = String(urlData[i][0]);
        if (userId.startsWith('TEST') || userId === '123456' || userId === '123457' || userId === '123458') {
          urlRowsToDelete.push(i + 1);
        }
      }

      for (var i = urlRowsToDelete.length - 1; i >= 0; i--) {
        urlSheet.deleteRow(urlRowsToDelete[i]);
        results.urlManagement++;
      }

      console.log('URL管理から削除: ' + results.urlManagement + '行');
    }

    // 4. 申請シートからテストデータを削除
    console.log('4. 申請シートのクリーンアップ');
    var applySheet = ss.getSheetByName('申請');
    if (applySheet) {
      var applyData = applySheet.getDataRange().getValues();
      var applyRowsToDelete = [];

      for (var i = 1; i < applyData.length; i++) {
        var userId = String(applyData[i][0]);
        if (userId.startsWith('TEST') || userId === '123456' || userId === '123457' || userId === '123458') {
          applyRowsToDelete.push(i + 1);
        }
      }

      for (var i = applyRowsToDelete.length - 1; i >= 0; i--) {
        applySheet.deleteRow(applyRowsToDelete[i]);
        results.applications++;
      }

      console.log('申請シートから削除: ' + results.applications + '行');
    }

    console.log('=== テストデータクリーンアップ完了 ===');
    console.log('削除サマリー:');
    console.log('  マスター: ' + results.master + '行');
    console.log('  付与履歴: ' + results.grantHistory + '行');
    console.log('  URL管理: ' + results.urlManagement + '行');
    console.log('  申請: ' + results.applications + '行');

    var totalDeleted = results.master + results.grantHistory + results.urlManagement + results.applications;

    return {
      success: true,
      message: 'テストデータをクリーンアップしました（合計' + totalDeleted + '行削除）',
      results: results,
      totalDeleted: totalDeleted
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
 * データ整合性修復関数
 * R01等の既存データの問題を修正
 */
function fixDataIntegrityIssues() {
  console.log('=== データ整合性修復開始 ===');

  try {
    var ss = getSpreadsheet();
    var results = [];

    // 問題1: R01がURL管理に存在するがマスターにない場合
    console.log('1. R01の問題を確認');
    var masterSheet = ss.getSheetByName('マスター');
    var urlSheet = ss.getSheetByName('URL管理');

    var masterData = masterSheet.getDataRange().getValues();
    var urlData = urlSheet.getDataRange().getValues();

    // R01がマスターに存在するか確認
    var r01InMaster = false;
    for (var i = 1; i < masterData.length; i++) {
      if (String(masterData[i][0]) === 'R01') {
        r01InMaster = true;
        break;
      }
    }

    // R01がURL管理に存在するか確認
    var r01InUrl = false;
    var r01UrlRow = -1;
    for (var i = 1; i < urlData.length; i++) {
      if (String(urlData[i][0]) === 'R01') {
        r01InUrl = true;
        r01UrlRow = i + 1;
        break;
      }
    }

    if (r01InUrl && !r01InMaster) {
      // URL管理にはあるがマスターにない → URL管理から削除
      console.log('R01をURL管理から削除（マスターに存在しないため）');
      urlSheet.deleteRow(r01UrlRow);
      results.push({
        issue: 'R01がマスターに存在しない',
        action: 'URL管理から削除',
        result: 'SUCCESS'
      });
    } else if (r01InMaster && !r01InUrl) {
      // マスターにはあるがURL管理にない → URL管理に追加
      console.log('R01のURLを生成してURL管理に追加');
      var urlKey = generateUrlKey();
      storeUrlKey('R01', urlKey);
      results.push({
        issue: 'R01がURL管理に存在しない',
        action: 'URL管理に追加',
        result: 'SUCCESS'
      });
    } else if (r01InMaster && r01InUrl) {
      console.log('R01は正常（マスターとURL管理両方に存在）');
      results.push({
        issue: 'R01の確認',
        action: '問題なし',
        result: 'SUCCESS'
      });
    } else {
      console.log('R01は存在しない（問題なし）');
      results.push({
        issue: 'R01の確認',
        action: '存在しない（正常）',
        result: 'SUCCESS'
      });
    }

    console.log('=== データ整合性修復完了 ===');
    results.forEach(function(result, index) {
      console.log((index + 1) + '. ' + result.issue + ': ' + result.action + ' (' + result.result + ')');
    });

    return {
      success: true,
      message: 'データ整合性を修復しました',
      results: results
    };

  } catch (error) {
    console.error('データ整合性修復エラー:', error);
    return {
      success: false,
      message: 'データ整合性修復に失敗しました: ' + error.message
    };
  }
}

/**
 * テスト準備のワンストップ関数
 * クリーンアップ → データ修復 → テストデータ準備を一括実行
 */
function prepareForTesting() {
  console.log('=== テスト準備（一括実行）開始 ===');

  var results = [];

  // ステップ1: クリーンアップ
  console.log('\n【ステップ1】テストデータクリーンアップ');
  var cleanupResult = cleanupTestData();
  results.push({
    step: 'クリーンアップ',
    result: cleanupResult
  });

  // ステップ2: データ整合性修復
  console.log('\n【ステップ2】データ整合性修復');
  var fixResult = fixDataIntegrityIssues();
  results.push({
    step: 'データ整合性修復',
    result: fixResult
  });

  // ステップ3: テストデータ準備
  console.log('\n【ステップ3】テストデータ準備');
  var setupResult = setupTestData();
  results.push({
    step: 'テストデータ準備',
    result: setupResult
  });

  console.log('\n=== テスト準備（一括実行）完了 ===');

  var allSuccess = results.every(function(r) { return r.result.success; });

  if (allSuccess) {
    console.log('✅ すべてのステップが成功しました');
    console.log('');
    console.log('次のステップ: 以下のテストを実行してください');
    console.log('  - runAllTests()');
    console.log('  - runEdgeCaseTests()');
    console.log('  - runDataIntegrityTests()');
  } else {
    console.log('❌ 一部のステップが失敗しました');
    results.forEach(function(r) {
      if (!r.result.success) {
        console.log('  失敗: ' + r.step + ' - ' + r.result.message);
      }
    });
  }

  return {
    success: allSuccess,
    message: allSuccess ? 'テスト準備が完了しました' : '一部のステップが失敗しました',
    results: results
  };
}
