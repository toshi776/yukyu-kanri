// =============================
// 有給付与システム
// =============================

/**
 * 付与ルールのテスト関数
 */
function testLeaveGrantCalculation() {
  console.log('=== 付与ルールテスト開始 ===');
  
  var testCases = [
    // 初回付与テスト
    { workYears: 0.5, weekDays: 5, isInitial: true, expected: 10, desc: '週5日・初回付与' },
    { workYears: 0.5, weekDays: 4, isInitial: true, expected: 7, desc: '週4日・初回付与' },
    { workYears: 0.5, weekDays: 3, isInitial: true, expected: 5, desc: '週3日・初回付与' },
    { workYears: 0.5, weekDays: 2, isInitial: true, expected: 3, desc: '週2日・初回付与' },
    { workYears: 0.5, weekDays: 1, isInitial: true, expected: 1, desc: '週1日・初回付与' },
    
    // 年次付与テスト（週5日）
    { workYears: 1.5, weekDays: 5, isInitial: false, expected: 11, desc: '週5日・1年6ヶ月' },
    { workYears: 2.5, weekDays: 5, isInitial: false, expected: 12, desc: '週5日・2年6ヶ月' },
    { workYears: 3.5, weekDays: 5, isInitial: false, expected: 14, desc: '週5日・3年6ヶ月' },
    { workYears: 4.5, weekDays: 5, isInitial: false, expected: 16, desc: '週5日・4年6ヶ月' },
    { workYears: 5.5, weekDays: 5, isInitial: false, expected: 18, desc: '週5日・5年6ヶ月' },
    { workYears: 6.5, weekDays: 5, isInitial: false, expected: 20, desc: '週5日・6年6ヶ月以上' },
    
    // 年次付与テスト（週4日）
    { workYears: 1.5, weekDays: 4, isInitial: false, expected: 8, desc: '週4日・1年6ヶ月' },
    { workYears: 6.5, weekDays: 4, isInitial: false, expected: 15, desc: '週4日・6年6ヶ月以上' },
    
    // 年次付与テスト（週3日）
    { workYears: 1.5, weekDays: 3, isInitial: false, expected: 6, desc: '週3日・1年6ヶ月' },
    { workYears: 6.5, weekDays: 3, isInitial: false, expected: 11, desc: '週3日・6年6ヶ月以上' }
  ];
  
  var passedTests = 0;
  var totalTests = testCases.length;
  
  testCases.forEach(function(testCase, index) {
    var result = calculateLeaveDays(testCase.workYears, testCase.weekDays, testCase.isInitial);
    var passed = result === testCase.expected;
    
    console.log('Test ' + (index + 1) + ' (' + testCase.desc + '):', {
      expected: testCase.expected,
      actual: result,
      passed: passed
    });
    
    if (passed) passedTests++;
  });
  
  console.log('=== テスト結果 ===');
  console.log('成功: ' + passedTests + '/' + totalTests);
  console.log('成功率: ' + Math.round((passedTests / totalTests) * 100) + '%');
  
  return {
    passed: passedTests,
    total: totalTests,
    success: passedTests === totalTests
  };
}

/**
 * 手動で付与処理をテスト実行（管理者用）
 */
function runLeaveGrantProcess() {
  try {
    console.log('=== 有給付与処理開始 ===');
    
    // テスト実行
    var testResult = testLeaveGrantCalculation();
    
    if (testResult.success) {
      return {
        success: true,
        message: '付与ルールテスト成功: ' + testResult.passed + '/' + testResult.total + ' テストケースが成功しました。\n\n実際の付与処理はマスターシートの構造を確定してから実装します。',
        testResults: testResult
      };
    } else {
      return {
        success: false,
        message: '付与ルールテスト失敗: ' + testResult.passed + '/' + testResult.total + ' テストケースが失敗しました。',
        testResults: testResult
      };
    }
    
  } catch (error) {
    console.error('付与処理エラー:', error);
    return {
      success: false,
      message: '付与処理中にエラーが発生しました: ' + error.message
    };
  }
}

// =============================
// 付与履歴管理システム
// =============================

/**
 * 付与履歴シートを取得または作成
 * @param {Spreadsheet} ss - スプレッドシート
 * @return {Sheet} 付与履歴シート
 */
function getOrCreateGrantHistorySheet(ss) {
  var sheetName = '付与履歴';
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    console.log('付与履歴シートを新規作成します');
    sheet = ss.insertSheet(sheetName);
    
    // ヘッダー行を追加
    sheet.getRange(1, 1, 1, 8).setValues([
      ['利用者番号', '付与日', '付与日数', '失効日', '残日数', '付与タイプ', '勤続年数', '作成日時']
    ]);
    
    // ヘッダー行の書式設定
    var headerRange = sheet.getRange(1, 1, 1, 8);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4CAF50');
    headerRange.setFontColor('white');
    
    // 列幅調整
    sheet.setColumnWidth(1, 100); // 利用者番号
    sheet.setColumnWidth(2, 120); // 付与日
    sheet.setColumnWidth(3, 80);  // 付与日数
    sheet.setColumnWidth(4, 120); // 失効日
    sheet.setColumnWidth(5, 80);  // 残日数
    sheet.setColumnWidth(6, 100); // 付与タイプ
    sheet.setColumnWidth(7, 80);  // 勤続年数
    sheet.setColumnWidth(8, 150); // 作成日時
  }
  
  return sheet;
}

/**
 * 有給を付与する
 * @param {string} userId - 利用者番号
 * @param {Date} grantDate - 付与日
 * @param {number} grantDays - 付与日数
 * @param {string} grantType - 付与タイプ（'初回' or '年次'）
 * @param {number} workYears - 勤続年数
 * @return {Object} 付与結果
 */
function grantLeave(userId, grantDate, grantDays, grantType, workYears) {
  try {
    console.log('有給付与処理開始:', {
      userId: userId,
      grantDate: grantDate,
      grantDays: grantDays,
      grantType: grantType,
      workYears: workYears
    });
    
    var ss = getSpreadsheet();
    var grantHistorySheet = getOrCreateGrantHistorySheet(ss);
    var masterSheet = ss.getSheetByName('マスター');
    
    if (!masterSheet) {
      throw new Error('マスターシートが見つかりません');
    }
    
    // 失効日を計算（付与日から2年後）
    var expiryDate = new Date(grantDate);
    expiryDate.setFullYear(expiryDate.getFullYear() + 2);
    
    // 付与履歴に追加
    grantHistorySheet.appendRow([
      userId,
      grantDate,
      grantDays,
      expiryDate,
      grantDays, // 残日数（初期値は付与日数と同じ）
      grantType,
      workYears,
      new Date()
    ]);
    
    // マスターシートの残日数を更新
    var masterData = masterSheet.getDataRange().getValues();
    var userRowIndex = -1;
    
    for (var i = 1; i < masterData.length; i++) {
      if (String(masterData[i][0]) === String(userId)) {
        userRowIndex = i;
        break;
      }
    }
    
    if (userRowIndex === -1) {
      throw new Error('利用者が見つかりません: ' + userId);
    }
    
    // 現在の残日数を取得して加算
    var currentRemaining = Number(masterData[userRowIndex][2] || 0);
    var newRemaining = currentRemaining + grantDays;
    
    // マスターシートを更新
    masterSheet.getRange(userRowIndex + 1, 3).setValue(newRemaining);
    
    console.log('有給付与完了:', {
      userId: userId,
      grantDays: grantDays,
      previousRemaining: currentRemaining,
      newRemaining: newRemaining
    });
    
    return {
      success: true,
      userId: userId,
      grantDays: grantDays,
      previousRemaining: currentRemaining,
      newRemaining: newRemaining,
      expiryDate: expiryDate,
      message: '有給を付与しました: ' + grantDays + '日'
    };
    
  } catch (error) {
    console.error('有給付与エラー:', error);
    return {
      success: false,
      userId: userId,
      error: error.message,
      message: '有給付与に失敗しました: ' + error.message
    };
  }
}

/**
 * 利用者の付与履歴を取得
 * @param {string} userId - 利用者番号
 * @return {Array} 付与履歴
 */
function getUserGrantHistory(userId) {
  try {
    var ss = getSpreadsheet();
    var grantHistorySheet = getOrCreateGrantHistorySheet(ss);
    
    var data = grantHistorySheet.getDataRange().getValues();
    var history = [];
    
    // ヘッダー行をスキップして処理
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      if (String(row[0]) === String(userId)) {
        history.push({
          userId: row[0],
          grantDate: row[1],
          grantDays: row[2],
          expiryDate: row[3],
          remainingDays: row[4],
          grantType: row[5],
          workYears: row[6],
          createdAt: row[7]
        });
      }
    }
    
    // 付与日の降順でソート（新しいものが上）
    history.sort(function(a, b) {
      return new Date(b.grantDate) - new Date(a.grantDate);
    });
    
    return history;
    
  } catch (error) {
    console.error('付与履歴取得エラー:', error);
    return [];
  }
}

/**
 * 最近の付与履歴を取得（管理画面用）
 * @param {number} limit - 取得件数（デフォルト: 50）
 * @return {Array} 付与履歴
 */
function getRecentGrantHistory(limit) {
  try {
    limit = limit || 50;

    var ss = getSpreadsheet();
    var grantHistorySheet = getOrCreateGrantHistorySheet(ss);
    var masterSheet = ss.getSheetByName('マスター');

    var data = grantHistorySheet.getDataRange().getValues();
    var history = [];

    // マスターシートから利用者名を取得するためのマップを作成
    var userNameMap = {};
    if (masterSheet) {
      var masterData = masterSheet.getDataRange().getValues();
      for (var i = 1; i < masterData.length; i++) {
        var userId = String(masterData[i][0]);
        var userName = String(masterData[i][1] || '');
        if (userId) {
          userNameMap[userId] = userName;
        }
      }
    }

    // ヘッダー行をスキップして処理
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var userId = String(row[0]);

      // 空行をスキップ
      if (!userId || userId.trim() === '') {
        continue;
      }

      history.push({
        userId: userId,
        userName: userNameMap[userId] || '-',
        grantDate: formatDate(row[1]),
        grantDays: row[2],
        expiryDate: formatDate(row[3]),
        remainingDays: row[4],
        grantType: row[5],
        workYears: row[6],
        createdAt: formatDate(row[7])
      });
    }

    // 付与日の降順でソート（新しいものが上）
    history.sort(function(a, b) {
      return new Date(b.grantDate) - new Date(a.grantDate);
    });

    // 上位limit件のみ返す
    return history.slice(0, limit);

  } catch (error) {
    console.error('最近の付与履歴取得エラー:', error);
    console.error('エラー詳細:', error.message);
    console.error('スタックトレース:', error.stack);
    return [];
  }
}

/**
 * FIFO方式で有給を消費
 * @param {string} userId - 利用者番号
 * @param {number} useDays - 使用日数
 * @return {Object} 消費結果
 */
function consumeLeave(userId, useDays) {
  try {
    console.log('有給消費処理開始:', userId, useDays);

    // 入力値のバリデーション
    if (useDays <= 0) {
      throw new Error('消費日数は0より大きい値を指定してください');
    }

    var ss = getSpreadsheet();
    var grantHistorySheet = getOrCreateGrantHistorySheet(ss);
    
    var data = grantHistorySheet.getDataRange().getValues();
    var remainingToConsume = useDays;
    var consumedFromGrants = [];
    
    // FIFO方式で消費（古い付与分から順に消費）
    for (var i = 1; i < data.length; i++) {
      if (remainingToConsume <= 0) break;
      
      var row = data[i];
      
      if (String(row[0]) === String(userId)) {
        var grantDate = new Date(row[1]);
        var availableDays = Number(row[4]); // 残日数
        
        if (availableDays > 0) {
          var consumeFromThisGrant = Math.min(remainingToConsume, availableDays);
          var newAvailable = availableDays - consumeFromThisGrant;
          
          // シートを更新
          grantHistorySheet.getRange(i + 1, 5).setValue(newAvailable);
          
          consumedFromGrants.push({
            grantDate: grantDate,
            consumed: consumeFromThisGrant,
            remaining: newAvailable
          });
          
          remainingToConsume -= consumeFromThisGrant;
        }
      }
    }
    
    if (remainingToConsume > 0) {
      throw new Error('有給残日数が不足しています');
    }
    
    console.log('有給消費完了:', consumedFromGrants);
    
    return {
      success: true,
      consumedFromGrants: consumedFromGrants,
      totalConsumed: useDays,
      message: '有給を消費しました: ' + useDays + '日'
    };
    
  } catch (error) {
    console.error('有給消費エラー:', error);
    return {
      success: false,
      error: error.message,
      message: '有給消費に失敗しました: ' + error.message
    };
  }
}

/**
 * 利用者の有効な残日数を計算
 * @param {string} userId - 利用者番号
 * @return {number} 有効な残日数
 */
function calculateEffectiveRemainingDays(userId) {
  try {
    var ss = getSpreadsheet();
    var grantHistorySheet = getOrCreateGrantHistorySheet(ss);
    
    var data = grantHistorySheet.getDataRange().getValues();
    var totalRemaining = 0;
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // 失効していない付与分の残日数を合計
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      if (String(row[0]) === String(userId)) {
        var expiryDate = new Date(row[3]);
        expiryDate.setHours(0, 0, 0, 0);
        var remainingDays = Number(row[4]);
        
        // 失効していない場合のみカウント
        if (expiryDate > today && remainingDays > 0) {
          totalRemaining += remainingDays;
        }
      }
    }
    
    return totalRemaining;
    
  } catch (error) {
    console.error('残日数計算エラー:', error);
    return 0;
  }
}

/**
 * 失効処理を実行
 * @return {Object} 失効処理結果
 */
function processExpiredLeaves() {
  try {
    console.log('失効処理開始');
    
    var ss = getSpreadsheet();
    var grantHistorySheet = getOrCreateGrantHistorySheet(ss);
    var masterSheet = ss.getSheetByName('マスター');
    
    var data = grantHistorySheet.getDataRange().getValues();
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    
    var expiredCount = 0;
    var expiredUsers = [];
    
    // 失効対象を検索
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var userId = String(row[0]);
      var expiryDate = new Date(row[3]);
      expiryDate.setHours(0, 0, 0, 0);
      var remainingDays = Number(row[4]);
      
      // 今日が失効日以降で、残日数がある場合
      if (expiryDate <= today && remainingDays > 0) {
        // 残日数を0に設定
        grantHistorySheet.getRange(i + 1, 5).setValue(0);
        
        expiredCount += remainingDays;
        expiredUsers.push({
          userId: userId,
          expiredDays: remainingDays,
          expiryDate: expiryDate
        });
        
        console.log('失効処理:', userId, remainingDays + '日失効');
      }
    }
    
    // マスターシートの残日数を更新
    if (expiredUsers.length > 0) {
      var masterData = masterSheet.getDataRange().getValues();
      var updatedUsers = [];
      
      expiredUsers.forEach(function(expired) {
        for (var i = 1; i < masterData.length; i++) {
          if (String(masterData[i][0]) === expired.userId) {
            var newRemaining = calculateEffectiveRemainingDays(expired.userId);
            masterSheet.getRange(i + 1, 3).setValue(newRemaining);
            updatedUsers.push({
              userId: expired.userId,
              newRemaining: newRemaining
            });
            break;
          }
        }
      });
    }
    
    console.log('失効処理完了:', expiredCount + '日失効、' + expiredUsers.length + '名に影響');
    
    return {
      success: true,
      expiredCount: expiredCount,
      expiredUsers: expiredUsers,
      message: expiredCount + '日の有給が失効しました（' + expiredUsers.length + '名）'
    };
    
  } catch (error) {
    console.error('失効処理エラー:', error);
    return {
      success: false,
      error: error.message,
      message: '失効処理に失敗しました: ' + error.message
    };
  }
}

/**
 * 失効予定の有給を取得（管理画面用）
 * @param {number} days - 何日以内の失効予定を取得するか
 * @return {Array} 失効予定リスト
 */
function getExpiringLeaves(days) {
  try {
    days = days || 30;

    var ss = getSpreadsheet();
    var grantHistorySheet = getOrCreateGrantHistorySheet(ss);
    var masterSheet = ss.getSheetByName('マスター');

    var data = grantHistorySheet.getDataRange().getValues();
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var targetDate = new Date(today);
    targetDate.setDate(targetDate.getDate() + days);

    var expiringLeaves = [];

    // マスターシートから利用者名を取得するためのマップを作成
    var userNameMap = {};
    if (masterSheet) {
      var masterData = masterSheet.getDataRange().getValues();
      for (var i = 1; i < masterData.length; i++) {
        var userId = String(masterData[i][0]);
        var userName = String(masterData[i][1] || '');
        if (userId) {
          userNameMap[userId] = userName;
        }
      }
    }

    // 失効予定を検索
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var userId = String(row[0]);
      var grantDate = row[1];
      var expiryDate = new Date(row[3]);
      expiryDate.setHours(0, 0, 0, 0);
      var remainingDays = Number(row[4]);

      // 失効日が対象期間内で、残日数がある場合
      if (expiryDate > today && expiryDate <= targetDate && remainingDays > 0) {
        expiringLeaves.push({
          userId: userId,
          userName: userNameMap[userId] || '-',
          grantDate: formatDate(grantDate),
          expiryDate: formatDate(expiryDate),
          remainingDays: remainingDays
        });
      }
    }

    // 失効日の昇順でソート（近い順）
    expiringLeaves.sort(function(a, b) {
      return new Date(a.expiryDate) - new Date(b.expiryDate);
    });

    return expiringLeaves;

  } catch (error) {
    console.error('失効予定取得エラー:', error);
    return [];
  }
}

// =============================
// 6ヶ月付与処理（初回付与）
// =============================

/**
 * 6ヶ月付与処理を実行
 * @return {Object} 処理結果
 */
function processSixMonthGrants() {
  try {
    console.log('=== 6ヶ月付与処理開始 ===');
    
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // 6ヶ月付与対象者を抽出
    var targets = getSixMonthGrantTargets();
    
    if (targets.length === 0) {
      console.log('6ヶ月付与対象者はいません');
      return {
        success: true,
        processedCount: 0,
        message: '本日の6ヶ月付与対象者はいません'
      };
    }
    
    console.log('6ヶ月付与対象者:', targets.length + '名');
    
    var results = [];
    var successCount = 0;
    
    // 各対象者を処理
    targets.forEach(function(target) {
      try {
        console.log('6ヶ月付与処理:', target.userId, target.name);
        
        // 出勤率チェック（80%以上）
        var attendanceRate = checkAttendanceRate(target.userId, target.hireDate, today);
        
        if (attendanceRate < 0.8) {
          console.log('出勤率不足により付与見送り:', target.userId, '出勤率:' + Math.round(attendanceRate * 100) + '%');
          results.push({
            userId: target.userId,
            name: target.name,
            success: false,
            reason: '出勤率不足',
            attendanceRate: Math.round(attendanceRate * 100) + '%'
          });
          return;
        }
        
        // 付与日数を計算
        var grantDays = calculateLeaveDays(0.5, target.weeklyWorkDays, true);
        if (grantDays <= 0) {
          console.log('付与テーブル外のためスキップ:', target.userId);
          results.push({
            userId: target.userId,
            name: target.name,
            success: false,
            reason: '付与対象外（テーブル外）'
          });
          return;
        }
        
        // 有給を付与
        var grantResult = grantLeave(
          target.userId,
          target.sixMonthDate,
          grantDays,
          '初回',
          0.5
        );
        
        if (grantResult.success) {
          // 利用者マスターに初回付与日を記録
          recordInitialGrantDate(target.userId, target.sixMonthDate);
          
          // 通知送信
          sendSixMonthGrantNotification(target, grantDays);
          
          successCount++;
          results.push({
            userId: target.userId,
            name: target.name,
            success: true,
            grantDays: grantDays,
            grantDate: target.sixMonthDate,
            attendanceRate: Math.round(attendanceRate * 100) + '%'
          });
          
          console.log('6ヶ月付与完了:', target.userId, grantDays + '日付与');
        } else {
          results.push({
            userId: target.userId,
            name: target.name,
            success: false,
            reason: grantResult.error
          });
        }
        
      } catch (error) {
        console.error('6ヶ月付与処理エラー:', target.userId, error);
        results.push({
          userId: target.userId,
          name: target.name,
          success: false,
          reason: error.message
        });
      }
    });
    
    console.log('6ヶ月付与処理完了:', successCount + '/' + targets.length + '名成功');
    
    return {
      success: true,
      processedCount: successCount,
      totalTargets: targets.length,
      results: results,
      message: successCount + '/' + targets.length + '名に6ヶ月付与を実行しました'
    };
    
  } catch (error) {
    console.error('6ヶ月付与処理エラー:', error);
    return {
      success: false,
      error: error.message,
      message: '6ヶ月付与処理に失敗しました: ' + error.message
    };
  }
}

/**
 * 6ヶ月付与対象者を抽出
 * @return {Array} 対象者リスト
 */
function getSixMonthGrantTargets() {
  try {
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');
    
    if (!masterSheet) {
      throw new Error('マスターシートが見つかりません');
    }
    
    var data = masterSheet.getDataRange().getValues();
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    
    var targets = [];
    
    // ヘッダー行をスキップして処理
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      var userId = String(row[0]);
      var name = String(row[1]);
      var hireDateStr = row[4]; // E列: 入社日
      var weeklyWorkDaysStr = row[5]; // F列: 週所定労働日数
      var initialGrantDateStr = row[6]; // G列: 初回付与日
      
      if (!hireDateStr || initialGrantDateStr) {
        continue; // 入社日がないか、既に初回付与済み
      }
      
      var hireDate = new Date(hireDateStr);
      hireDate.setHours(0, 0, 0, 0);
      
      // 6ヶ月後の日付を計算
      var sixMonthDate = new Date(hireDate);
      sixMonthDate.setMonth(sixMonthDate.getMonth() + 6);
      sixMonthDate.setHours(0, 0, 0, 0);
      
      // 今日が6ヶ月後以降かチェック
      if (today >= sixMonthDate) {
        var weeklyWorkDays = Number(weeklyWorkDaysStr) || 5; // デフォルト5日
        var grantDays = calculateInitialLeaveDays(weeklyWorkDays);
        var daysFromHire = Math.floor((today - hireDate) / (1000 * 60 * 60 * 24));

        targets.push({
          userId: userId,
          userName: name,  // nameではなくuserName
          hireDate: formatDate(hireDate),
          sixMonthDate: formatDate(sixMonthDate),
          weeklyWorkDays: weeklyWorkDays,
          grantDays: grantDays,
          daysFromHire: daysFromHire
        });

        console.log('6ヶ月付与対象:', userId, name, '6ヶ月経過日:' + Utilities.formatDate(sixMonthDate, 'JST', 'yyyy/MM/dd'));
      }
    }

    console.log('=== getSixMonthGrantTargets() 返り値 ===');
    console.log('targets配列の長さ:', targets.length);
    console.log('targets配列の内容:', JSON.stringify(targets));
    console.log('返り値のtargets:', targets);

    return targets;

  } catch (error) {
    console.error('6ヶ月付与対象者抽出エラー:', error);
    console.error('エラー詳細:', error.message);
    console.error('スタックトレース:', error.stack);
    return [];
  }
}

/**
 * 出勤率をチェック（簡易版）
 * @param {string} userId - 利用者番号
 * @param {Date} hireDate - 入社日
 * @param {Date} checkDate - チェック日
 * @return {number} 出勤率（0-1）
 */
function checkAttendanceRate(userId, hireDate, checkDate) {
  try {
    // 実際の実装では勤怠データから出勤率を計算
    // ここでは簡易版として、申請データから概算する
    
    var ss = getSpreadsheet();
    var applySheet = ss.getSheetByName('申請');
    
    if (!applySheet) {
      console.log('申請シートが見つからないため、出勤率100%とみなします');
      return 1.0; // 申請データがない場合は100%とみなす
    }
    
    var data = applySheet.getDataRange().getValues();
    var totalLeaveDays = 0;
    
    // 6ヶ月間の有給取得日数を計算
    var sixMonthsAgo = new Date(checkDate);
    sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      if (String(row[0]) === String(userId) && row[3] && row[5] === 'Approved') {
        var applyDate = new Date(row[3]);
        
        if (applyDate >= hireDate && applyDate >= sixMonthsAgo && applyDate <= checkDate) {
          totalLeaveDays += Number(row[7] || 1); // 申請日数
        }
      }
    }
    
    // 6ヶ月間の稼働日数を計算（平日のみ、祝日除外は簡易版では省略）
    var totalWorkDays = 0;
    var currentDate = new Date(Math.max(hireDate, sixMonthsAgo));
    
    while (currentDate <= checkDate) {
      var dayOfWeek = currentDate.getDay();
      if (dayOfWeek !== 0 && dayOfWeek !== 6) { // 土日以外
        totalWorkDays++;
      }
      currentDate.setDate(currentDate.getDate() + 1);
    }
    
    if (totalWorkDays === 0) {
      return 1.0; // 稼働日がない場合は100%
    }
    
    var attendanceDays = totalWorkDays - totalLeaveDays;
    var attendanceRate = attendanceDays / totalWorkDays;
    
    console.log('出勤率計算:', {
      userId: userId,
      totalWorkDays: totalWorkDays,
      totalLeaveDays: totalLeaveDays,
      attendanceDays: attendanceDays,
      attendanceRate: Math.round(attendanceRate * 100) + '%'
    });
    
    return Math.max(0, Math.min(1, attendanceRate));
    
  } catch (error) {
    console.error('出勤率チェックエラー:', error);
    return 1.0; // エラー時は100%とみなす
  }
}

/**
 * 初回付与日をマスターシートに記録
 * @param {string} userId - 利用者番号
 * @param {Date} grantDate - 付与日
 */
function recordInitialGrantDate(userId, grantDate) {
  try {
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');
    
    var data = masterSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userId)) {
        // G列に初回付与日を記録
        masterSheet.getRange(i + 1, 7).setValue(grantDate);
        console.log('初回付与日記録:', userId, Utilities.formatDate(grantDate, 'JST', 'yyyy/MM/dd'));
        break;
      }
    }
    
  } catch (error) {
    console.error('初回付与日記録エラー:', error);
  }
}

/**
 * 6ヶ月付与通知を送信
 * @param {Object} target - 対象者情報
 * @param {number} grantDays - 付与日数
 */
function sendSixMonthGrantNotification(target, grantDays) {
  try {
    // 通知システムを使用（テストモードで実行）
    var notificationData = {
      userId: target.userId,
      userName: target.name,
      grantDays: grantDays,
      grantDate: Utilities.formatDate(target.sixMonthDate, 'JST', 'yyyy年M月d日'),
      grantType: '初回付与（入社6ヶ月）'
    };
    
    // 通知送信（実装は notification.js の関数を使用）
    console.log('6ヶ月付与通知:', notificationData);
    
    // 実際の通知送信はテストモードなので省略
    // sendGrantNotification(notificationData);
    
  } catch (error) {
    console.error('6ヶ月付与通知エラー:', error);
  }
}

/**
 * マスターシートの拡張（必要列の追加）
 */
function setupMasterSheetColumns() {
  try {
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');

    if (!masterSheet) {
      throw new Error('マスターシートが見つかりません');
    }

    // 現在の列数を確認
    var lastColumn = masterSheet.getLastColumn();

    // 必要な列が不足している場合は追加
    var requiredColumns = [
      'A: 利用者番号',
      'B: 利用者名',
      'C: 残有給日数',
      'D: 備考',
      'E: 入社日',
      'F: 週所定労働日数',
      'G: 初回付与日',
      'H: 最新年次付与日'
    ];

    if (lastColumn < 8) {
      console.log('マスターシートに列を追加します');

      // ヘッダー行を更新
      var headerRange = masterSheet.getRange(1, 1, 1, 8);
      headerRange.setValues([
        ['利用者番号', '利用者名', '残有給日数', '備考', '入社日', '週所定労働日数', '初回付与日', '最新年次付与日']
      ]);

      console.log('マスターシートの列構成を更新しました');
    }

    return {
      success: true,
      message: 'マスターシートの設定完了'
    };

  } catch (error) {
    console.error('マスターシート設定エラー:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * マスターシートのフォーマットを整える（完全版）
 * @return {Object} 処理結果
 */
function formatMasterSheet() {
  try {
    console.log('=== マスターシートフォーマット設定開始 ===');

    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');

    if (!masterSheet) {
      throw new Error('マスターシートが見つかりません');
    }

    // 1. ヘッダー行を設定
    var headers = [
      '利用者番号',
      '利用者名',
      '残有給日数',
      '備考',
      '入社日',
      '週所定労働日数',
      '初回付与日',
      '最新年次付与日'
    ];

    var headerRange = masterSheet.getRange(1, 1, 1, 8);
    headerRange.setValues([headers]);

    // ヘッダー行の書式設定
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4CAF50');
    headerRange.setFontColor('white');
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');

    // 2. 列幅を設定
    masterSheet.setColumnWidth(1, 100);  // A: 利用者番号
    masterSheet.setColumnWidth(2, 150);  // B: 利用者名
    masterSheet.setColumnWidth(3, 100);  // C: 残有給日数
    masterSheet.setColumnWidth(4, 200);  // D: 備考
    masterSheet.setColumnWidth(5, 120);  // E: 入社日
    masterSheet.setColumnWidth(6, 140);  // F: 週所定労働日数
    masterSheet.setColumnWidth(7, 120);  // G: 初回付与日
    masterSheet.setColumnWidth(8, 140);  // H: 最新年次付与日

    // 3. データ行のフォーマット設定
    var lastRow = Math.max(masterSheet.getLastRow(), 100); // 最低100行確保

    // C列（残有給日数）: 数値フォーマット、中央揃え
    var remainingRange = masterSheet.getRange(2, 3, lastRow - 1, 1);
    remainingRange.setNumberFormat('0');
    remainingRange.setHorizontalAlignment('center');

    // E列（入社日）: 日付フォーマット
    var hireDateRange = masterSheet.getRange(2, 5, lastRow - 1, 1);
    hireDateRange.setNumberFormat('yyyy/mm/dd');

    // F列（週所定労働日数）: 数値フォーマット、中央揃え
    var weekDaysRange = masterSheet.getRange(2, 6, lastRow - 1, 1);
    weekDaysRange.setNumberFormat('0');
    weekDaysRange.setHorizontalAlignment('center');

    // G列（初回付与日）: 日付フォーマット
    var initialGrantRange = masterSheet.getRange(2, 7, lastRow - 1, 1);
    initialGrantRange.setNumberFormat('yyyy/mm/dd');

    // H列（最新年次付与日）: 日付フォーマット
    var annualGrantRange = masterSheet.getRange(2, 8, lastRow - 1, 1);
    annualGrantRange.setNumberFormat('yyyy/mm/dd');

    // 4. データ検証を設定
    // F列（週所定労働日数）: 1-5の整数のみ
    var weekDaysValidation = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, 5)
      .setAllowInvalid(false)
      .setHelpText('週所定労働日数は1〜5の整数で入力してください')
      .build();
    weekDaysRange.setDataValidation(weekDaysValidation);

    // 5. 条件付き書式を設定（残有給日数の色分け）
    var rules = masterSheet.getConditionalFormatRules();

    // 残日数0日: 赤色
    var zeroRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setBackground('#FFCDD2')
      .setFontColor('#C62828')
      .setRanges([remainingRange])
      .build();

    // 残日数1-5日: オレンジ色
    var lowRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(1, 5)
      .setBackground('#FFE0B2')
      .setFontColor('#E65100')
      .setRanges([remainingRange])
      .build();

    // 残日数10日以上: 緑色
    var goodRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(10)
      .setBackground('#C8E6C9')
      .setFontColor('#2E7D32')
      .setRanges([remainingRange])
      .build();

    rules.push(zeroRule);
    rules.push(lowRule);
    rules.push(goodRule);
    masterSheet.setConditionalFormatRules(rules);

    // 6. 行の固定（ヘッダー行）
    masterSheet.setFrozenRows(1);

    console.log('=== マスターシートフォーマット設定完了 ===');

    return {
      success: true,
      message: 'マスターシートのフォーマットを設定しました',
      details: {
        columns: 8,
        headers: headers,
        formattedRows: lastRow - 1,
        validations: 'F列（週所定労働日数）: 1-5の整数',
        conditionalFormats: '残日数の色分け（0日:赤、1-5日:オレンジ、10日以上:緑）'
      }
    };

  } catch (error) {
    console.error('マスターシートフォーマット設定エラー:', error);
    return {
      success: false,
      error: error.message,
      message: 'マスターシートのフォーマット設定に失敗しました: ' + error.message
    };
  }
}

/**
 * 付与処理の対象者を抽出（統合版）
 */
function getGrantTargetEmployees() {
  try {
    console.log('=== 付与対象者抽出 ===');
    
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // 6ヶ月付与対象者
    var sixMonthTargets = getSixMonthGrantTargets();
    
    // 年次付与対象者（4月1日チェック）
    var annualTargets = [];
    if (today.getMonth() === 3 && today.getDate() === 1) { // 4月1日
      annualTargets = getAnnualGrantTargets();
    }
    
    return {
      success: true,
      sixMonthTargets: sixMonthTargets,
      annualTargets: annualTargets,
      summary: {
        sixMonthCount: sixMonthTargets.length,
        annualCount: annualTargets.length,
        totalCount: sixMonthTargets.length + annualTargets.length
      }
    };
    
  } catch (error) {
    console.error('付与対象者抽出エラー:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 年次付与対象者を抽出（4月1日用）
 */
function getAnnualGrantTargets() {
  try {
    console.log('=== 年次付与対象者抽出開始 ===');
    
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');
    
    if (!masterSheet) {
      throw new Error('マスターシートが見つかりません');
    }
    
    var data = masterSheet.getDataRange().getValues();
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    
    var targets = [];
    
    // ヘッダー行をスキップして処理
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      var userId = String(row[0]);
      var name = String(row[1]);
      var hireDateStr = row[4]; // E列: 入社日
      var weeklyWorkDaysStr = row[5]; // F列: 週所定労働日数
      var initialGrantDateStr = row[6]; // G列: 初回付与日
      var latestAnnualGrantDateStr = row[7]; // H列: 最新年次付与日

      if (!hireDateStr || !initialGrantDateStr) {
        continue; // 入社日がないか、初回付与がまだの場合は対象外
      }

      var hireDate = new Date(hireDateStr);
      var initialGrantDate = new Date(initialGrantDateStr);
      hireDate.setHours(0, 0, 0, 0);
      initialGrantDate.setHours(0, 0, 0, 0);

      // 勤続年数を計算（入社日からの年数）
      var workYears = calculateWorkYears(hireDate, today);

      // 初回付与を受けている場合は年次付与対象候補
      var weeklyWorkDays = Number(weeklyWorkDaysStr) || 5; // デフォルト5日

      // 次回付与予定日を計算
      var nextGrantDate;
      if (latestAnnualGrantDateStr) {
        // 最新年次付与日がある場合は、それに1年を加算
        var latestAnnualGrantDate = new Date(latestAnnualGrantDateStr);
        latestAnnualGrantDate.setHours(0, 0, 0, 0);
        nextGrantDate = new Date(latestAnnualGrantDate);
        nextGrantDate.setFullYear(nextGrantDate.getFullYear() + 1);
      } else {
        // 最新年次付与日がない場合は、初回付与日に1年を加算
        nextGrantDate = new Date(initialGrantDate);
        nextGrantDate.setFullYear(nextGrantDate.getFullYear() + 1);
      }

      // 次回付与日が今日以前の場合、付与対象
      if (nextGrantDate <= today) {
        var grantDays = calculateAnnualLeaveDays(Math.floor(workYears), weeklyWorkDays);

        targets.push({
          userId: userId,
          userName: name,  // nameではなくuserName
          hireDate: formatDate(hireDate),
          initialGrantDate: formatDate(initialGrantDate),
          latestAnnualGrantDate: latestAnnualGrantDateStr ? formatDate(new Date(latestAnnualGrantDateStr)) : null,
          workYears: Math.floor(workYears),  // 整数に丸める
          weeklyWorkDays: weeklyWorkDays,
          nextGrantDate: formatDate(nextGrantDate),
          grantDays: grantDays
        });

        console.log('年次付与対象:', userId, name, '次回付与日:' + Utilities.formatDate(nextGrantDate, 'JST', 'yyyy/MM/dd'));
      }
    }
    
    console.log('年次付与対象者:', targets.length + '名');
    return targets;
    
  } catch (error) {
    console.error('年次付与対象者抽出エラー:', error);
    console.error('エラー詳細:', error.message);
    console.error('スタックトレース:', error.stack);
    return [];
  }
}

// =============================
// 年次付与処理（4月1日一括付与）
// =============================

/**
 * 年次付与処理を実行（4月1日用）
 * @return {Object} 処理結果
 */
function processAnnualGrants() {
  try {
    console.log('=== 年次付与処理開始 ===');
    
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // 年次付与対象者を抽出
    var targets = getAnnualGrantTargets();
    
    if (targets.length === 0) {
      console.log('年次付与対象者はいません');
      return {
        success: true,
        processedCount: 0,
        message: '本日の年次付与対象者はいません'
      };
    }
    
    console.log('年次付与対象者:', targets.length + '名');
    
    var results = [];
    var successCount = 0;
    
    // 各対象者を処理
    targets.forEach(function(target) {
      try {
        console.log('年次付与処理:', target.userId, target.name);
        
        // 出勤率チェック（80%以上）
        var attendanceRate = checkAnnualAttendanceRate(target.userId, target.initialGrantDate, today);
        
        if (attendanceRate < 0.8) {
          console.log('出勤率不足により付与見送り:', target.userId, '出勤率:' + Math.round(attendanceRate * 100) + '%');
          results.push({
            userId: target.userId,
            name: target.name,
            success: false,
            reason: '出勤率不足',
            attendanceRate: Math.round(attendanceRate * 100) + '%'
          });
          return;
        }
        
        // 付与日数を計算（年次付与）
        var grantDays = calculateLeaveDays(target.workYears, target.weeklyWorkDays, false);
        if (grantDays <= 0) {
          console.log('付与テーブル外のためスキップ:', target.userId);
          results.push({
            userId: target.userId,
            name: target.name,
            success: false,
            reason: '付与対象外（テーブル外）'
          });
          return;
        }
        
        // 有給を付与
        var grantResult = grantLeave(
          target.userId,
          today, // 年次付与は4月1日に付与
          grantDays,
          '年次',
          target.workYears
        );
        
        if (grantResult.success) {
          // 年次付与履歴を記録
          recordAnnualGrantDate(target.userId, today, target.fiscalYear);
          
          // 通知送信
          sendAnnualGrantNotification(target, grantDays);
          
          successCount++;
          results.push({
            userId: target.userId,
            name: target.name,
            success: true,
            grantDays: grantDays,
            workYears: Math.floor(target.workYears),
            attendanceRate: Math.round(attendanceRate * 100) + '%'
          });
          
          console.log('年次付与完了:', target.userId, grantDays + '日付与');
        } else {
          results.push({
            userId: target.userId,
            name: target.name,
            success: false,
            reason: grantResult.error
          });
        }
        
      } catch (error) {
        console.error('年次付与処理エラー:', target.userId, error);
        results.push({
          userId: target.userId,
          name: target.name,
          success: false,
          reason: error.message
        });
      }
    });
    
    console.log('年次付与処理完了:', successCount + '/' + targets.length + '名成功');
    
    return {
      success: true,
      processedCount: successCount,
      totalTargets: targets.length,
      results: results,
      message: successCount + '/' + targets.length + '名に年次付与を実行しました'
    };
    
  } catch (error) {
    console.error('年次付与処理エラー:', error);
    return {
      success: false,
      error: error.message,
      message: '年次付与処理に失敗しました: ' + error.message
    };
  }
}

/**
 * 年次付与用の出勤率チェック（1年間）
 * @param {string} userId - 利用者番号
 * @param {Date} fromDate - 開始日（前回付与日）
 * @param {Date} toDate - 終了日（今日）
 * @return {number} 出勤率（0-1）
 */
function checkAnnualAttendanceRate(userId, fromDate, toDate) {
  try {
    var ss = getSpreadsheet();
    var applySheet = ss.getSheetByName('申請');
    
    if (!applySheet) {
      console.log('申請シートが見つからないため、出勤率100%とみなします');
      return 1.0;
    }
    
    var data = applySheet.getDataRange().getValues();
    var totalLeaveDays = 0;
    
    // 1年間の有給取得日数を計算
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      if (String(row[0]) === String(userId) && row[3] && row[5] === 'Approved') {
        var applyDate = new Date(row[3]);
        
        if (applyDate >= fromDate && applyDate <= toDate) {
          totalLeaveDays += Number(row[7] || 1);
        }
      }
    }
    
    // 1年間の稼働日数を計算
    var totalWorkDays = 0;
    var currentDate = new Date(fromDate);
    
    while (currentDate <= toDate) {
      var dayOfWeek = currentDate.getDay();
      if (dayOfWeek !== 0 && dayOfWeek !== 6) { // 土日以外
        totalWorkDays++;
      }
      currentDate.setDate(currentDate.getDate() + 1);
    }
    
    if (totalWorkDays === 0) {
      return 1.0;
    }
    
    var attendanceDays = totalWorkDays - totalLeaveDays;
    var attendanceRate = attendanceDays / totalWorkDays;
    
    console.log('年次出勤率計算:', {
      userId: userId,
      period: Utilities.formatDate(fromDate, 'JST', 'yyyy/MM/dd') + ' - ' + Utilities.formatDate(toDate, 'JST', 'yyyy/MM/dd'),
      totalWorkDays: totalWorkDays,
      totalLeaveDays: totalLeaveDays,
      attendanceDays: attendanceDays,
      attendanceRate: Math.round(attendanceRate * 100) + '%'
    });
    
    return Math.max(0, Math.min(1, attendanceRate));
    
  } catch (error) {
    console.error('年次出勤率チェックエラー:', error);
    return 1.0;
  }
}

/**
 * 年次付与日をマスターシートに記録
 * @param {string} userId - 利用者番号
 * @param {Date} grantDate - 付与日
 * @param {number} fiscalYear - 年度
 */
function recordAnnualGrantDate(userId, grantDate, fiscalYear) {
  try {
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');
    
    // H列に最新の年次付与日を記録
    var data = masterSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userId)) {
        // H列に年次付与日を記録
        if (masterSheet.getLastColumn() < 8) {
          // H列がない場合は追加
          masterSheet.getRange(1, 8).setValue('最新年次付与日');
        }
        masterSheet.getRange(i + 1, 8).setValue(grantDate);
        console.log('年次付与日記録:', userId, Utilities.formatDate(grantDate, 'JST', 'yyyy/MM/dd'));
        break;
      }
    }
    
  } catch (error) {
    console.error('年次付与日記録エラー:', error);
  }
}

/**
 * 年次付与通知を送信
 * @param {Object} target - 対象者情報
 * @param {number} grantDays - 付与日数
 */
function sendAnnualGrantNotification(target, grantDays) {
  try {
    var notificationData = {
      userId: target.userId,
      userName: target.name,
      grantDays: grantDays,
      grantDate: Utilities.formatDate(new Date(), 'JST', 'yyyy年M月d日'),
      grantType: '年次付与（' + Math.floor(target.workYears) + '年目）',
      workYears: Math.floor(target.workYears)
    };

    console.log('年次付与通知:', notificationData);

    // 実際の通知送信はテストモードなので省略
    // sendGrantNotification(notificationData);

  } catch (error) {
    console.error('年次付与通知エラー:', error);
  }
}

// =============================
// テストデータ生成機能
// =============================

/**
 * テストデータを生成してスプレッドシートに投入
 * @return {Object} 処理結果
 */
function generateTestData() {
  try {
    console.log('=== テストデータ生成開始 ===');

    var ss = getSpreadsheet();
    var results = {
      masterUsers: 0,
      grantRecords: 0,
      scenarios: []
    };

    // 1. マスターシートにテストユーザーを追加
    var masterResult = generateMasterTestData(ss);
    results.masterUsers = masterResult.count;

    // 2. 付与履歴シートにテストデータを追加
    var grantResult = generateGrantHistoryTestData(ss);
    results.grantRecords = grantResult.count;
    results.scenarios = grantResult.scenarios;

    // 3. マスターシートの残日数を更新
    updateMasterRemainingDays(ss);

    console.log('=== テストデータ生成完了 ===');

    return {
      success: true,
      message: 'テストデータを生成しました',
      details: results
    };

  } catch (error) {
    console.error('テストデータ生成エラー:', error);
    return {
      success: false,
      error: error.message,
      message: 'テストデータ生成に失敗しました: ' + error.message
    };
  }
}

/**
 * マスターシートにテストユーザーを追加
 */
function generateMasterTestData(ss) {
  var masterSheet = ss.getSheetByName('マスター');
  var today = new Date();

  var testUsers = [
    // 6ヶ月付与対象者（入社6ヶ月経過、初回付与前）
    {
      userId: 'TEST01',
      name: 'テスト太郎（6ヶ月経過）',
      remaining: 0,
      remarks: 'テストデータ：6ヶ月付与対象',
      hireDate: new Date(today.getFullYear(), today.getMonth() - 6, today.getDate()),
      weeklyWorkDays: 5,
      initialGrantDate: null
    },

    // 6ヶ月付与対象者（週3日勤務）
    {
      userId: 'TEST02',
      name: 'テスト花子（6ヶ月経過・週3日）',
      remaining: 0,
      remarks: 'テストデータ：6ヶ月付与対象（週3日）',
      hireDate: new Date(today.getFullYear(), today.getMonth() - 7, 15),
      weeklyWorkDays: 3,
      initialGrantDate: null
    },

    // 年次付与対象者（1年6ヶ月経過）
    {
      userId: 'TEST03',
      name: 'テスト次郎（1年6ヶ月）',
      remaining: 5,
      remarks: 'テストデータ：年次付与対象',
      hireDate: new Date(today.getFullYear() - 2, 3, 1),
      weeklyWorkDays: 5,
      initialGrantDate: new Date(today.getFullYear() - 1, 9, 1)
    },

    // 付与履歴あり（失効間近あり）
    {
      userId: 'TEST04',
      name: 'テスト三郎（失効間近）',
      remaining: 8,
      remarks: 'テストデータ：失効間近データあり',
      hireDate: new Date(today.getFullYear() - 3, 6, 1),
      weeklyWorkDays: 5,
      initialGrantDate: new Date(today.getFullYear() - 3, 11, 1)
    },

    // 付与履歴あり（失効済みデータあり）
    {
      userId: 'TEST05',
      name: 'テスト四郎（失効済みあり）',
      remaining: 12,
      remarks: 'テストデータ：失効済みデータあり',
      hireDate: new Date(today.getFullYear() - 4, 4, 1),
      weeklyWorkDays: 5,
      initialGrantDate: new Date(today.getFullYear() - 4, 9, 1)
    },

    // FIFO確認用（複数付与履歴）
    {
      userId: 'TEST06',
      name: 'テスト五郎（FIFO確認用）',
      remaining: 25,
      remarks: 'テストデータ：FIFO方式確認用',
      hireDate: new Date(today.getFullYear() - 5, 3, 1),
      weeklyWorkDays: 5,
      initialGrantDate: new Date(today.getFullYear() - 5, 8, 1)
    }
  ];

  testUsers.forEach(function(user) {
    masterSheet.appendRow([
      user.userId,
      user.name,
      user.remaining,
      user.remarks,
      user.hireDate,
      user.weeklyWorkDays,
      user.initialGrantDate,
      null // 最新年次付与日
    ]);
  });

  console.log('マスターシートにテストユーザーを追加:', testUsers.length + '名');

  return {
    count: testUsers.length,
    users: testUsers
  };
}

/**
 * 付与履歴シートにテストデータを追加
 */
function generateGrantHistoryTestData(ss) {
  var grantHistorySheet = getOrCreateGrantHistorySheet(ss);
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var testGrants = [];
  var scenarios = [];

  // TEST03: 年次付与対象者のデータ
  // 初回付与（1年6ヶ月前）
  var test03Grant1Date = new Date(today.getFullYear() - 1, 9, 1);
  var test03Grant1Expiry = new Date(test03Grant1Date);
  test03Grant1Expiry.setFullYear(test03Grant1Expiry.getFullYear() + 2);
  testGrants.push({
    userId: 'TEST03',
    grantDate: test03Grant1Date,
    grantDays: 10,
    expiryDate: test03Grant1Expiry,
    remainingDays: 5, // 一部消費済み
    grantType: '初回',
    workYears: 0.5,
    createdAt: test03Grant1Date
  });

  // TEST04: 失効間近データ
  // 古い付与（失効まであと5日）
  var test04Grant1Date = new Date(today);
  test04Grant1Date.setFullYear(test04Grant1Date.getFullYear() - 2);
  test04Grant1Date.setDate(test04Grant1Date.getDate() + 5); // 失効まで5日
  var test04Grant1Expiry = new Date(test04Grant1Date);
  test04Grant1Expiry.setFullYear(test04Grant1Expiry.getFullYear() + 2);
  testGrants.push({
    userId: 'TEST04',
    grantDate: test04Grant1Date,
    grantDays: 10,
    expiryDate: test04Grant1Expiry,
    remainingDays: 3, // 失効間近
    grantType: '初回',
    workYears: 0.5,
    createdAt: test04Grant1Date
  });
  scenarios.push('TEST04: 5日後に失効予定（残3日）');

  // 新しい付与
  var test04Grant2Date = new Date(today.getFullYear() - 1, 3, 1);
  var test04Grant2Expiry = new Date(test04Grant2Date);
  test04Grant2Expiry.setFullYear(test04Grant2Expiry.getFullYear() + 2);
  testGrants.push({
    userId: 'TEST04',
    grantDate: test04Grant2Date,
    grantDays: 11,
    expiryDate: test04Grant2Expiry,
    remainingDays: 5,
    grantType: '年次',
    workYears: 1.5,
    createdAt: test04Grant2Date
  });

  // TEST05: 失効済みデータ
  // 失効済み（失効日が3ヶ月前）
  var test05Grant1Date = new Date(today);
  test05Grant1Date.setFullYear(test05Grant1Date.getFullYear() - 2);
  test05Grant1Date.setMonth(test05Grant1Date.getMonth() - 3);
  var test05Grant1Expiry = new Date(test05Grant1Date);
  test05Grant1Expiry.setFullYear(test05Grant1Expiry.getFullYear() + 2);
  testGrants.push({
    userId: 'TEST05',
    grantDate: test05Grant1Date,
    grantDays: 10,
    expiryDate: test05Grant1Expiry,
    remainingDays: 0, // 失効済み
    grantType: '初回',
    workYears: 0.5,
    createdAt: test05Grant1Date
  });
  scenarios.push('TEST05: 失効済み（3ヶ月前に失効、残0日）');

  // 有効なデータ
  var test05Grant2Date = new Date(today.getFullYear() - 1, 3, 1);
  var test05Grant2Expiry = new Date(test05Grant2Date);
  test05Grant2Expiry.setFullYear(test05Grant2Expiry.getFullYear() + 2);
  testGrants.push({
    userId: 'TEST05',
    grantDate: test05Grant2Date,
    grantDays: 11,
    expiryDate: test05Grant2Expiry,
    remainingDays: 11,
    grantType: '年次',
    workYears: 1.5,
    createdAt: test05Grant2Date
  });

  var test05Grant3Date = new Date(today.getFullYear(), 3, 1);
  var test05Grant3Expiry = new Date(test05Grant3Date);
  test05Grant3Expiry.setFullYear(test05Grant3Expiry.getFullYear() + 2);
  testGrants.push({
    userId: 'TEST05',
    grantDate: test05Grant3Date,
    grantDays: 12,
    expiryDate: test05Grant3Expiry,
    remainingDays: 1, // ほぼ消費済み
    grantType: '年次',
    workYears: 2.5,
    createdAt: test05Grant3Date
  });

  // TEST06: FIFO確認用（複数付与、段階的に消費）
  // 1回目の付与（古い）
  var test06Grant1Date = new Date(today.getFullYear() - 4, 8, 1);
  var test06Grant1Expiry = new Date(test06Grant1Date);
  test06Grant1Expiry.setFullYear(test06Grant1Expiry.getFullYear() + 2);
  testGrants.push({
    userId: 'TEST06',
    grantDate: test06Grant1Date,
    grantDays: 10,
    expiryDate: test06Grant1Expiry,
    remainingDays: 3, // 一部消費
    grantType: '初回',
    workYears: 0.5,
    createdAt: test06Grant1Date
  });

  // 2回目の付与（1年後）
  var test06Grant2Date = new Date(today.getFullYear() - 3, 8, 1);
  var test06Grant2Expiry = new Date(test06Grant2Date);
  test06Grant2Expiry.setFullYear(test06Grant2Expiry.getFullYear() + 2);
  testGrants.push({
    userId: 'TEST06',
    grantDate: test06Grant2Date,
    grantDays: 11,
    expiryDate: test06Grant2Expiry,
    remainingDays: 6, // 一部消費
    grantType: '年次',
    workYears: 1.5,
    createdAt: test06Grant2Date
  });

  // 3回目の付与（2年後）
  var test06Grant3Date = new Date(today.getFullYear() - 2, 8, 1);
  var test06Grant3Expiry = new Date(test06Grant3Date);
  test06Grant3Expiry.setFullYear(test06Grant3Expiry.getFullYear() + 2);
  testGrants.push({
    userId: 'TEST06',
    grantDate: test06Grant3Date,
    grantDays: 12,
    expiryDate: test06Grant3Expiry,
    remainingDays: 12, // 未消費
    grantType: '年次',
    workYears: 2.5,
    createdAt: test06Grant3Date
  });

  // 4回目の付与（最新）
  var test06Grant4Date = new Date(today.getFullYear() - 1, 8, 1);
  var test06Grant4Expiry = new Date(test06Grant4Date);
  test06Grant4Expiry.setFullYear(test06Grant4Expiry.getFullYear() + 2);
  testGrants.push({
    userId: 'TEST06',
    grantDate: test06Grant4Date,
    grantDays: 14,
    expiryDate: test06Grant4Expiry,
    remainingDays: 4, // 一部消費
    grantType: '年次',
    workYears: 3.5,
    createdAt: test06Grant4Date
  });
  scenarios.push('TEST06: FIFO確認用（4回の付与、古い順に一部消費）');

  // 失効予定のバリエーションを追加
  // 1週間後に失効
  var weekExpiry = new Date(today);
  weekExpiry.setDate(weekExpiry.getDate() + 7);
  var weekGrant = new Date(weekExpiry);
  weekGrant.setFullYear(weekGrant.getFullYear() - 2);
  testGrants.push({
    userId: 'TEST04',
    grantDate: weekGrant,
    grantDays: 5,
    expiryDate: weekExpiry,
    remainingDays: 2,
    grantType: '特別付与',
    workYears: 2.0,
    createdAt: weekGrant
  });
  scenarios.push('TEST04: 1週間後に失効予定（残2日）');

  // 1ヶ月後に失効
  var monthExpiry = new Date(today);
  monthExpiry.setMonth(monthExpiry.getMonth() + 1);
  var monthGrant = new Date(monthExpiry);
  monthGrant.setFullYear(monthGrant.getFullYear() - 2);
  testGrants.push({
    userId: 'TEST05',
    grantDate: monthGrant,
    grantDays: 3,
    expiryDate: monthExpiry,
    remainingDays: 3,
    grantType: '特別付与',
    workYears: 2.5,
    createdAt: monthGrant
  });
  scenarios.push('TEST05: 1ヶ月後に失効予定（残3日）');

  // 3ヶ月後に失効
  var threeMonthExpiry = new Date(today);
  threeMonthExpiry.setMonth(threeMonthExpiry.getMonth() + 3);
  var threeMonthGrant = new Date(threeMonthExpiry);
  threeMonthGrant.setFullYear(threeMonthGrant.getFullYear() - 2);
  testGrants.push({
    userId: 'TEST06',
    grantDate: threeMonthGrant,
    grantDays: 5,
    expiryDate: threeMonthExpiry,
    remainingDays: 5,
    grantType: '特別付与',
    workYears: 3.0,
    createdAt: threeMonthGrant
  });
  scenarios.push('TEST06: 3ヶ月後に失効予定（残5日）');

  // データを投入
  testGrants.forEach(function(grant) {
    grantHistorySheet.appendRow([
      grant.userId,
      grant.grantDate,
      grant.grantDays,
      grant.expiryDate,
      grant.remainingDays,
      grant.grantType,
      grant.workYears,
      grant.createdAt
    ]);
  });

  console.log('付与履歴シートにテストデータを追加:', testGrants.length + '件');

  return {
    count: testGrants.length,
    scenarios: scenarios
  };
}

/**
 * マスターシートの残日数を更新（付与履歴から計算）
 */
function updateMasterRemainingDays(ss) {
  var masterSheet = ss.getSheetByName('マスター');
  var data = masterSheet.getDataRange().getValues();

  var testUserIds = ['TEST01', 'TEST02', 'TEST03', 'TEST04', 'TEST05', 'TEST06'];

  for (var i = 1; i < data.length; i++) {
    var userId = String(data[i][0]);

    if (testUserIds.indexOf(userId) !== -1) {
      var remaining = calculateEffectiveRemainingDays(userId);
      masterSheet.getRange(i + 1, 3).setValue(remaining);
      console.log('残日数更新:', userId, remaining + '日');
    }
  }
}

// =============================
// シートマニュアル機能
// =============================

/**
 * 各シートにマニュアルを追加
 * @return {Object} 処理結果
 */
function addSheetManuals() {
  try {
    console.log('=== シートマニュアル追加開始 ===');

    var ss = getSpreadsheet();
    var results = [];

    // マスターシートにマニュアルを追加
    var masterResult = addMasterSheetManual(ss);
    results.push(masterResult);

    // 付与履歴シートにマニュアルを追加
    var grantHistoryResult = addGrantHistorySheetManual(ss);
    results.push(grantHistoryResult);

    console.log('=== シートマニュアル追加完了 ===');

    return {
      success: true,
      results: results,
      message: 'シートマニュアルを追加しました'
    };

  } catch (error) {
    console.error('シートマニュアル追加エラー:', error);
    return {
      success: false,
      error: error.message,
      message: 'シートマニュアルの追加に失敗しました: ' + error.message
    };
  }
}

/**
 * マスターシートにマニュアルを追加
 * @param {Spreadsheet} ss - スプレッドシート
 * @return {Object} 処理結果
 */
function addMasterSheetManual(ss) {
  try {
    var masterSheet = ss.getSheetByName('マスター');

    if (!masterSheet) {
      throw new Error('マスターシートが見つかりません');
    }

    // J列（10列目）からマニュアルを記載
    var manualStartCol = 10;

    // マニュアルタイトル
    masterSheet.getRange(1, manualStartCol).setValue('【マスターシート 使い方マニュアル】');
    masterSheet.getRange(1, manualStartCol).setFontWeight('bold');
    masterSheet.getRange(1, manualStartCol).setBackground('#2196F3');
    masterSheet.getRange(1, manualStartCol).setFontColor('white');
    masterSheet.setColumnWidth(manualStartCol, 300);

    // マニュアル内容
    var manualContent = [
      [''],
      ['■ このシートの役割'],
      ['利用者の基本情報と有給残日数を管理します。'],
      [''],
      ['■ 列の説明'],
      ['A列: 利用者番号（例: R01, P01, S01, E01）'],
      ['B列: 利用者名'],
      ['C列: 残有給日数（自動計算・手動修正可）'],
      ['D列: 備考（任意のメモ）'],
      ['E列: 入社日（付与計算に使用）'],
      ['F列: 週所定労働日数（1-5、付与日数の計算に使用）'],
      ['G列: 初回付与日（入社6ヶ月後に自動記録）'],
      ['H列: 最新年次付与日（4月1日の年次付与時に自動記録）'],
      [''],
      ['■ 手動修正が必要な場合'],
      ['【パターン1: 残日数のみ修正】'],
      ['  C列の残日数を直接編集してください。'],
      ['  翌日の日次処理で整合性チェックが行われます。'],
      [''],
      ['【パターン2: 厳密に管理する場合】'],
      ['  1. C列の残日数を修正'],
      ['  2. 付与履歴シートに修正記録を追加'],
      ['  （利用者番号、付与日、付与日数、失効日、残日数、付与タイプ=「手動調整」等）'],
      [''],
      ['■ 新規利用者の追加手順'],
      ['1. 利用者番号を設定（事業所コード+連番）'],
      ['2. 利用者名を入力'],
      ['3. 残有給日数は0または初期値'],
      ['4. 入社日を入力（必須）'],
      ['5. 週所定労働日数を入力（1-5、デフォルト5）'],
      ['6. G列・H列は自動処理で記録されるため空欄でOK'],
      [''],
      ['■ 注意事項'],
      ['・利用者番号は変更しないでください（システム全体で使用）'],
      ['・入社日が正しくないと付与処理が正しく動作しません'],
      ['・残日数の手動修正は慎重に行ってください'],
      ['・退職者は削除せず、備考欄に「退職」と記入してください'],
      [''],
      ['■ 自動処理について'],
      ['・毎日10時: 6ヶ月付与チェック（入社6ヶ月後の自動付与）'],
      ['・毎年4月1日 8時: 年次付与処理（勤続年数に応じた付与）'],
      ['・毎日9時30分: 失効処理（有効期限切れの自動失効）'],
      ['・毎日9時: データ整合性チェック'],
      ['']
    ];

    // マニュアルを書き込み
    for (var i = 0; i < manualContent.length; i++) {
      var cell = masterSheet.getRange(i + 2, manualStartCol);
      cell.setValue(manualContent[i][0]);
      cell.setWrap(true);

      // 見出し行の書式設定
      if (manualContent[i][0].indexOf('■') === 0) {
        cell.setFontWeight('bold');
        cell.setBackground('#E3F2FD');
        cell.setFontColor('#1976D2');
      } else if (manualContent[i][0].indexOf('【') === 0) {
        cell.setFontWeight('bold');
        cell.setFontColor('#FF6F00');
      }
    }

    console.log('マスターシートにマニュアルを追加しました');

    return {
      success: true,
      sheet: 'マスター',
      message: 'マスターシートにマニュアルを追加しました'
    };

  } catch (error) {
    console.error('マスターシートマニュアル追加エラー:', error);
    return {
      success: false,
      sheet: 'マスター',
      error: error.message
    };
  }
}

/**
 * 付与履歴シートにマニュアルを追加
 * @param {Spreadsheet} ss - スプレッドシート
 * @return {Object} 処理結果
 */
function addGrantHistorySheetManual(ss) {
  try {
    var grantHistorySheet = getOrCreateGrantHistorySheet(ss);

    // J列（10列目）からマニュアルを記載
    var manualStartCol = 10;

    // マニュアルタイトル
    grantHistorySheet.getRange(1, manualStartCol).setValue('【付与履歴シート 使い方マニュアル】');
    grantHistorySheet.getRange(1, manualStartCol).setFontWeight('bold');
    grantHistorySheet.getRange(1, manualStartCol).setBackground('#4CAF50');
    grantHistorySheet.getRange(1, manualStartCol).setFontColor('white');
    grantHistorySheet.setColumnWidth(manualStartCol, 300);

    // マニュアル内容
    var manualContent = [
      [''],
      ['■ このシートの役割'],
      ['有給の付与履歴を記録し、FIFO方式（先入先出）で消費を管理します。'],
      ['失効処理もこのシートで管理されます。'],
      [''],
      ['■ 列の説明'],
      ['A列: 利用者番号'],
      ['B列: 付与日（有給が付与された日）'],
      ['C列: 付与日数（付与された日数）'],
      ['D列: 失効日（付与日から2年後）'],
      ['E列: 残日数（この付与分の残り日数）'],
      ['F列: 付与タイプ（初回/年次/手動調整など）'],
      ['G列: 勤続年数（付与時点の勤続年数）'],
      ['H列: 作成日時（レコード作成日時）'],
      [''],
      ['■ データの見方'],
      ['【有効な有給】'],
      ['  失効日 > 今日 かつ 残日数 > 0'],
      [''],
      ['【失効済みの有給】'],
      ['  失効日 <= 今日 または 残日数 = 0'],
      ['  ※失効処理により残日数が0に更新されます'],
      [''],
      ['【消費済みの有給】'],
      ['  残日数 = 0（ただし失効日前）'],
      [''],
      ['■ 手動で修正する場合'],
      ['【ケース1: 付与日数を追加したい】'],
      ['  新しい行を追加して以下を入力：'],
      ['  - 利用者番号'],
      ['  - 付与日（今日または任意の日付）'],
      ['  - 付与日数'],
      ['  - 失効日（付与日から2年後を計算）'],
      ['  - 残日数（初期値は付与日数と同じ）'],
      ['  - 付与タイプ（「手動調整」など）'],
      ['  - 勤続年数（参考値）'],
      ['  - 作成日時（今日の日時）'],
      ['  ※追加後、マスターシートのC列も手動で更新してください'],
      [''],
      ['【ケース2: 誤って消費された分を戻したい】'],
      ['  該当行のE列（残日数）を修正'],
      ['  ※マスターシートのC列も手動で更新してください'],
      [''],
      ['【ケース3: 特例で失効を取り消したい】'],
      ['  1. E列（残日数）を元の値に戻す'],
      ['  2. 必要ならD列（失効日）を延長'],
      ['  ※マスターシートのC列も手動で更新してください'],
      [''],
      ['■ FIFO方式（先入先出）について'],
      ['有給を使用する際、付与日が古いものから順に消費されます。'],
      ['これにより、失効リスクの高い有給から優先的に使用されます。'],
      [''],
      ['例:'],
      ['  2023/04/01 付与 10日 残5日'],
      ['  2024/04/01 付与 11日 残11日'],
      ['  → 3日使用すると、2023年分から消費され残2日になります'],
      [''],
      ['■ 自動処理との連携'],
      ['・有給申請が承認されると、このシートの残日数が自動的に減ります'],
      ['・失効処理（毎日9:30）で失効日を過ぎた行の残日数が0になります'],
      ['・付与処理（6ヶ月/年次）で新しい行が自動追加されます'],
      [''],
      ['■ 注意事項'],
      ['・このシートを直接編集した場合、マスターシートも手動更新が必要です'],
      ['・削除は避け、修正する場合は慎重に行ってください'],
      ['・過去のデータは監査証跡として保持することを推奨します'],
      ['・整合性チェック（毎日9時）で不整合があればログに記録されます'],
      ['']
    ];

    // マニュアルを書き込み
    for (var i = 0; i < manualContent.length; i++) {
      var cell = grantHistorySheet.getRange(i + 2, manualStartCol);
      cell.setValue(manualContent[i][0]);
      cell.setWrap(true);

      // 見出し行の書式設定
      if (manualContent[i][0].indexOf('■') === 0) {
        cell.setFontWeight('bold');
        cell.setBackground('#C8E6C9');
        cell.setFontColor('#2E7D32');
      } else if (manualContent[i][0].indexOf('【') === 0) {
        cell.setFontWeight('bold');
        cell.setFontColor('#F57C00');
      }
    }

    console.log('付与履歴シートにマニュアルを追加しました');

    return {
      success: true,
      sheet: '付与履歴',
      message: '付与履歴シートにマニュアルを追加しました'
    };

  } catch (error) {
    console.error('付与履歴シートマニュアル追加エラー:', error);
    return {
      success: false,
      sheet: '付与履歴',
      error: error.message
    };
  }
}
