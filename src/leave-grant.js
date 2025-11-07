// =============================
// 有給付与システム
// =============================

/**
 * 有給付与日数を計算
 * @param {number} workYears - 勤続年数（年単位）
 * @param {number} weeklyWorkDays - 週所定労働日数
 * @param {boolean} isInitial - 初回付与かどうか
 * @return {number} 付与日数
 */
function calculateLeaveDays(workYears, weeklyWorkDays, isInitial) {
  try {
    var years = Math.floor(workYears);
    var weekDays = Math.floor(weeklyWorkDays);
    
    console.log('付与日数計算:', {
      workYears: workYears,
      years: years,
      weeklyWorkDays: weeklyWorkDays,
      weekDays: weekDays,
      isInitial: isInitial
    });
    
    if (isInitial) {
      return calculateInitialLeaveDays(weekDays);
    }
    
    return calculateAnnualLeaveDays(years, weekDays);
    
  } catch (error) {
    console.error('付与日数計算エラー:', error);
    return 0;
  }
}

/**
 * 初回付与日数を計算（入社6ヶ月後）
 */
function calculateInitialLeaveDays(weeklyWorkDays) {
  var weekDays = Math.floor(weeklyWorkDays);
  
  // 基本設計書ベースの初回付与日数
  if (weekDays >= 5) return 10;
  if (weekDays === 4) return 7;
  if (weekDays === 3) return 5;
  if (weekDays === 2) return 3;
  if (weekDays === 1) return 1;
  
  return 0;
}

/**
 * 年次付与日数を計算（2回目以降）
 */
function calculateAnnualLeaveDays(workYears, weeklyWorkDays) {
  var years = Math.floor(workYears);
  var weekDays = Math.floor(weeklyWorkDays);
  
  // 週5日以上勤務者
  if (weekDays >= 5) {
    if (years < 1) return 10;
    if (years < 2) return 11;
    if (years < 3) return 12;
    if (years < 4) return 14;
    if (years < 5) return 16;
    if (years < 6) return 18;
    return 20;
  }
  
  // 週4日勤務者
  if (weekDays === 4) {
    if (years < 1) return 7;
    if (years < 2) return 8;
    if (years < 3) return 9;
    if (years < 4) return 10;
    if (years < 5) return 12;
    if (years < 6) return 13;
    return 15;
  }
  
  // 週3日勤務者
  if (weekDays === 3) {
    if (years < 1) return 5;
    if (years < 2) return 6;
    if (years < 3) return 6;
    if (years < 4) return 8;
    if (years < 5) return 9;
    if (years < 6) return 10;
    return 11;
  }
  
  // 週2日勤務者
  if (weekDays === 2) {
    if (years < 1) return 3;
    if (years < 2) return 4;
    if (years < 3) return 4;
    if (years < 4) return 5;
    if (years < 5) return 6;
    if (years < 6) return 6;
    return 7;
  }
  
  // 週1日勤務者
  if (weekDays === 1) {
    if (years < 1) return 1;
    if (years < 2) return 2;
    if (years < 3) return 2;
    if (years < 4) return 2;
    if (years < 5) return 3;
    if (years < 6) return 3;
    return 3;
  }
  
  return 0;
}

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

/**
 * 勤続年数を計算
 * @param {Date} hireDate - 入社日
 * @param {Date} baseDate - 基準日（付与日や今日）
 * @return {number} 勤続年数（小数点仕切き含む）
 */
function calculateWorkYears(hireDate, baseDate) {
  if (!hireDate || !baseDate) return 0;
  
  var hire = new Date(hireDate);
  var base = new Date(baseDate);
  
  hire.setHours(0, 0, 0, 0);
  base.setHours(0, 0, 0, 0);
  
  // 日数差を年数に変換
  var diffDays = Math.floor((base - hire) / (1000 * 60 * 60 * 24));
  var workYears = diffDays / 365.25; // うるう年を考慮
  
  return Math.max(0, workYears);
}

/**
 * 年5日取得義務の対象者かどうかを判定
 * @param {number} grantedDays - 付与日数
 * @return {boolean} 年5日取得義務対象かどうか
 */
function isFiveDayObligationTarget(grantedDays) {
  return grantedDays >= 10;
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
 * FIFO方式で有給を消費
 * @param {string} userId - 利用者番号
 * @param {number} useDays - 使用日数
 * @return {Object} 消費結果
 */
function consumeLeave(userId, useDays) {
  try {
    console.log('有給消費処理開始:', userId, useDays);
    
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
        
        targets.push({
          userId: userId,
          name: name,
          hireDate: hireDate,
          sixMonthDate: sixMonthDate,
          weeklyWorkDays: weeklyWorkDays
        });
        
        console.log('6ヶ月付与対象:', userId, name, '6ヶ月経過日:' + Utilities.formatDate(sixMonthDate, 'JST', 'yyyy/MM/dd'));
      }
    }
    
    return targets;
    
  } catch (error) {
    console.error('6ヶ月付与対象者抽出エラー:', error);
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
      'G: 初回付与日'
    ];
    
    if (lastColumn < 7) {
      console.log('マスターシートに列を追加します');
      
      // ヘッダー行を更新
      var headerRange = masterSheet.getRange(1, 1, 1, 7);
      headerRange.setValues([
        ['利用者番号', '利用者名', '残有給日数', '備考', '入社日', '週所定労働日数', '初回付与日']
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
      
      if (!hireDateStr || !initialGrantDateStr) {
        continue; // 入社日がないか、初回付与がまだの場合は対象外
      }
      
      var hireDate = new Date(hireDateStr);
      var initialGrantDate = new Date(initialGrantDateStr);
      hireDate.setHours(0, 0, 0, 0);
      initialGrantDate.setHours(0, 0, 0, 0);
      
      // 勤続年数を計算（初回付与日からの年数）
      var workYears = calculateWorkYears(initialGrantDate, today);
      
      // 1年以上経過している場合は年次付与対象
      if (workYears >= 1) {
        var weeklyWorkDays = Number(weeklyWorkDaysStr) || 5; // デフォルト5日
        
        // 次回付与予定日を計算（初回付与日から1年ごと）
        var nextGrantYear = Math.floor(workYears) + 1;
        var nextGrantDate = new Date(initialGrantDate);
        nextGrantDate.setFullYear(nextGrantDate.getFullYear() + nextGrantYear);
        
        // 今年の4月1日以降で、まだ今年度の付与を受けていない場合
        var currentFiscalYear = today.getFullYear();
        if (today.getMonth() < 3) { // 1-3月の場合は前年度
          currentFiscalYear--;
        }
        
        var fiscalYearStart = new Date(currentFiscalYear, 3, 1); // 4月1日
        
        // 今年度に付与すべきかチェック
        if (nextGrantDate <= today && initialGrantDate < fiscalYearStart) {
          targets.push({
            userId: userId,
            name: name,
            hireDate: hireDate,
            initialGrantDate: initialGrantDate,
            workYears: workYears,
            weeklyWorkDays: weeklyWorkDays,
            nextGrantDate: nextGrantDate,
            fiscalYear: currentFiscalYear
          });
          
          console.log('年次付与対象:', userId, name, '勤続年数:' + Math.floor(workYears) + '年');
        }
      }
    }
    
    console.log('年次付与対象者:', targets.length + '名');
    return targets;
    
  } catch (error) {
    console.error('年次付与対象者抽出エラー:', error);
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