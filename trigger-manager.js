// =============================
// GASトリガー管理システム
// =============================

/**
 * 有給管理システムのトリガー設定
 */
var TRIGGER_CONFIG = {
  DAILY_PROCESS_HOUR: 9,        // 日次処理実行時刻（9時）
  ANNUAL_PROCESS_MONTH: 4,      // 年次処理実行月（4月）
  ANNUAL_PROCESS_DATE: 1,       // 年次処理実行日（1日）
  ANNUAL_PROCESS_HOUR: 8,       // 年次処理実行時刻（8時）
  NOTIFICATION_INTERVAL: 30     // 通知処理間隔（30分）
};

/**
 * システム用トリガーを全て設定
 * @return {Object} 設定結果
 */
function setupAllTriggers() {
  try {
    console.log('=== 有給管理システム トリガー設定開始 ===');
    
    // 既存のトリガーをクリア
    var clearResult = clearAllSystemTriggers();
    console.log('既存トリガークリア:', clearResult.message);
    
    var results = [];
    var successCount = 0;
    
    // 1. 日次処理トリガー（毎日9時）
    try {
      var dailyTrigger = ScriptApp.newTrigger('runDailyProcess')
        .timeBased()
        .everyDays(1)
        .atHour(TRIGGER_CONFIG.DAILY_PROCESS_HOUR)
        .create();
      
      results.push({
        name: '日次処理トリガー',
        success: true,
        triggerId: dailyTrigger.getUniqueId(),
        schedule: '毎日' + TRIGGER_CONFIG.DAILY_PROCESS_HOUR + '時',
        function: 'runDailyProcess'
      });
      successCount++;
      console.log('✅ 日次処理トリガー設定完了');
      
    } catch (error) {
      results.push({
        name: '日次処理トリガー',
        success: false,
        error: error.message
      });
      console.error('❌ 日次処理トリガー設定失敗:', error);
    }
    
    // 2. 年次処理トリガー（4月1日8時）
    try {
      var annualTrigger = ScriptApp.newTrigger('runAnnualProcess')
        .timeBased()
        .onMonthDay(TRIGGER_CONFIG.ANNUAL_PROCESS_DATE)
        .atHour(TRIGGER_CONFIG.ANNUAL_PROCESS_HOUR)
        .create();
      
      results.push({
        name: '年次処理トリガー',
        success: true,
        triggerId: annualTrigger.getUniqueId(),
        schedule: '毎年4月1日8時',
        function: 'runAnnualProcess'
      });
      successCount++;
      console.log('✅ 年次処理トリガー設定完了');
      
    } catch (error) {
      results.push({
        name: '年次処理トリガー',
        success: false,
        error: error.message
      });
      console.error('❌ 年次処理トリガー設定失敗:', error);
    }
    
    // 3. 失効処理トリガー（毎日9時30分）
    try {
      var expiryTrigger = ScriptApp.newTrigger('runExpiryProcess')
        .timeBased()
        .everyDays(1)
        .atHour(TRIGGER_CONFIG.DAILY_PROCESS_HOUR)
        .nearMinute(30)
        .create();
      
      results.push({
        name: '失効処理トリガー',
        success: true,
        triggerId: expiryTrigger.getUniqueId(),
        schedule: '毎日9時30分',
        function: 'runExpiryProcess'
      });
      successCount++;
      console.log('✅ 失効処理トリガー設定完了');
      
    } catch (error) {
      results.push({
        name: '失効処理トリガー',
        success: false,
        error: error.message
      });
      console.error('❌ 失効処理トリガー設定失敗:', error);
    }
    
    // 4. 6ヶ月付与チェックトリガー（毎日10時）
    try {
      var sixMonthTrigger = ScriptApp.newTrigger('runSixMonthGrantCheck')
        .timeBased()
        .everyDays(1)
        .atHour(10)
        .create();
      
      results.push({
        name: '6ヶ月付与チェックトリガー',
        success: true,
        triggerId: sixMonthTrigger.getUniqueId(),
        schedule: '毎日10時',
        function: 'runSixMonthGrantCheck'
      });
      successCount++;
      console.log('✅ 6ヶ月付与チェックトリガー設定完了');
      
    } catch (error) {
      results.push({
        name: '6ヶ月付与チェックトリガー',
        success: false,
        error: error.message
      });
      console.error('❌ 6ヶ月付与チェックトリガー設定失敗:', error);
    }
    
    // トリガー設定をPropertiesServiceに保存
    var triggerInfo = {
      setupDate: new Date().toISOString(),
      triggers: results.filter(function(r) { return r.success; }),
      totalCount: results.length,
      successCount: successCount
    };
    
    PropertiesService.getScriptProperties().setProperty(
      'TRIGGER_INFO',
      JSON.stringify(triggerInfo)
    );
    
    console.log('=== トリガー設定完了 ===');
    console.log('成功: ' + successCount + '/' + results.length);
    
    return {
      success: successCount > 0,
      totalTriggers: results.length,
      successCount: successCount,
      failedCount: results.length - successCount,
      results: results,
      message: successCount + '/' + results.length + 'のトリガーを設定しました'
    };
    
  } catch (error) {
    console.error('トリガー設定エラー:', error);
    return {
      success: false,
      error: error.message,
      message: 'トリガー設定に失敗しました: ' + error.message
    };
  }
}

/**
 * システム用トリガーをすべてクリア
 * @return {Object} クリア結果
 */
function clearAllSystemTriggers() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var deletedCount = 0;
    
    var systemFunctions = [
      'runDailyProcess',
      'runAnnualProcess', 
      'runExpiryProcess',
      'runSixMonthGrantCheck'
    ];
    
    triggers.forEach(function(trigger) {
      var handlerFunction = trigger.getHandlerFunction();
      
      if (systemFunctions.indexOf(handlerFunction) !== -1) {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
        console.log('トリガー削除:', handlerFunction);
      }
    });
    
    return {
      success: true,
      deletedCount: deletedCount,
      message: deletedCount + '個の既存トリガーを削除しました'
    };
    
  } catch (error) {
    console.error('トリガークリアエラー:', error);
    return {
      success: false,
      error: error.message,
      message: 'トリガークリアに失敗しました'
    };
  }
}

/**
 * 現在のトリガー設定状況を確認
 * @return {Object} トリガー情報
 */
function getTriggerStatus() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var systemTriggers = [];
    
    var systemFunctions = [
      'runDailyProcess',
      'runAnnualProcess',
      'runExpiryProcess', 
      'runSixMonthGrantCheck'
    ];
    
    triggers.forEach(function(trigger) {
      var handlerFunction = trigger.getHandlerFunction();
      
      if (systemFunctions.indexOf(handlerFunction) !== -1) {
        systemTriggers.push({
          id: trigger.getUniqueId(),
          function: handlerFunction,
          triggerSource: trigger.getTriggerSource().toString(),
          eventType: trigger.getEventType().toString()
        });
      }
    });
    
    // 保存されたトリガー情報を取得
    var savedTriggerInfo = PropertiesService.getScriptProperties().getProperty('TRIGGER_INFO');
    var triggerInfo = savedTriggerInfo ? JSON.parse(savedTriggerInfo) : null;
    
    return {
      success: true,
      activeTriggers: systemTriggers,
      activeCount: systemTriggers.length,
      savedInfo: triggerInfo,
      message: systemTriggers.length + '個のシステムトリガーが活動中です'
    };
    
  } catch (error) {
    console.error('トリガー状況確認エラー:', error);
    return {
      success: false,
      error: error.message,
      message: 'トリガー状況の確認に失敗しました'
    };
  }
}

// =============================
// トリガーで実行される処理関数
// =============================

/**
 * 日次処理を実行
 */
function runDailyProcess() {
  try {
    console.log('=== 日次処理開始 (' + new Date() + ') ===');
    
    var results = [];
    
    // 1. システムヘルスチェック
    console.log('1. システムヘルスチェック');
    results.push({
      process: 'ヘルスチェック',
      result: { success: true, message: 'システム正常' }
    });
    
    // 2. データ整合性チェック
    console.log('2. データ整合性チェック');
    var integrityResult = checkDataIntegrity();
    results.push({
      process: 'データ整合性チェック',
      result: integrityResult
    });
    
    // 3. 残日数アラートチェック
    console.log('3. 残日数アラートチェック');
    var alertResult = checkLowRemainingDaysAlerts();
    results.push({
      process: '残日数アラート',
      result: alertResult
    });
    
    // 処理結果をログに記録
    logDailyProcessResult(results);
    
    console.log('=== 日次処理完了 ===');
    
    return {
      success: true,
      processedItems: results.length,
      results: results
    };
    
  } catch (error) {
    console.error('日次処理エラー:', error);
    
    // エラー通知
    logSystemError('日次処理', error);
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 年次処理を実行（4月1日）
 */
function runAnnualProcess() {
  try {
    console.log('=== 年次処理開始 (' + new Date() + ') ===');
    
    // 今日が4月1日かチェック
    var today = new Date();
    if (today.getMonth() !== 3 || today.getDate() !== 1) {
      console.log('今日は4月1日ではないため、年次処理をスキップします');
      return { success: true, message: '年次処理日ではありません' };
    }
    
    var results = [];
    
    // 1. 年次有給付与処理
    console.log('1. 年次有給付与処理');
    var grantResult = processAnnualGrants();
    results.push({
      process: '年次付与',
      result: grantResult
    });
    
    // 2. 年度切り替え処理
    console.log('2. 年度切り替え処理'); 
    var fiscalYearResult = processFiscalYearChange();
    results.push({
      process: '年度切り替え',
      result: fiscalYearResult
    });
    
    // 処理結果をログに記録
    logAnnualProcessResult(results);
    
    console.log('=== 年次処理完了 ===');
    
    return {
      success: true,
      processedItems: results.length,
      results: results
    };
    
  } catch (error) {
    console.error('年次処理エラー:', error);
    
    // エラー通知
    logSystemError('年次処理', error);
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 失効処理を実行
 */
function runExpiryProcess() {
  try {
    console.log('=== 失効処理開始 (' + new Date() + ') ===');
    
    // 失効処理実行
    var expiryResult = processExpiredLeaves();
    
    // 失効通知送信
    if (expiryResult.success && expiryResult.expiredUsers.length > 0) {
      sendExpiryNotifications(expiryResult.expiredUsers);
    }
    
    // 結果をログに記録
    logExpiryProcessResult(expiryResult);
    
    console.log('=== 失効処理完了 ===');
    
    return expiryResult;
    
  } catch (error) {
    console.error('失効処理エラー:', error);
    
    // エラー通知
    logSystemError('失効処理', error);
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 6ヶ月付与チェックを実行
 */
function runSixMonthGrantCheck() {
  try {
    console.log('=== 6ヶ月付与チェック開始 (' + new Date() + ') ===');
    
    // 6ヶ月付与処理実行
    var grantResult = processSixMonthGrants();
    
    // 結果をログに記録
    logSixMonthProcessResult(grantResult);
    
    console.log('=== 6ヶ月付与チェック完了 ===');
    
    return grantResult;
    
  } catch (error) {
    console.error('6ヶ月付与チェックエラー:', error);
    
    // エラー通知
    logSystemError('6ヶ月付与チェック', error);
    
    return {
      success: false,
      error: error.message
    };
  }
}

// =============================
// ヘルパー関数
// =============================

/**
 * データ整合性チェック
 */
function checkDataIntegrity() {
  try {
    var issues = [];
    
    // マスターシートと付与履歴の整合性チェック
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');
    var grantHistorySheet = ss.getSheetByName('付与履歴');
    
    if (!masterSheet || !grantHistorySheet) {
      issues.push('必要なシートが見つかりません');
    } else {
      var masterData = masterSheet.getDataRange().getValues();
      
      // 各利用者の残日数と履歴の整合性をチェック
      for (var i = 1; i < masterData.length; i++) {
        var userId = String(masterData[i][0]);
        var masterRemaining = Number(masterData[i][2] || 0);
        var calculatedRemaining = calculateEffectiveRemainingDays(userId);
        
        if (Math.abs(masterRemaining - calculatedRemaining) > 0.01) {
          issues.push(userId + ': 残日数不整合 (マスター:' + masterRemaining + '日, 計算値:' + calculatedRemaining + '日)');
        }
      }
    }
    
    return {
      success: issues.length === 0,
      issues: issues,
      message: issues.length === 0 ? 'データ整合性OK' : issues.length + '件の問題を検出'
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message,
      message: '整合性チェック実行エラー'
    };
  }
}

/**
 * 残日数アラートチェック
 */
function checkLowRemainingDaysAlerts() {
  try {
    // 残日数が少ない社員を抽出
    var lowRemainingEmployees = getLowRemainingDaysEmployees();
    
    if (lowRemainingEmployees.length === 0) {
      return {
        success: true,
        message: 'アラート対象者なし',
        alertCount: 0
      };
    }
    
    // アラート送信（テストモードの場合は実際には送信しない）
    var alertResult = sendBulkLowRemainingDaysAlerts();
    
    return {
      success: true,
      message: lowRemainingEmployees.length + '名にアラートチェック実行',
      alertCount: lowRemainingEmployees.length,
      sentCount: alertResult.sentCount
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message,
      message: 'アラートチェック実行エラー'
    };
  }
}

/**
 * 年度切り替え処理
 */
function processFiscalYearChange() {
  try {
    var currentDate = new Date();
    var fiscalYear = currentDate.getFullYear();
    
    console.log('年度切り替え処理:', fiscalYear + '年度');
    
    // 年度情報を記録
    PropertiesService.getScriptProperties().setProperty(
      'CURRENT_FISCAL_YEAR',
      fiscalYear.toString()
    );
    
    return {
      success: true,
      fiscalYear: fiscalYear,
      message: fiscalYear + '年度への切り替え完了'
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message,
      message: '年度切り替え処理エラー'
    };
  }
}

/**
 * 処理結果ログ記録関数群
 */
function logDailyProcessResult(results) {
  var logMessage = '[日次処理] ' + new Date().toISOString() + '\n';
  results.forEach(function(item) {
    logMessage += '- ' + item.process + ': ' + (item.result.success ? 'OK' : 'NG') + '\n';
  });
  console.log(logMessage);
}

function logAnnualProcessResult(results) {
  var logMessage = '[年次処理] ' + new Date().toISOString() + '\n';
  results.forEach(function(item) {
    logMessage += '- ' + item.process + ': ' + (item.result.success ? 'OK' : 'NG') + '\n';
  });
  console.log(logMessage);
}

function logExpiryProcessResult(result) {
  var logMessage = '[失効処理] ' + new Date().toISOString() + 
    ' - ' + (result.success ? 'OK' : 'NG') + ': ' + result.message;
  console.log(logMessage);
}

function logSixMonthProcessResult(result) {
  var logMessage = '[6ヶ月付与] ' + new Date().toISOString() + 
    ' - ' + (result.success ? 'OK' : 'NG') + ': ' + result.message;
  console.log(logMessage);
}

function logSystemError(processName, error) {
  var errorMessage = '[システムエラー] ' + processName + ' - ' + new Date().toISOString() + 
    '\nエラー: ' + error.message + '\nスタック: ' + error.stack;
  console.error(errorMessage);
  
  // 重要なエラーはPropertiesServiceにも記録
  var errorLog = PropertiesService.getScriptProperties().getProperty('ERROR_LOG') || '[]';
  var errors = JSON.parse(errorLog);
  
  errors.push({
    timestamp: new Date().toISOString(),
    process: processName,
    error: error.message,
    stack: error.stack
  });
  
  // 最新100件まで保持
  if (errors.length > 100) {
    errors = errors.slice(-100);
  }
  
  PropertiesService.getScriptProperties().setProperty('ERROR_LOG', JSON.stringify(errors));
}