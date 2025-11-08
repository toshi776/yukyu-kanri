/**
 * 付与管理機能のデバッグスクリプト
 */

/**
 * 6ヶ月付与対象者のデバッグ
 */
function debugSixMonthTargets() {
  console.log('=== 6ヶ月付与対象者デバッグ ===');

  try {
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');
    var data = masterSheet.getDataRange().getValues();
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    console.log('今日:', Utilities.formatDate(today, 'JST', 'yyyy/MM/dd'));
    console.log('マスターシート総行数:', data.length);
    console.log('');

    // TEST01とTEST02だけ詳細チェック
    for (var i = 1; i < data.length; i++) {
      var userId = String(data[i][0]);
      if (userId === 'TEST01' || userId === 'TEST02') {
        var name = String(data[i][1]);
        var hireDateStr = data[i][4]; // E列: 入社日
        var weeklyWorkDaysStr = data[i][5]; // F列: 週所定労働日数
        var initialGrantDateStr = data[i][6]; // G列: 初回付与日

        console.log(userId + ' (' + name + '):');
        console.log('  入社日(raw):', hireDateStr);
        console.log('  初回付与日(raw):', initialGrantDateStr);
        console.log('  初回付与日(type):', typeof initialGrantDateStr);
        console.log('  初回付与日(bool):', !!initialGrantDateStr);

        if (!hireDateStr || initialGrantDateStr) {
          console.log('  → スキップされる: !入社日=' + !hireDateStr + ', 初回付与済み=' + !!initialGrantDateStr);
          console.log('');
          continue;
        }

        var hireDate = new Date(hireDateStr);
        hireDate.setHours(0, 0, 0, 0);

        var sixMonthDate = new Date(hireDate);
        sixMonthDate.setMonth(sixMonthDate.getMonth() + 6);
        sixMonthDate.setHours(0, 0, 0, 0);

        console.log('  入社日:', Utilities.formatDate(hireDate, 'JST', 'yyyy/MM/dd'));
        console.log('  6ヶ月後:', Utilities.formatDate(sixMonthDate, 'JST', 'yyyy/MM/dd'));
        console.log('  6ヶ月経過?:', today >= sixMonthDate);
        console.log('  週所定労働日数:', weeklyWorkDaysStr);
        console.log('');
      }
    }

    // 実際の関数を実行
    console.log('=== getSixMonthGrantTargets() 実行 ===');
    var targets = getSixMonthGrantTargets();
    console.log('対象者数:', targets.length);
    targets.forEach(function(target) {
      console.log('  - ' + target.userId + ': ' + target.name);
    });

  } catch (error) {
    console.error('エラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

/**
 * 年次付与対象者のデバッグ
 */
function debugAnnualTargets() {
  console.log('=== 年次付与対象者デバッグ ===');

  try {
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');
    var data = masterSheet.getDataRange().getValues();
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    console.log('今日:', Utilities.formatDate(today, 'JST', 'yyyy/MM/dd'));
    console.log('');

    // TEST03だけ詳細チェック
    for (var i = 1; i < data.length; i++) {
      var userId = String(data[i][0]);
      if (userId === 'TEST03') {
        var name = String(data[i][1]);
        var hireDateStr = data[i][4];
        var initialGrantDateStr = data[i][6];
        var latestAnnualGrantDateStr = data[i][7];

        console.log(userId + ' (' + name + '):');
        console.log('  入社日(raw):', hireDateStr);
        console.log('  初回付与日(raw):', initialGrantDateStr);
        console.log('  最新年次付与日(raw):', latestAnnualGrantDateStr);

        if (!hireDateStr || !initialGrantDateStr) {
          console.log('  → スキップされる: 入社日または初回付与日がない');
          console.log('');
          continue;
        }

        var hireDate = new Date(hireDateStr);
        var initialGrantDate = new Date(initialGrantDateStr);
        var latestAnnualGrantDate = latestAnnualGrantDateStr ? new Date(latestAnnualGrantDateStr) : null;

        console.log('  入社日:', Utilities.formatDate(hireDate, 'JST', 'yyyy/MM/dd'));
        console.log('  初回付与日:', Utilities.formatDate(initialGrantDate, 'JST', 'yyyy/MM/dd'));
        if (latestAnnualGrantDate) {
          console.log('  最新年次付与日:', Utilities.formatDate(latestAnnualGrantDate, 'JST', 'yyyy/MM/dd'));
        } else {
          console.log('  最新年次付与日: なし');
        }

        // 次回付与日の計算
        var nextGrantDate = new Date(initialGrantDate);
        nextGrantDate.setFullYear(nextGrantDate.getFullYear() + 1);

        if (latestAnnualGrantDate) {
          nextGrantDate = new Date(latestAnnualGrantDate);
          nextGrantDate.setFullYear(nextGrantDate.getFullYear() + 1);
        }

        console.log('  次回付与日:', Utilities.formatDate(nextGrantDate, 'JST', 'yyyy/MM/dd'));
        console.log('  付与日到来?:', today >= nextGrantDate);
        console.log('');
      }
    }

    // 実際の関数を実行
    console.log('=== getAnnualGrantTargets() 実行 ===');
    var targets = getAnnualGrantTargets();
    console.log('対象者数:', targets.length);
    targets.forEach(function(target) {
      console.log('  - ' + target.userId + ': ' + target.name);
    });

  } catch (error) {
    console.error('エラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

/**
 * 付与履歴読み込みのデバッグ
 */
function debugGrantHistory() {
  console.log('=== 付与履歴読み込みデバッグ ===');

  try {
    var result = getRecentGrantHistory(50);
    console.log('取得件数:', result.length);
    console.log('');

    if (result.length > 0) {
      console.log('最新5件:');
      for (var i = 0; i < Math.min(5, result.length); i++) {
        var record = result[i];
        console.log((i + 1) + '. ' + record.userId + ' - ' + record.userName);
        console.log('   付与日: ' + record.grantDate);
        console.log('   付与日数: ' + record.grantDays + '日');
        console.log('   失効日: ' + record.expiryDate);
        console.log('   残日数: ' + record.remainingDays + '日');
        console.log('');
      }
    } else {
      console.log('付与履歴が取得できませんでした');
    }

  } catch (error) {
    console.error('エラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

/**
 * 失効予定確認のデバッグ
 */
function debugExpiringLeaves() {
  console.log('=== 失効予定確認デバッグ ===');

  try {
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    console.log('今日:', Utilities.formatDate(today, 'JST', 'yyyy/MM/dd'));
    console.log('');

    // 1週間以内
    console.log('--- 1週間以内の失効予定 ---');
    var leaves7days = getExpiringLeaves(7);
    console.log('件数:', leaves7days.length);
    leaves7days.forEach(function(leave) {
      var daysUntil = Math.ceil((new Date(leave.expiryDate) - today) / (1000 * 60 * 60 * 24));
      console.log('  - ' + leave.userId + ': ' + leave.userName);
      console.log('    失効日: ' + leave.expiryDate + ' (あと' + daysUntil + '日)');
      console.log('    残日数: ' + leave.remainingDays + '日');
    });
    console.log('');

    // 1ヶ月以内
    console.log('--- 1ヶ月以内の失効予定 ---');
    var leaves30days = getExpiringLeaves(30);
    console.log('件数:', leaves30days.length);
    leaves30days.forEach(function(leave) {
      var daysUntil = Math.ceil((new Date(leave.expiryDate) - today) / (1000 * 60 * 60 * 24));
      console.log('  - ' + leave.userId + ': ' + leave.userName);
      console.log('    失効日: ' + leave.expiryDate + ' (あと' + daysUntil + '日)');
      console.log('    残日数: ' + leave.remainingDays + '日');
    });
    console.log('');

    // 3ヶ月以内
    console.log('--- 3ヶ月以内の失効予定 ---');
    var leaves90days = getExpiringLeaves(90);
    console.log('件数:', leaves90days.length);
    console.log('最初の5件:');
    for (var i = 0; i < Math.min(5, leaves90days.length); i++) {
      var leave = leaves90days[i];
      var daysUntil = Math.ceil((new Date(leave.expiryDate) - today) / (1000 * 60 * 60 * 24));
      console.log('  ' + (i + 1) + '. ' + leave.userId + ': ' + leave.userName);
      console.log('     失効日: ' + leave.expiryDate + ' (あと' + daysUntil + '日)');
      console.log('     残日数: ' + leave.remainingDays + '日');
    }

  } catch (error) {
    console.error('エラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}
