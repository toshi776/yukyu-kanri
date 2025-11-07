// =============================
// 統計レポート機能システム
// =============================

/**
 * 統計レポートの設定
 */
var REPORT_CONFIG = {
  REPORT_TYPES: {
    MONTHLY: '月次レポート',
    QUARTERLY: '四半期レポート', 
    ANNUAL: '年次レポート',
    CUSTOM: 'カスタムレポート'
  },
  FISCAL_YEAR_START: 4, // 4月開始
  MAX_HISTORY_MONTHS: 24 // 最大24ヶ月の履歴保持
};

/**
 * 月次統計レポートを生成
 * @param {number} year - 年
 * @param {number} month - 月（1-12）
 * @return {Object} レポートデータ
 */
function generateMonthlyReport(year, month) {
  try {
    console.log('月次レポート生成開始:', year + '年' + month + '月');
    
    var reportDate = new Date(year, month - 1, 1);
    var nextMonth = new Date(year, month, 1);
    
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');
    var applySheet = ss.getSheetByName('申請');
    var grantHistorySheet = ss.getSheetByName('付与履歴');
    
    // 基本統計データ
    var basicStats = calculateBasicStatistics(reportDate);
    
    // 申請統計データ
    var applicationStats = calculateApplicationStatistics(applySheet, reportDate, nextMonth);
    
    // 付与統計データ
    var grantStats = calculateGrantStatistics(grantHistorySheet, reportDate, nextMonth);
    
    // 利用率統計データ
    var usageStats = calculateUsageStatistics(masterSheet, applySheet, reportDate, nextMonth);
    
    // 事業所別統計データ
    var divisionStats = calculateDivisionStatistics(reportDate, nextMonth);
    
    var reportData = {
      reportType: 'MONTHLY',
      period: {
        year: year,
        month: month,
        displayName: year + '年' + month + '月'
      },
      generatedAt: new Date(),
      basicStatistics: basicStats,
      applicationStatistics: applicationStats,
      grantStatistics: grantStats,
      usageStatistics: usageStats,
      divisionStatistics: divisionStats
    };
    
    console.log('月次レポート生成完了');
    
    return {
      success: true,
      reportData: reportData,
      message: year + '年' + month + '月の月次レポートを生成しました'
    };
    
  } catch (error) {
    console.error('月次レポート生成エラー:', error);
    return {
      success: false,
      error: error.message,
      message: '月次レポートの生成に失敗しました: ' + error.message
    };
  }
}

/**
 * 基本統計データを計算
 * @param {Date} reportDate - レポート基準日
 * @return {Object} 基本統計データ
 */
function calculateBasicStatistics(reportDate) {
  try {
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');
    var data = masterSheet.getDataRange().getValues();
    
    var totalEmployees = 0;
    var totalRemainingDays = 0;
    var lowRemainingCount = 0; // 5日以下
    var zeroRemainingCount = 0; // 0日
    
    var divisionCounts = {
      'R': 0, // ライズ
      'P': 0, // パロン
      'S': 0, // シエル
      'E': 0  // EBISU
    };
    
    // ヘッダー行をスキップして処理
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var userId = String(row[0]);
      var remaining = Number(row[2] || 0);
      
      if (userId) {
        totalEmployees++;
        totalRemainingDays += remaining;
        
        if (remaining <= 5 && remaining > 0) {
          lowRemainingCount++;
        } else if (remaining === 0) {
          zeroRemainingCount++;
        }
        
        // 事業所別カウント
        var division = userId.substring(0, 1);
        if (divisionCounts.hasOwnProperty(division)) {
          divisionCounts[division]++;
        }
      }
    }
    
    var averageRemainingDays = totalEmployees > 0 ? 
      Math.round((totalRemainingDays / totalEmployees) * 10) / 10 : 0;
    
    return {
      totalEmployees: totalEmployees,
      totalRemainingDays: totalRemainingDays,
      averageRemainingDays: averageRemainingDays,
      lowRemainingCount: lowRemainingCount,
      zeroRemainingCount: zeroRemainingCount,
      divisionCounts: divisionCounts
    };
    
  } catch (error) {
    console.error('基本統計計算エラー:', error);
    return {
      totalEmployees: 0,
      totalRemainingDays: 0,
      averageRemainingDays: 0,
      lowRemainingCount: 0,
      zeroRemainingCount: 0,
      divisionCounts: { 'R': 0, 'P': 0, 'S': 0, 'E': 0 }
    };
  }
}

/**
 * 申請統計データを計算
 * @param {Sheet} applySheet - 申請シート
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Object} 申請統計データ
 */
function calculateApplicationStatistics(applySheet, startDate, endDate) {
  try {
    if (!applySheet) {
      return {
        totalApplications: 0,
        approvedApplications: 0,
        rejectedApplications: 0,
        pendingApplications: 0,
        totalAppliedDays: 0,
        averageApplicationDays: 0
      };
    }
    
    var data = applySheet.getDataRange().getValues();
    var totalApplications = 0;
    var approvedApplications = 0;
    var rejectedApplications = 0;
    var pendingApplications = 0;
    var totalAppliedDays = 0;
    
    // ヘッダー行をスキップして処理
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var applyDate = row[3] ? new Date(row[3]) : null;
      var status = String(row[5] || '');
      var appliedDays = Number(row[7] || 0);
      
      // 期間内の申請のみカウント
      if (applyDate && applyDate >= startDate && applyDate < endDate) {
        totalApplications++;
        totalAppliedDays += appliedDays;
        
        switch (status) {
          case 'Approved':
            approvedApplications++;
            break;
          case 'Rejected':
            rejectedApplications++;
            break;
          default:
            pendingApplications++;
            break;
        }
      }
    }
    
    var averageApplicationDays = totalApplications > 0 ? 
      Math.round((totalAppliedDays / totalApplications) * 10) / 10 : 0;
    
    return {
      totalApplications: totalApplications,
      approvedApplications: approvedApplications,
      rejectedApplications: rejectedApplications,
      pendingApplications: pendingApplications,
      totalAppliedDays: totalAppliedDays,
      averageApplicationDays: averageApplicationDays,
      approvalRate: totalApplications > 0 ? 
        Math.round((approvedApplications / totalApplications) * 100) : 0
    };
    
  } catch (error) {
    console.error('申請統計計算エラー:', error);
    return {
      totalApplications: 0,
      approvedApplications: 0,
      rejectedApplications: 0,
      pendingApplications: 0,
      totalAppliedDays: 0,
      averageApplicationDays: 0,
      approvalRate: 0
    };
  }
}

/**
 * 付与統計データを計算
 * @param {Sheet} grantHistorySheet - 付与履歴シート
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Object} 付与統計データ
 */
function calculateGrantStatistics(grantHistorySheet, startDate, endDate) {
  try {
    if (!grantHistorySheet) {
      return {
        totalGrants: 0,
        totalGrantedDays: 0,
        initialGrants: 0,
        annualGrants: 0,
        averageGrantDays: 0
      };
    }
    
    var data = grantHistorySheet.getDataRange().getValues();
    var totalGrants = 0;
    var totalGrantedDays = 0;
    var initialGrants = 0;
    var annualGrants = 0;
    
    // ヘッダー行をスキップして処理
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var grantDate = row[1] ? new Date(row[1]) : null;
      var grantDays = Number(row[2] || 0);
      var grantType = String(row[5] || '');
      
      // 期間内の付与のみカウント
      if (grantDate && grantDate >= startDate && grantDate < endDate) {
        totalGrants++;
        totalGrantedDays += grantDays;
        
        if (grantType === '初回') {
          initialGrants++;
        } else if (grantType === '年次') {
          annualGrants++;
        }
      }
    }
    
    var averageGrantDays = totalGrants > 0 ? 
      Math.round((totalGrantedDays / totalGrants) * 10) / 10 : 0;
    
    return {
      totalGrants: totalGrants,
      totalGrantedDays: totalGrantedDays,
      initialGrants: initialGrants,
      annualGrants: annualGrants,
      averageGrantDays: averageGrantDays
    };
    
  } catch (error) {
    console.error('付与統計計算エラー:', error);
    return {
      totalGrants: 0,
      totalGrantedDays: 0,
      initialGrants: 0,
      annualGrants: 0,
      averageGrantDays: 0
    };
  }
}

/**
 * 利用率統計データを計算
 * @param {Sheet} masterSheet - マスターシート
 * @param {Sheet} applySheet - 申請シート
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Object} 利用率統計データ
 */
function calculateUsageStatistics(masterSheet, applySheet, startDate, endDate) {
  try {
    // 年5日取得義務対象者の抽出と分析
    var fiveDayObligationStats = calculateFiveDayObligationStats(startDate, endDate);
    
    // 全体の利用率計算
    var overallUsageRate = calculateOverallUsageRate(masterSheet, applySheet, startDate, endDate);
    
    return {
      overallUsageRate: overallUsageRate,
      fiveDayObligation: fiveDayObligationStats
    };
    
  } catch (error) {
    console.error('利用率統計計算エラー:', error);
    return {
      overallUsageRate: 0,
      fiveDayObligation: {
        targetEmployees: 0,
        compliantEmployees: 0,
        complianceRate: 0
      }
    };
  }
}

/**
 * 年5日取得義務の統計を計算
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Object} 年5日義務統計データ
 */
function calculateFiveDayObligationStats(startDate, endDate) {
  try {
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');
    var applySheet = ss.getSheetByName('申請');
    var grantHistorySheet = ss.getSheetByName('付与履歴');
    
    if (!masterSheet || !applySheet || !grantHistorySheet) {
      return {
        targetEmployees: 0,
        compliantEmployees: 0,
        complianceRate: 0,
        details: []
      };
    }
    
    var masterData = masterSheet.getDataRange().getValues();
    var applyData = applySheet.getDataRange().getValues();
    
    var targetEmployees = 0;
    var compliantEmployees = 0;
    var details = [];
    
    // 各従業員の年5日取得義務をチェック
    for (var i = 1; i < masterData.length; i++) {
      var userId = String(masterData[i][0]);
      
      if (!userId) continue;
      
      // この従業員が年5日取得義務の対象かチェック
      var isTarget = checkFiveDayObligationTarget(userId, startDate);
      
      if (isTarget) {
        targetEmployees++;
        
        // 期間内の取得日数を計算
        var takenDays = calculateTakenDaysInPeriod(applyData, userId, startDate, endDate);
        var isCompliant = takenDays >= 5;
        
        if (isCompliant) {
          compliantEmployees++;
        }
        
        details.push({
          userId: userId,
          userName: String(masterData[i][1] || ''),
          takenDays: takenDays,
          isCompliant: isCompliant
        });
      }
    }
    
    var complianceRate = targetEmployees > 0 ? 
      Math.round((compliantEmployees / targetEmployees) * 100) : 0;
    
    return {
      targetEmployees: targetEmployees,
      compliantEmployees: compliantEmployees,
      complianceRate: complianceRate,
      details: details
    };
    
  } catch (error) {
    console.error('年5日義務統計計算エラー:', error);
    return {
      targetEmployees: 0,
      compliantEmployees: 0,
      complianceRate: 0,
      details: []
    };
  }
}

/**
 * 年5日取得義務の対象者かチェック
 * @param {string} userId - 利用者番号
 * @param {Date} checkDate - チェック日
 * @return {boolean} 対象かどうか
 */
function checkFiveDayObligationTarget(userId, checkDate) {
  try {
    var ss = getSpreadsheet();
    var grantHistorySheet = ss.getSheetByName('付与履歴');
    
    if (!grantHistorySheet) return false;
    
    var data = grantHistorySheet.getDataRange().getValues();
    var totalGranted = 0;
    
    // 直近の付与日数を合計（失効前のもの）
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      if (String(row[0]) === userId) {
        var grantDate = new Date(row[1]);
        var expiryDate = new Date(row[3]);
        var grantDays = Number(row[2] || 0);
        
        // チェック日時点で有効な付与分を合計
        if (grantDate <= checkDate && expiryDate > checkDate) {
          totalGranted += grantDays;
        }
      }
    }
    
    // 10日以上付与されている場合は年5日取得義務対象
    return totalGranted >= 10;
    
  } catch (error) {
    console.error('年5日義務対象者チェックエラー:', error);
    return false;
  }
}

/**
 * 期間内の取得日数を計算
 * @param {Array} applyData - 申請データ
 * @param {string} userId - 利用者番号
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {number} 取得日数
 */
function calculateTakenDaysInPeriod(applyData, userId, startDate, endDate) {
  try {
    var takenDays = 0;
    
    for (var i = 1; i < applyData.length; i++) {
      var row = applyData[i];
      
      if (String(row[0]) === userId && row[5] === 'Approved') {
        var applyDate = new Date(row[3]);
        var appliedDays = Number(row[7] || 0);
        
        if (applyDate >= startDate && applyDate < endDate) {
          takenDays += appliedDays;
        }
      }
    }
    
    return takenDays;
    
  } catch (error) {
    console.error('取得日数計算エラー:', error);
    return 0;
  }
}

/**
 * 全体利用率を計算
 * @param {Sheet} masterSheet - マスターシート
 * @param {Sheet} applySheet - 申請シート
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {number} 利用率（%）
 */
function calculateOverallUsageRate(masterSheet, applySheet, startDate, endDate) {
  try {
    if (!masterSheet || !applySheet) return 0;
    
    var masterData = masterSheet.getDataRange().getValues();
    var applyData = applySheet.getDataRange().getValues();
    
    var totalAvailable = 0;
    var totalUsed = 0;
    
    for (var i = 1; i < masterData.length; i++) {
      var userId = String(masterData[i][0]);
      var currentRemaining = Number(masterData[i][2] || 0);
      
      if (userId) {
        // 期間内の使用日数を計算
        var usedInPeriod = calculateTakenDaysInPeriod(applyData, userId, startDate, endDate);
        
        // 利用可能日数 = 現在の残日数 + 期間内使用日数
        var availableInPeriod = currentRemaining + usedInPeriod;
        
        totalAvailable += availableInPeriod;
        totalUsed += usedInPeriod;
      }
    }
    
    return totalAvailable > 0 ? Math.round((totalUsed / totalAvailable) * 100) : 0;
    
  } catch (error) {
    console.error('全体利用率計算エラー:', error);
    return 0;
  }
}

/**
 * 事業所別統計データを計算
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Object} 事業所別統計データ
 */
function calculateDivisionStatistics(startDate, endDate) {
  try {
    var divisions = {
      'R': { name: 'ライズ', employees: 0, totalRemaining: 0, applications: 0, averageRemaining: 0 },
      'P': { name: 'パロン', employees: 0, totalRemaining: 0, applications: 0, averageRemaining: 0 },
      'S': { name: 'シエル', employees: 0, totalRemaining: 0, applications: 0, averageRemaining: 0 },
      'E': { name: 'EBISU', employees: 0, totalRemaining: 0, applications: 0, averageRemaining: 0 }
    };
    
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName('マスター');
    var applySheet = ss.getSheetByName('申請');
    
    if (masterSheet) {
      var masterData = masterSheet.getDataRange().getValues();
      
      for (var i = 1; i < masterData.length; i++) {
        var userId = String(masterData[i][0]);
        var remaining = Number(masterData[i][2] || 0);
        
        if (userId) {
          var divisionCode = userId.substring(0, 1);
          
          if (divisions[divisionCode]) {
            divisions[divisionCode].employees++;
            divisions[divisionCode].totalRemaining += remaining;
          }
        }
      }
    }
    
    if (applySheet) {
      var applyData = applySheet.getDataRange().getValues();
      
      for (var i = 1; i < applyData.length; i++) {
        var userId = String(applyData[i][0]);
        var applyDate = applyData[i][3] ? new Date(applyData[i][3]) : null;
        
        if (userId && applyDate && applyDate >= startDate && applyDate < endDate) {
          var divisionCode = userId.substring(0, 1);
          
          if (divisions[divisionCode]) {
            divisions[divisionCode].applications++;
          }
        }
      }
    }
    
    // 平均残日数を計算
    Object.keys(divisions).forEach(function(code) {
      var div = divisions[code];
      if (div.employees > 0) {
        div.averageRemaining = Math.round((div.totalRemaining / div.employees) * 10) / 10;
      }
    });
    
    return divisions;
    
  } catch (error) {
    console.error('事業所別統計計算エラー:', error);
    return {
      'R': { name: 'ライズ', employees: 0, totalRemaining: 0, applications: 0, averageRemaining: 0 },
      'P': { name: 'パロン', employees: 0, totalRemaining: 0, applications: 0, averageRemaining: 0 },
      'S': { name: 'シエル', employees: 0, totalRemaining: 0, applications: 0, averageRemaining: 0 },
      'E': { name: 'EBISU', employees: 0, totalRemaining: 0, applications: 0, averageRemaining: 0 }
    };
  }
}

/**
 * 年次レポートを生成
 * @param {number} fiscalYear - 年度
 * @return {Object} 年次レポートデータ
 */
function generateAnnualReport(fiscalYear) {
  try {
    console.log('年次レポート生成開始:', fiscalYear + '年度');
    
    var startDate = new Date(fiscalYear, REPORT_CONFIG.FISCAL_YEAR_START - 1, 1); // 4月1日
    var endDate = new Date(fiscalYear + 1, REPORT_CONFIG.FISCAL_YEAR_START - 1, 1); // 次年度4月1日
    
    var monthlyReports = [];
    
    // 12ヶ月分の月次レポートを生成
    for (var month = 0; month < 12; month++) {
      var reportMonth = new Date(startDate);
      reportMonth.setMonth(startDate.getMonth() + month);
      
      var monthlyReport = generateMonthlyReport(reportMonth.getFullYear(), reportMonth.getMonth() + 1);
      if (monthlyReport.success) {
        monthlyReports.push(monthlyReport.reportData);
      }
    }
    
    // 年間集計データを計算
    var annualSummary = calculateAnnualSummary(monthlyReports);
    
    var reportData = {
      reportType: 'ANNUAL',
      fiscalYear: fiscalYear,
      period: {
        startDate: startDate,
        endDate: endDate,
        displayName: fiscalYear + '年度'
      },
      generatedAt: new Date(),
      monthlyReports: monthlyReports,
      annualSummary: annualSummary
    };
    
    console.log('年次レポート生成完了');
    
    return {
      success: true,
      reportData: reportData,
      message: fiscalYear + '年度の年次レポートを生成しました'
    };
    
  } catch (error) {
    console.error('年次レポート生成エラー:', error);
    return {
      success: false,
      error: error.message,
      message: '年次レポートの生成に失敗しました: ' + error.message
    };
  }
}

/**
 * 年間集計データを計算
 * @param {Array} monthlyReports - 月次レポート配列
 * @return {Object} 年間集計データ
 */
function calculateAnnualSummary(monthlyReports) {
  try {
    var totalApplications = 0;
    var totalApprovedApplications = 0;
    var totalAppliedDays = 0;
    var totalGrantedDays = 0;
    var totalGrants = 0;
    
    monthlyReports.forEach(function(report) {
      if (report.applicationStatistics) {
        totalApplications += report.applicationStatistics.totalApplications;
        totalApprovedApplications += report.applicationStatistics.approvedApplications;
        totalAppliedDays += report.applicationStatistics.totalAppliedDays;
      }
      
      if (report.grantStatistics) {
        totalGrantedDays += report.grantStatistics.totalGrantedDays;
        totalGrants += report.grantStatistics.totalGrants;
      }
    });
    
    return {
      totalApplications: totalApplications,
      totalApprovedApplications: totalApprovedApplications,
      totalAppliedDays: totalAppliedDays,
      totalGrantedDays: totalGrantedDays,
      totalGrants: totalGrants,
      averageApplicationsPerMonth: monthlyReports.length > 0 ? 
        Math.round(totalApplications / monthlyReports.length * 10) / 10 : 0,
      overallApprovalRate: totalApplications > 0 ? 
        Math.round((totalApprovedApplications / totalApplications) * 100) : 0
    };
    
  } catch (error) {
    console.error('年間集計計算エラー:', error);
    return {
      totalApplications: 0,
      totalApprovedApplications: 0,
      totalAppliedDays: 0,
      totalGrantedDays: 0,
      totalGrants: 0,
      averageApplicationsPerMonth: 0,
      overallApprovalRate: 0
    };
  }
}

/**
 * レポートデータをHTMLに変換
 * @param {Object} reportData - レポートデータ
 * @return {string} HTMLレポート
 */
function generateReportHTML(reportData) {
  try {
    var html = '';
    
    html += '<div class="report-container" style="font-family: Arial, sans-serif; max-width: 1000px; margin: 0 auto; padding: 20px;">';
    html += '<h1 style="color: #333; text-align: center;">有給管理システム レポート</h1>';
    html += '<h2 style="color: #4CAF50; text-align: center;">' + reportData.period.displayName + '</h2>';
    
    // 基本統計
    if (reportData.basicStatistics) {
      html += generateBasicStatsHTML(reportData.basicStatistics);
    }
    
    // 申請統計
    if (reportData.applicationStatistics) {
      html += generateApplicationStatsHTML(reportData.applicationStatistics);
    }
    
    // 付与統計
    if (reportData.grantStatistics) {
      html += generateGrantStatsHTML(reportData.grantStatistics);
    }
    
    // 利用率統計
    if (reportData.usageStatistics) {
      html += generateUsageStatsHTML(reportData.usageStatistics);
    }
    
    // 事業所別統計
    if (reportData.divisionStatistics) {
      html += generateDivisionStatsHTML(reportData.divisionStatistics);
    }
    
    html += '<p style="text-align: right; color: #666; margin-top: 30px;">生成日時: ' + 
            Utilities.formatDate(reportData.generatedAt, 'JST', 'yyyy/MM/dd HH:mm') + '</p>';
    html += '</div>';
    
    return html;
    
  } catch (error) {
    console.error('HTMLレポート生成エラー:', error);
    return '<p>レポートの生成でエラーが発生しました: ' + error.message + '</p>';
  }
}

/**
 * 基本統計のHTML生成
 */
function generateBasicStatsHTML(stats) {
  var html = '<div class="stats-section" style="margin: 20px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px;">';
  html += '<h3 style="color: #333; border-bottom: 2px solid #4CAF50; padding-bottom: 5px;">基本統計</h3>';
  html += '<div style="display: flex; flex-wrap: wrap; gap: 20px;">';
  
  html += '<div style="flex: 1; min-width: 200px;">';
  html += '<p><strong>総従業員数:</strong> ' + stats.totalEmployees + '名</p>';
  html += '<p><strong>総残日数:</strong> ' + stats.totalRemainingDays + '日</p>';
  html += '<p><strong>平均残日数:</strong> ' + stats.averageRemainingDays + '日</p>';
  html += '</div>';
  
  html += '<div style="flex: 1; min-width: 200px;">';
  html += '<p><strong>残日数少ない従業員:</strong> ' + stats.lowRemainingCount + '名</p>';
  html += '<p><strong>残日数0の従業員:</strong> ' + stats.zeroRemainingCount + '名</p>';
  html += '</div>';
  
  html += '</div></div>';
  
  return html;
}

/**
 * 申請統計のHTML生成
 */
function generateApplicationStatsHTML(stats) {
  var html = '<div class="stats-section" style="margin: 20px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px;">';
  html += '<h3 style="color: #333; border-bottom: 2px solid #4CAF50; padding-bottom: 5px;">申請統計</h3>';
  html += '<div style="display: flex; flex-wrap: wrap; gap: 20px;">';
  
  html += '<div style="flex: 1; min-width: 200px;">';
  html += '<p><strong>総申請数:</strong> ' + stats.totalApplications + '件</p>';
  html += '<p><strong>承認済み:</strong> ' + stats.approvedApplications + '件</p>';
  html += '<p><strong>却下済み:</strong> ' + stats.rejectedApplications + '件</p>';
  html += '</div>';
  
  html += '<div style="flex: 1; min-width: 200px;">';
  html += '<p><strong>承認率:</strong> ' + stats.approvalRate + '%</p>';
  html += '<p><strong>総申請日数:</strong> ' + stats.totalAppliedDays + '日</p>';
  html += '<p><strong>平均申請日数:</strong> ' + stats.averageApplicationDays + '日</p>';
  html += '</div>';
  
  html += '</div></div>';
  
  return html;
}

/**
 * 付与統計のHTML生成
 */
function generateGrantStatsHTML(stats) {
  var html = '<div class="stats-section" style="margin: 20px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px;">';
  html += '<h3 style="color: #333; border-bottom: 2px solid #4CAF50; padding-bottom: 5px;">付与統計</h3>';
  html += '<div style="display: flex; flex-wrap: wrap; gap: 20px;">';
  
  html += '<div style="flex: 1; min-width: 200px;">';
  html += '<p><strong>総付与数:</strong> ' + stats.totalGrants + '件</p>';
  html += '<p><strong>総付与日数:</strong> ' + stats.totalGrantedDays + '日</p>';
  html += '<p><strong>平均付与日数:</strong> ' + stats.averageGrantDays + '日</p>';
  html += '</div>';
  
  html += '<div style="flex: 1; min-width: 200px;">';
  html += '<p><strong>初回付与:</strong> ' + stats.initialGrants + '件</p>';
  html += '<p><strong>年次付与:</strong> ' + stats.annualGrants + '件</p>';
  html += '</div>';
  
  html += '</div></div>';
  
  return html;
}

/**
 * 利用率統計のHTML生成
 */
function generateUsageStatsHTML(stats) {
  var html = '<div class="stats-section" style="margin: 20px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px;">';
  html += '<h3 style="color: #333; border-bottom: 2px solid #4CAF50; padding-bottom: 5px;">利用率統計</h3>';
  html += '<div style="display: flex; flex-wrap: wrap; gap: 20px;">';
  
  html += '<div style="flex: 1; min-width: 200px;">';
  html += '<p><strong>全体利用率:</strong> ' + stats.overallUsageRate + '%</p>';
  html += '</div>';
  
  if (stats.fiveDayObligation) {
    html += '<div style="flex: 1; min-width: 200px;">';
    html += '<h4>年5日取得義務</h4>';
    html += '<p><strong>対象者:</strong> ' + stats.fiveDayObligation.targetEmployees + '名</p>';
    html += '<p><strong>達成者:</strong> ' + stats.fiveDayObligation.compliantEmployees + '名</p>';
    html += '<p><strong>達成率:</strong> ' + stats.fiveDayObligation.complianceRate + '%</p>';
    html += '</div>';
  }
  
  html += '</div></div>';
  
  return html;
}

/**
 * 事業所別統計のHTML生成
 */
function generateDivisionStatsHTML(stats) {
  var html = '<div class="stats-section" style="margin: 20px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px;">';
  html += '<h3 style="color: #333; border-bottom: 2px solid #4CAF50; padding-bottom: 5px;">事業所別統計</h3>';
  html += '<div style="display: flex; flex-wrap: wrap; gap: 20px;">';
  
  Object.keys(stats).forEach(function(code) {
    var div = stats[code];
    
    html += '<div style="flex: 1; min-width: 200px; border: 1px solid #eee; padding: 10px; border-radius: 3px;">';
    html += '<h4 style="margin: 0 0 10px 0; color: #4CAF50;">' + div.name + '</h4>';
    html += '<p><strong>従業員数:</strong> ' + div.employees + '名</p>';
    html += '<p><strong>平均残日数:</strong> ' + div.averageRemaining + '日</p>';
    html += '<p><strong>申請数:</strong> ' + div.applications + '件</p>';
    html += '</div>';
  });
  
  html += '</div></div>';
  
  return html;
}