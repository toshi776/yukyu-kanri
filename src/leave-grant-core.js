/**
 * leave-grant.js で利用される純粋関数群を切り出した共有モジュール。
 * GAS 環境ではグローバル関数として利用され、Node 環境では module.exports で参照できる。
 */
(function (root, factory) {
  var core = factory();

  // Node.js / CommonJS
  if (typeof module === 'object' && module.exports) {
    module.exports = core;
  }

  // GAS / ブラウザ環境（グローバルにエクスポート）
  if (root && typeof root === 'object') {
    Object.keys(core).forEach(function (key) {
      if (typeof root[key] === 'undefined') {
        root[key] = core[key];
      }
    });
  }
})(typeof this !== 'undefined' ? this : globalThis, function () {
  /**
   * 初回付与日数を計算（入社6ヶ月後）
   * @param {number} weeklyWorkDays
   * @return {number}
   */
  function calculateInitialLeaveDays(weeklyWorkDays) {
    var weekDays = Math.floor(weeklyWorkDays);

    if (weekDays >= 5) return 10;
    if (weekDays === 4) return 7;
    if (weekDays === 3) return 5;
    if (weekDays === 2) return 3;
    if (weekDays === 1) return 1;

    return 0;
  }

  /**
   * 年次付与日数を計算（2回目以降）
   * @param {number} workYears
   * @param {number} weeklyWorkDays
   * @return {number}
   */
  function calculateAnnualLeaveDays(workYears, weeklyWorkDays) {
    var years = Math.floor(workYears);
    var weekDays = Math.floor(weeklyWorkDays);

    if (weekDays >= 5) {
      if (years < 1) return 10;
      if (years < 2) return 11;
      if (years < 3) return 12;
      if (years < 4) return 14;
      if (years < 5) return 16;
      if (years < 6) return 18;
      return 20;
    }

    if (weekDays === 4) {
      if (years < 1) return 7;
      if (years < 2) return 8;
      if (years < 3) return 9;
      if (years < 4) return 10;
      if (years < 5) return 12;
      if (years < 6) return 13;
      return 15;
    }

    if (weekDays === 3) {
      if (years < 1) return 5;
      if (years < 2) return 6;
      if (years < 3) return 6;
      if (years < 4) return 8;
      if (years < 5) return 9;
      if (years < 6) return 10;
      return 11;
    }

    if (weekDays === 2) {
      if (years < 1) return 3;
      if (years < 2) return 4;
      if (years < 3) return 4;
      if (years < 4) return 5;
      if (years < 5) return 6;
      if (years < 6) return 6;
      return 7;
    }

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
   * 有給付与日数を計算
   * @param {number} workYears
   * @param {number} weeklyWorkDays
   * @param {boolean} isInitial
   * @return {number}
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
   * 勤続年数を計算
   * @param {Date} hireDate
   * @param {Date} baseDate
   * @return {number}
   */
  function calculateWorkYears(hireDate, baseDate) {
    if (!hireDate || !baseDate) return 0;

    var hire = new Date(hireDate);
    var base = new Date(baseDate);

    hire.setHours(0, 0, 0, 0);
    base.setHours(0, 0, 0, 0);

    var diffDays = Math.floor((base - hire) / (1000 * 60 * 60 * 24));
    var workYears = diffDays / 365.25;

    return Math.max(0, workYears);
  }

  /**
   * 年5日取得義務対象か判定
   * @param {number} grantedDays
   * @return {boolean}
   */
  function isFiveDayObligationTarget(grantedDays) {
    return grantedDays >= 10;
  }

  /**
   * 日付をYYYY/MM/DD形式にフォーマット
   * @param {Date|string|number} date
   * @return {string}
   */
  function formatDate(date) {
    if (!date) return '-';

    var d = new Date(date);
    if (isNaN(d.getTime())) return '-';

    var year = d.getFullYear();
    var month = ('0' + (d.getMonth() + 1)).slice(-2);
    var day = ('0' + d.getDate()).slice(-2);

    return year + '/' + month + '/' + day;
  }

  return {
    calculateLeaveDays: calculateLeaveDays,
    calculateInitialLeaveDays: calculateInitialLeaveDays,
    calculateAnnualLeaveDays: calculateAnnualLeaveDays,
    calculateWorkYears: calculateWorkYears,
    isFiveDayObligationTarget: isFiveDayObligationTarget,
    formatDate: formatDate
  };
});
