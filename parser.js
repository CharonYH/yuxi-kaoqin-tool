(function attachParser(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.YuxiAttendanceParser = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function createParser() {
  'use strict';

  const WEEKDAYS = ['日', '一', '二', '三', '四', '五', '六'];
  const REQUIRED_SHEETS = ['排班信息', '考勤汇总', '考勤记录', '异常统计'];
  const FORMAT_ERROR_MESSAGE = '请上传格式正确的考勤表';

  function clean(value) {
    if (value === null || value === undefined) return '';
    return String(value).replace(/\r/g, '\n').replace(/\n+/g, ' ').trim();
  }

  function normalized(value) {
    return clean(value).replace(/\s+/g, '');
  }

  function extractTimes(value) {
    const text = clean(value);
    if (!text) return [];
    const matches = text.match(/(?:[01]?\d|2[0-3]):[0-5]\d/g);
    return matches || [text];
  }

  function findLabelValue(row, labelText) {
    const labelIndex = row.findIndex(cell => normalized(cell).includes(labelText));
    if (labelIndex < 0) return '';
    for (let index = labelIndex + 1; index < row.length; index += 1) {
      const value = clean(row[index]);
      if (value) return value;
    }
    return '';
  }

  function validateWorkbookStructure(workbook) {
    const sheetNames = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
    const normalizedNames = sheetNames.map(normalized);
    const valid = sheetNames.length === REQUIRED_SHEETS.length
      && new Set(normalizedNames).size === REQUIRED_SHEETS.length
      && REQUIRED_SHEETS.every(name => normalizedNames.includes(name));
    if (!valid) throw new Error(FORMAT_ERROR_MESSAGE);
  }

  function parseReport(workbook, fileName, XLSX) {
    validateWorkbookStructure(workbook);
    const sheetName = workbook.SheetNames.find(name => normalized(name) === '考勤记录');
    if (!sheetName) throw new Error(FORMAT_ERROR_MESSAGE);

    // 业务数据只读取“考勤记录”；其他三个 Tab 仅参与工作簿名称校验。
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', raw: false, blankrows: true });
    const dayHeaderIndex = rows.findIndex(row => {
      const values = row.map(value => Number.parseInt(clean(value), 10)).filter(value => value >= 1 && value <= 31);
      return new Set(values).size >= 20;
    });
    if (dayHeaderIndex < 0) throw new Error(FORMAT_ERROR_MESSAGE);

    const dayColumns = [];
    rows[dayHeaderIndex].forEach((value, columnIndex) => {
      const day = Number.parseInt(clean(value), 10);
      if (day >= 1 && day <= 31 && !dayColumns.some(item => item.day === day)) dayColumns.push({ day, columnIndex });
    });
    dayColumns.sort((a, b) => a.day - b.day);

    const topText = rows.slice(0, dayHeaderIndex + 1).flat().map(clean).join(' ');
    const dateRange = topText.match(/(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})\s*[~至—-]+\s*(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if (!dateRange) throw new Error(FORMAT_ERROR_MESSAGE);
    const year = Number(dateRange[1]);
    const month = Number(dateRange[2]);
    const startDay = Number(dateRange[3]);
    const endYear = Number(dateRange[4]);
    const endMonth = Number(dateRange[5]);
    const endDay = Number(dateRange[6]);
    if (month < 1 || month > 12 || endMonth < 1 || endMonth > 12) throw new Error(FORMAT_ERROR_MESSAGE);

    const daysInMonth = new Date(Date.UTC(year, month, 0)).getUTCDate();
    const completeMonth = year === endYear && month === endMonth && startDay === 1 && endDay === daysInMonth;
    const matchingColumns = dayColumns.length === daysInMonth && dayColumns.every(({ day }, index) => day === index + 1);
    if (!completeMonth || !matchingColumns) throw new Error(FORMAT_ERROR_MESSAGE);

    const employees = [];
    for (let rowIndex = dayHeaderIndex + 1; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex] || [];
      const rowText = row.map(normalized).join('|');
      if (!rowText.includes('工号') || !rowText.includes('姓名')) continue;
      const id = findLabelValue(row, '工号');
      const name = findLabelValue(row, '姓名');
      const department = findLabelValue(row, '部门') || '未标注';
      if (!id && !name) continue;
      const punchRow = rows[rowIndex + 1] || [];
      const days = dayColumns.map(({ day, columnIndex }) => ({
        day,
        rawValue: clean(punchRow[columnIndex]),
        times: extractTimes(punchRow[columnIndex])
      }));
      employees.push({ id, name: name || `工号 ${id}`, department, days });
      rowIndex += 1;
    }
    if (!employees.length || new Set(employees.map(employee => employee.id)).size !== employees.length) {
      throw new Error(FORMAT_ERROR_MESSAGE);
    }

    const totals = employees.reduce((result, employee) => {
      employee.days.forEach(entry => {
        result.punches += entry.times.length;
        if (entry.times.length) result.recorded += 1; else result.empty += 1;
      });
      return result;
    }, { punches: 0, recorded: 0, empty: 0 });

    const period = `${dateRange[1]}-${dateRange[2].padStart(2, '0')}-${dateRange[3].padStart(2, '0')} — ${dateRange[4]}-${dateRange[5].padStart(2, '0')}-${dateRange[6].padStart(2, '0')}`;
    return {
      fileName, sheetName, year, month, period,
      days: Array.from({ length: daysInMonth }, (_, index) => index + 1),
      employees, totals
    };
  }

  return { WEEKDAYS, REQUIRED_SHEETS, FORMAT_ERROR_MESSAGE, clean, normalized, extractTimes, parseReport, validateWorkbookStructure };
});
