(() => {
  'use strict';

  const $ = (selector, root = document) => root.querySelector(selector);
  const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];
  const state = { workbook: null, report: null, selectedEmployeeIndex: 0, selectedDay: 1, filter: 'all', query: '' };
  const WEEKDAYS = ['日', '一', '二', '三', '四', '五', '六'];
  const REQUIRED_SHEETS = ['排班信息', '考勤汇总', '考勤记录', '异常统计'];
  const FORMAT_ERROR_MESSAGE = '请上传格式正确的考勤表';

  const els = {
    uploadView: $('#uploadView'), dashboard: $('#dashboard'), fileInput: $('#fileInput'), dropZone: $('#dropZone'),
    backButton: $('#backButton'), exportButton: $('#exportButton'), fileName: $('#fileName'), periodText: $('#periodText'),
    employeeCount: $('#employeeCount'), punchCount: $('#punchCount'), recordedDays: $('#recordedDays'), emptyDays: $('#emptyDays'),
    coverageChart: $('#coverageChart'), attendanceMatrix: $('#attendanceMatrix'), employeePanel: $('#employeePanel'),
    searchInput: $('#searchInput'), toast: $('#toast')
  };

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
    for (let i = labelIndex + 1; i < row.length; i += 1) {
      const value = clean(row[i]);
      if (value) return value;
    }
    return '';
  }

  function findAttendanceSheet(workbook) {
    return workbook.SheetNames.find(name => normalized(name) === '考勤记录');
  }

  function validateWorkbookStructure(workbook) {
    const sheetNames = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
    const normalizedNames = sheetNames.map(normalized);
    const hasExactSheetSet = sheetNames.length === REQUIRED_SHEETS.length
      && new Set(normalizedNames).size === REQUIRED_SHEETS.length
      && REQUIRED_SHEETS.every(requiredName => normalizedNames.includes(requiredName));

    if (!hasExactSheetSet) throw new Error(FORMAT_ERROR_MESSAGE);
  }

  function parseReport(workbook, fileName) {
    validateWorkbookStructure(workbook);
    const sheetName = findAttendanceSheet(workbook);
    if (!sheetName) throw new Error(FORMAT_ERROR_MESSAGE);

    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', raw: false, blankrows: true });
    const dayHeaderIndex = rows.findIndex(row => {
      const dayValues = row.map(value => Number.parseInt(clean(value), 10)).filter(value => value >= 1 && value <= 31);
      return new Set(dayValues).size >= 20;
    });
    if (dayHeaderIndex < 0) throw new Error(FORMAT_ERROR_MESSAGE);

    const dayHeader = rows[dayHeaderIndex];
    const dayColumns = [];
    dayHeader.forEach((value, columnIndex) => {
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
    const isCompleteMonth = year === endYear && month === endMonth && startDay === 1 && endDay === daysInMonth;
    const hasMatchingDayColumns = dayColumns.length === daysInMonth
      && dayColumns.every(({ day }, index) => day === index + 1);
    if (!isCompleteMonth || !hasMatchingDayColumns) throw new Error(FORMAT_ERROR_MESSAGE);
    const period = `${dateRange[1]}-${dateRange[2].padStart(2, '0')}-${dateRange[3].padStart(2, '0')} — ${dateRange[4]}-${dateRange[5].padStart(2, '0')}-${dateRange[6].padStart(2, '0')}`;

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
      const days = dayColumns.map(({ day, columnIndex }) => ({ day, times: extractTimes(punchRow[columnIndex]) }));
      employees.push({ id, name: name || `工号 ${id}`, department, days });
      rowIndex += 1;
    }
    if (!employees.length) throw new Error(FORMAT_ERROR_MESSAGE);

    const fullDays = Array.from({ length: daysInMonth }, (_, index) => index + 1);
    const totals = employees.reduce((acc, employee) => {
      employee.days.forEach(entry => {
        acc.punches += entry.times.length;
        if (entry.times.length) acc.recorded += 1; else acc.empty += 1;
      });
      return acc;
    }, { punches: 0, recorded: 0, empty: 0 });

    return { fileName, sheetName, year, month, period, days: fullDays, employees, totals };
  }

  function getWeekday(day) {
    return new Date(Date.UTC(state.report.year, state.report.month - 1, day)).getUTCDay();
  }

  function isWeekend(day) {
    const weekday = getWeekday(day);
    return weekday === 0 || weekday === 6;
  }

  function statusFor(times) {
    if (!times.length) return 'empty';
    if (times.length === 1) return 'single';
    if (times.length === 2) return 'pair';
    return 'multiple';
  }

  function formatNumber(value) {
    return new Intl.NumberFormat('zh-CN').format(value);
  }

  function escapeHtml(value) {
    return clean(value).replace(/[&<>'"]/g, character => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' })[character]);
  }

  function showToast(message, isError = false) {
    els.toast.textContent = message;
    els.toast.classList.toggle('is-error', isError);
    els.toast.classList.add('is-visible');
    window.clearTimeout(showToast.timer);
    showToast.timer = window.setTimeout(() => els.toast.classList.remove('is-visible'), 3200);
  }

  async function loadFile(file) {
    if (!file) return;
    if (!/\.xlsx?$/i.test(file.name)) { showToast(FORMAT_ERROR_MESSAGE, true); return; }
    els.dropZone.classList.add('is-loading');
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
      state.workbook = workbook;
      state.report = parseReport(workbook, file.name);
      state.selectedEmployeeIndex = 0;
      state.selectedDay = 1;
      state.filter = 'all';
      state.query = '';
      renderDashboard();
      els.uploadView.hidden = true;
      els.dashboard.hidden = false;
      window.scrollTo({ top: 0, behavior: 'smooth' });
      showToast(`已读取 ${state.report.employees.length} 名员工的考勤记录`);
    } catch (error) {
      console.error(error);
      showToast(FORMAT_ERROR_MESSAGE, true);
    } finally {
      els.dropZone.classList.remove('is-loading');
      els.fileInput.value = '';
    }
  }

  function renderDashboard() {
    const report = state.report;
    els.fileName.textContent = report.fileName;
    els.periodText.textContent = `${report.period} · ${report.sheetName}`;
    els.employeeCount.textContent = formatNumber(report.employees.length);
    els.punchCount.textContent = formatNumber(report.totals.punches);
    els.recordedDays.textContent = formatNumber(report.totals.recorded);
    els.emptyDays.textContent = formatNumber(report.totals.empty);
    els.searchInput.value = '';
    $$('.filter-chip').forEach(button => button.classList.toggle('is-active', button.dataset.filter === 'all'));
    renderCoverage();
    renderMatrix();
    renderEmployee();
  }

  function renderCoverage() {
    const report = state.report;
    els.coverageChart.style.gridTemplateColumns = `repeat(${report.days.length}, minmax(14px, 1fr))`;
    els.coverageChart.style.minWidth = `${Math.max(640, report.days.length * 23)}px`;
    const counts = report.days.map(day => report.employees.filter(employee => employee.days.find(item => item.day === day)?.times.length).length);
    const max = Math.max(...counts, 1);
    els.coverageChart.innerHTML = report.days.map((day, index) => {
      const count = counts[index];
      const weekend = isWeekend(day);
      return `<div class="coverage-bar-wrap ${weekend ? 'is-weekend' : ''}">
        <span class="bar-tooltip">${day} 日 · ${count}/${report.employees.length} 人</span>
        <span class="coverage-bar" style="height:${Math.max(3, (count / max) * 88)}px"></span>
        <small>${day}</small>
      </div>`;
    }).join('');
  }

  function filteredEmployees() {
    const query = state.query.toLowerCase();
    return state.report.employees.map((employee, originalIndex) => ({ employee, originalIndex })).filter(({ employee }) => {
      const matchesQuery = !query || `${employee.name} ${employee.id} ${employee.department}`.toLowerCase().includes(query);
      const matchesStatus = state.filter === 'all' || employee.days.some(entry => statusFor(entry.times) === state.filter);
      return matchesQuery && matchesStatus;
    });
  }

  function renderMatrix() {
    const report = state.report;
    const people = filteredEmployees();
    if (people.length && !people.some(({ originalIndex }) => originalIndex === state.selectedEmployeeIndex)) {
      state.selectedEmployeeIndex = people[0].originalIndex;
    }
    const headers = `<div class="matrix-cell matrix-corner">员工 / 日期</div>${report.days.map(day => {
      const weekday = getWeekday(day);
      return `<div class="matrix-cell matrix-day ${isWeekend(day) ? 'is-weekend' : ''}"><span>${String(day).padStart(2, '0')}</span><small>周${WEEKDAYS[weekday]}</small></div>`;
    }).join('')}`;

    const rows = people.length ? people.map(({ employee, originalIndex }) => {
      const name = `<button class="matrix-cell matrix-name ${originalIndex === state.selectedEmployeeIndex ? 'is-selected' : ''}" type="button" data-employee="${originalIndex}"><span>${escapeHtml(employee.name)}</span><small>${escapeHtml(employee.id)}</small></button>`;
      const cells = report.days.map(day => {
        const entry = employee.days.find(item => item.day === day) || { times: [] };
        const status = statusFor(entry.times);
        const label = `${employee.name}，${report.month}月${day}日，${entry.times.length ? entry.times.join('、') : '无记录'}`;
        return `<div class="matrix-cell"><button type="button" class="day-button status-${status} ${isWeekend(day) ? 'is-weekend' : ''} ${originalIndex === state.selectedEmployeeIndex && day === state.selectedDay ? 'is-selected' : ''}" data-employee="${originalIndex}" data-day="${day}" data-count="${entry.times.length}" aria-label="${escapeHtml(label)}" title="${escapeHtml(label)}"></button></div>`;
      }).join('');
      return name + cells;
    }).join('') : '<div class="matrix-no-results">没有符合当前条件的员工</div>';

    els.attendanceMatrix.style.gridTemplateColumns = `154px repeat(${report.days.length}, 40px)`;
    els.attendanceMatrix.style.minWidth = `${154 + report.days.length * 40 + 56}px`;
    els.attendanceMatrix.innerHTML = headers + rows;
    $$('[data-employee]', els.attendanceMatrix).forEach(button => button.addEventListener('click', () => {
      state.selectedEmployeeIndex = Number(button.dataset.employee);
      if (button.dataset.day) state.selectedDay = Number(button.dataset.day);
      renderMatrix();
      renderEmployee();
    }));
  }

  function employeeStats(employee) {
    return employee.days.reduce((acc, entry) => {
      const status = statusFor(entry.times);
      acc.punches += entry.times.length;
      if (status === 'empty') acc.empty += 1;
      if (status === 'single') acc.single += 1;
      return acc;
    }, { punches: 0, empty: 0, single: 0 });
  }

  function renderEmployee() {
    const report = state.report;
    const employee = report.employees[state.selectedEmployeeIndex];
    if (!employee) return;
    const stats = employeeStats(employee);
    const firstWeekday = getWeekday(1);
    const blankDays = Array.from({ length: firstWeekday }, () => '<div class="calendar-day calendar-day--blank"></div>').join('');
    const calendarDays = report.days.map(day => {
      const entry = employee.days.find(item => item.day === day) || { times: [] };
      return `<button type="button" class="calendar-day ${isWeekend(day) ? 'is-weekend' : ''} ${entry.times.length ? '' : 'is-empty'} ${day === state.selectedDay ? 'is-active' : ''}" data-calendar-day="${day}" aria-label="${report.month}月${day}日，${entry.times.length ? entry.times.join('、') : '无记录'}">
        <b>${day}</b><span>${entry.times.length ? escapeHtml(entry.times.join(' · ')) : '—'}</span>
      </button>`;
    }).join('');
    const selected = employee.days.find(item => item.day === state.selectedDay) || { times: [] };
    els.employeePanel.innerHTML = `
      <div class="employee-head">
        <div class="employee-identity"><span class="employee-avatar">${escapeHtml(employee.name.slice(-1))}</span><div><h2>${escapeHtml(employee.name)}</h2><p>NO. ${escapeHtml(employee.id)} · ${escapeHtml(employee.department)}</p></div></div>
        <div class="employee-nav">
          <button class="icon-button" type="button" data-nav="prev" aria-label="上一名员工"><svg viewBox="0 0 24 24"><path d="m15 18-6-6 6-6"/></svg></button>
          <button class="icon-button" type="button" data-nav="next" aria-label="下一名员工"><svg viewBox="0 0 24 24"><path d="m9 18 6-6-6-6"/></svg></button>
        </div>
      </div>
      <div class="employee-stats">
        <div class="mini-stat"><strong>${stats.punches}</strong><span>打卡次数</span></div>
        <div class="mini-stat"><strong>${stats.single}</strong><span>单次记录</span></div>
        <div class="mini-stat"><strong>${stats.empty}</strong><span>无记录日</span></div>
      </div>
      <div class="calendar-labels">${WEEKDAYS.map(day => `<span>${day}</span>`).join('')}</div>
      <div class="calendar-grid">${blankDays}${calendarDays}</div>
      <div class="selected-day-detail"><p>${report.year} 年 ${report.month} 月 ${state.selectedDay} 日 · 周${WEEKDAYS[getWeekday(state.selectedDay)]}</p><strong class="${selected.times.length ? '' : 'no-punch'}">${selected.times.length ? selected.times.map(escapeHtml).join(' / ') : '无打卡记录'}</strong></div>`;

    $$('[data-calendar-day]', els.employeePanel).forEach(button => button.addEventListener('click', () => {
      state.selectedDay = Number(button.dataset.calendarDay); renderMatrix(); renderEmployee();
    }));
    $$('[data-nav]', els.employeePanel).forEach(button => button.addEventListener('click', () => {
      const direction = button.dataset.nav === 'next' ? 1 : -1;
      state.selectedEmployeeIndex = (state.selectedEmployeeIndex + direction + report.employees.length) % report.employees.length;
      renderMatrix(); renderEmployee();
    }));
  }

  function exportCsv() {
    const report = state.report;
    const rows = [['工号', '姓名', '部门', '日期', '星期', '打卡记录', '打卡次数']];
    report.employees.forEach(employee => employee.days.forEach(entry => {
      rows.push([employee.id, employee.name, employee.department, `${report.year}-${String(report.month).padStart(2, '0')}-${String(entry.day).padStart(2, '0')}`, `周${WEEKDAYS[getWeekday(entry.day)]}`, entry.times.join(' / '), entry.times.length]);
    }));
    const csv = '\ufeff' + rows.map(row => row.map(value => `"${String(value).replace(/"/g, '""')}"`).join(',')).join('\r\n');
    const link = document.createElement('a');
    link.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    link.download = `${report.fileName.replace(/\.xlsx?$/i, '')}_考勤记录明细.csv`;
    link.click();
    URL.revokeObjectURL(link.href);
    showToast('明细 CSV 已生成');
  }

  els.fileInput.addEventListener('change', event => loadFile(event.target.files[0]));
  ['dragenter', 'dragover'].forEach(type => els.dropZone.addEventListener(type, event => { event.preventDefault(); els.dropZone.classList.add('is-dragging'); }));
  ['dragleave', 'drop'].forEach(type => els.dropZone.addEventListener(type, event => { event.preventDefault(); els.dropZone.classList.remove('is-dragging'); }));
  els.dropZone.addEventListener('drop', event => loadFile(event.dataTransfer.files[0]));
  els.backButton.addEventListener('click', () => { els.dashboard.hidden = true; els.uploadView.hidden = false; state.report = null; });
  els.exportButton.addEventListener('click', exportCsv);
  els.searchInput.addEventListener('input', event => {
    state.query = event.target.value.trim();
    renderMatrix();
    renderEmployee();
  });
  $$('.filter-chip').forEach(button => button.addEventListener('click', () => {
    state.filter = button.dataset.filter;
    $$('.filter-chip').forEach(item => item.classList.toggle('is-active', item === button));
    renderMatrix();
    renderEmployee();
  }));
})();
